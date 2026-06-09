#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs OutlookAI v3 (LiteLLM connector) for all users on RDS / Terminal Server.

.DESCRIPTION
    v3 install:
      - Hardcoded install path: C:\Program Files\OutlookAI
      - Backs up any v1 config to C:\ProgramData\OutlookAI\Backups
      - Writes a fresh v3 config when one is missing
      - Renames any per-user v1 %APPDATA%\OutlookAI\config.xml to a backup
        so per-user files don't override the new server-authoritative
        LiteLLM endpoint/model settings

    Run as Administrator.

.PARAMETER SourcePath
    Path to the published OutlookAI files (containing OutlookAI.vsto and the
    Application Files folder). Defaults to C:\OutlookAI.

.EXAMPLE
    .\Install-OutlookAI.ps1
#>

param(
    [string]$SourcePath = "C:\OutlookAI",
    [string]$LiteLlmBaseUrl = "https://litellm.example.com/v1",
    [string]$LiteLlmModel = "gpt-4.1-mini",
    [AllowEmptyString()][string]$LiteLlmVoiceModel = "",
    [double]$Temperature = 0.2,
    [int]$MaxTokens = 4096,
    [int]$MaxBulkExportRows = 2000
)

$ErrorActionPreference = "Stop"
$InstallPath        = "C:\Program Files\OutlookAI"
$ProgramDataPath    = "C:\ProgramData\OutlookAI"
$BackupRoot         = Join-Path $ProgramDataPath "Backups"
$ConfigFilePath     = Join-Path $InstallPath "config.xml"
$ProgramFilesX86    = ${env:ProgramFiles(x86)}
$ConfigMirrorFilePath = if ([string]::IsNullOrWhiteSpace($ProgramFilesX86)) { "" } else { Join-Path $ProgramFilesX86 "OutlookAI\config.xml" }
$Timestamp          = Get-Date -Format "yyyyMMdd-HHmmss"

function Normalize-OptionalValue {
    param([AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    $trimmed = $Value.Trim()
    if ($trimmed -in @("null", "none", "no", "off", "-")) {
        return ""
    }

    return $trimmed
}

$LiteLlmVoiceModel = Normalize-OptionalValue $LiteLlmVoiceModel

# Cleans every known OutlookAI registration for one Windows user. Designed
# to be called once per user hive (offline-loaded for non-logged-in users,
# directly under HKEY_USERS\<sid> for logged-on ones). Idempotent.
#
# Why all of these?  VSTO's ClickOnceAddInDeploymentManager throws
# AddInAlreadyInstalledException at Outlook startup if any of the following
# point at an older or differently-located manifest than HKLM does:
#   - HKCU\Software\Microsoft\VSTA\Solutions\<guid>          (ClickOnce subscription)
#   - HKCU\Software\Microsoft\VSTO\SolutionMetadata\<guid>   (SolutionId â†’ addin metadata)
#   - HKCU\Software\Microsoft\VSTO\SolutionMetadata          (URL â†’ SolutionId map values)
#   - HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid> (trust list)
#   - HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\<guid> (Add/Remove Programs)
#   - HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI
#   - HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI
#   - HKCU\Software\Microsoft\Office\Outlook\AddinsData\OutlookAI
#   - HKCU\Software\Microsoft\Office\16.0\Outlook\AddinsData\OutlookAI
#   - %LOCALAPPDATA%\Apps\2.0\...\outl..vsto_*               (ClickOnce app cache)
# Even with `|vstolocal` in HKLM, Outlook prefers any matching HKCU entry,
# and any of the VSTA/VSTO state above can re-trigger the ClickOnce path.
function Clean-StaleOutlookAIRegistrations {
    param(
        [Parameter(Mandatory=$true)] [string] $UserRegistryRoot,
        [string] $ProfilePath,
        [string] $VstoInstaller
    )

    function _RemoveKey($path) {
        if (Test-Path $path) {
            try { Remove-Item -Path $path -Recurse -Force -ErrorAction Stop; return $true } catch { return $false }
        }
        return $false
    }

    # 1) Add/Remove Programs subscription (Uninstall key). Run its own
    #    UninstallString first when VSTOInstaller is available; then drop
    #    the registry entry whether or not the helper succeeded.
    $uninstallRoot = "$UserRegistryRoot\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    if (Test-Path $uninstallRoot) {
        Get-ChildItem -Path $uninstallRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if (-not $entry) { return }
            if ($entry.DisplayName -notlike "*OutlookAI*") { return }

            if ($VstoInstaller -and $entry.UninstallString -and $entry.UninstallString -match 'file:[^\s"]+OutlookAI\.vsto') {
                $manifest = $Matches[0]
                Write-Host ("  Uninstalling ClickOnce subscription: {0}" -f $manifest) -ForegroundColor Gray
                & $VstoInstaller /Uninstall $manifest /s 2>$null | Out-Null
            }
            try { Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop } catch { }
        }
    }

    # 2) VSTA Solutions subscription (the one VerifySolutionCodebaseIsUnchanged
    #    actually checks). Matched by ProductName/Url containing OutlookAI.
    $vstaRoot = "$UserRegistryRoot\Software\Microsoft\VSTA\Solutions"
    if (Test-Path $vstaRoot) {
        Get-ChildItem -Path $vstaRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and ($entry.ProductName -like "*OutlookAI*" -or $entry.Url -like "*OutlookAI*")) {
                Write-Host ("  Removing VSTA subscription: {0}" -f $entry.Url) -ForegroundColor Gray
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 3) VSTO SolutionMetadata: subkeys (SolutionId â†’ addin name) and
    #    values directly on the parent (URL â†’ SolutionId map).
    $metaRoot = "$UserRegistryRoot\Software\Microsoft\VSTO\SolutionMetadata"
    if (Test-Path $metaRoot) {
        Get-ChildItem -Path $metaRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and ($entry.addInName -like "*OutlookAI*" -or $entry.friendlyName -like "*OutlookAI*")) {
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        $metaKey = Get-Item -Path $metaRoot -ErrorAction SilentlyContinue
        if ($metaKey) {
            foreach ($valName in $metaKey.GetValueNames()) {
                if ($valName -like "*OutlookAI*") {
                    Remove-ItemProperty -Path $metaRoot -Name $valName -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    # 4) VSTO Security Inclusion list (trust). Stale entries point at moved
    #    build paths and confuse the manifest resolver.
    $inclusionRoot = "$UserRegistryRoot\Software\Microsoft\VSTO\Security\Inclusion"
    if (Test-Path $inclusionRoot) {
        Get-ChildItem -Path $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and $entry.Url -like "*OutlookAI*") {
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 5) HKCU Outlook Addin entries (all known variants).
    foreach ($leaf in @(
        "Software\Microsoft\Office\Outlook\Addins\OutlookAI",
        "Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI",
        "Software\Microsoft\Office\Outlook\AddinsData\OutlookAI",
        "Software\Microsoft\Office\16.0\Outlook\AddinsData\OutlookAI"
    )) {
        _RemoveKey "$UserRegistryRoot\$leaf" | Out-Null
    }

    # 6) Per-user Outlook cache state. These are harmless if left, but Outlook
    #    occasionally surfaces stale ribbon validation results when an addin
    #    is reinstalled with different ribbon XML.
    $loadTimes = "$UserRegistryRoot\Software\Microsoft\Office\16.0\Outlook\AddInLoadTimes"
    if (Test-Path $loadTimes) {
        Remove-ItemProperty -Path $loadTimes -Name "OutlookAI" -Force -ErrorAction SilentlyContinue
    }
    $uiCache = "$UserRegistryRoot\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
    if (Test-Path $uiCache) {
        $uiKey = Get-Item -Path $uiCache -ErrorAction SilentlyContinue
        if ($uiKey) {
            foreach ($valName in $uiKey.GetValueNames()) {
                if ($valName -like "OutlookAI*") {
                    Remove-ItemProperty -Path $uiCache -Name $valName -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    # 7) ClickOnce app cache directories under this profile.
    if ($ProfilePath) {
        $appsRoot = Join-Path $ProfilePath "AppData\Local\Apps\2.0"
        if (Test-Path $appsRoot) {
            Get-ChildItem -Path $appsRoot -Recurse -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "outl*.vsto*" -or $_.Name -like "OutlookAI*" } |
                ForEach-Object {
                    try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
                }
        }
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI v3 Installer (LiteLLM)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Admin check ---------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $SourcePath)) {
    Write-Host "ERROR: Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

$vstoFile    = Join-Path $SourcePath "OutlookAI.vsto"
$appFilesDir = Join-Path $SourcePath "Application Files"

if (!(Test-Path $vstoFile)) {
    Write-Host "ERROR: OutlookAI.vsto not found in $SourcePath" -ForegroundColor Red
    exit 1
}
if (!(Test-Path $appFilesDir)) {
    Write-Host "ERROR: 'Application Files' folder not found in $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Source : $SourcePath"     -ForegroundColor Gray
Write-Host "Target : $InstallPath"    -ForegroundColor Gray
Write-Host "Config : $ConfigFilePath" -ForegroundColor Gray
Write-Host ""

# --- 0. Clean up any stale OutlookAI registrations ----------------------
# A previous OutlookAI install (Visual Studio publish, ClickOnce setup.exe,
# F5/debug deploy, or v1 installer) may have left behind:
#   - HKCU Outlook add-in entry pointing at a stale build path
#   - HKCU\...\Uninstall ClickOnce subscription pointing at an old manifest
#   - %LOCALAPPDATA%\Apps\2.0 cached ClickOnce manifest
# Even with |vstolocal in our HKLM entry, Outlook prefers HKCU and the
# ClickOnce runtime throws AddInAlreadyInstalledException when the cached
# subscription's manifest URL no longer matches the current install path.
#
# This block cleans up every Windows user profile we can see on this
# machine, not just the admin running the installer. ClickOnce state is
# per-user, so cleaning only the admin profile leaves other users broken.
Write-Host "[0/10] Removing stale OutlookAI registrations for all users..." -ForegroundColor Yellow

# Stop Outlook so the cache files aren't locked.
Get-Process OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  Stopping running Outlook process..." -ForegroundColor Gray
    try { $_ | Stop-Process -Force -ErrorAction Stop } catch { }
}

$vstoInstallerCandidates = @(
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
)
$vstoInstaller = $vstoInstallerCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

# Iterate every loadable user hive. Load offline hives, clean, unload.
$userProfiles = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("Public", "Default", "Default User", "All Users") }

foreach ($profile in $userProfiles) {
    $ntuser = Join-Path $profile.FullName "NTUSER.DAT"
    if (!(Test-Path $ntuser)) { continue }

    $hiveName = "OutlookAICleanup_" + $profile.Name
    $loaded = $false
    try {
        # Try to load the offline hive. Fails if the user is logged in (their
        # hive is already mounted under HKEY_USERS\<sid>); in that case we
        # walk HKEY_USERS directly below. Also fails for locked hives.
        $prevEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        & reg.exe load "HKU\$hiveName" "$ntuser" 2>$null | Out-Null
        $ErrorActionPreference = $prevEAP
        if ($LASTEXITCODE -eq 0) { $loaded = $true }
    } catch {
        # Swallow: locked hives, access denied, etc. are expected for
        # currently-logged-in users and service accounts.
    }

    if ($loaded) {
        $userRoot = "Registry::HKEY_USERS\$hiveName"
        try {
            Clean-StaleOutlookAIRegistrations -UserRegistryRoot $userRoot -ProfilePath $profile.FullName -VstoInstaller $vstoInstaller
        } catch {
            Write-Host ("  Skipped offline-hive cleanup for {0}: {1}" -f $profile.Name, $_.Exception.Message) -ForegroundColor DarkGray
        } finally {
            $prevEAP = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            & reg.exe unload "HKU\$hiveName" 2>$null | Out-Null
            $ErrorActionPreference = $prevEAP
        }
    }
}

# Also clean every currently-loaded HKEY_USERS hive (the admin running this
# script and any other logged-on session). Each iteration is wrapped so a
# single access-denied SID (service accounts we can't read) doesn't kill
# the whole cleanup pass.
Get-ChildItem "Registry::HKEY_USERS" -ErrorAction SilentlyContinue | Where-Object {
    $_.PSChildName -match '^S-1-5-21-' -and $_.PSChildName -notmatch '_Classes$'
} | ForEach-Object {
    $sid = $_.PSChildName
    $userRoot = "Registry::HKEY_USERS\$sid"
    try {
        $profilePath = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
        Clean-StaleOutlookAIRegistrations -UserRegistryRoot $userRoot -ProfilePath $profilePath -VstoInstaller $vstoInstaller
    } catch {
        Write-Host ("  Skipped live-hive cleanup for {0}: {1}" -f $sid, $_.Exception.Message) -ForegroundColor DarkGray
    }
}

Write-Host "  Done." -ForegroundColor Green

# --- 1. External config backup BEFORE we touch the install dir -----------
Write-Host "[1/10] Backing up any existing v1 config..." -ForegroundColor Yellow
if (!(Test-Path $BackupRoot)) {
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
}
if (Test-Path $ConfigFilePath) {
    $backupTarget = Join-Path $BackupRoot ("config.xml.v1.backup." + $Timestamp)
    Copy-Item -Path $ConfigFilePath -Destination $backupTarget -Force
    Write-Host "  Backed up to $backupTarget" -ForegroundColor Gray
} else {
    Write-Host "  No existing config.xml to back up." -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

# --- 2. Replace install directory ----------------------------------------
Write-Host "[2/10] Preparing install directory..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
}
New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
Write-Host "  Done." -ForegroundColor Green

# --- 3. Copy publish payload ---------------------------------------------
Write-Host "[3/10] Copying files..." -ForegroundColor Yellow
Copy-Item -Path $vstoFile -Destination $InstallPath -Force
Copy-Item -Path $appFilesDir -Destination $InstallPath -Recurse -Force

# Mirror .deploy artifacts as plain DLLs at the root (mirrors v1 layout).
$latestVersionDir = Get-ChildItem -Path (Join-Path $InstallPath "Application Files") -Directory |
    Sort-Object Name -Descending | Select-Object -First 1
if ($latestVersionDir) {
    Get-ChildItem -Path $latestVersionDir.FullName -Filter "*.deploy" | ForEach-Object {
        $newName = $_.Name -replace '\.deploy$', ''
        Copy-Item -Path $_.FullName -Destination (Join-Path $InstallPath $newName) -Force
    }
    $manifestFile = Join-Path $latestVersionDir.FullName "OutlookAI.dll.manifest"
    if (Test-Path $manifestFile) {
        Copy-Item -Path $manifestFile -Destination $InstallPath -Force
    }
}
Write-Host "  Done." -ForegroundColor Green

# --- 4. Unblock files ----------------------------------------------------
Write-Host "[4/10] Unblocking files..." -ForegroundColor Yellow
Get-ChildItem -Path $InstallPath -Recurse | Unblock-File -ErrorAction SilentlyContinue
Write-Host "  Done." -ForegroundColor Green

# --- 5. Write v3 config --------------------------------------------------
Write-Host "[5/10] Writing v3 config.xml..." -ForegroundColor Yellow

# Carry over only the AdminPassword from any previous file. LiteLLM
# connection defaults are server-authoritative and are supplied by this
# installer invocation. User API keys are intentionally not written here.
$preservedAdminPassword = "admin"
$latestBackup = Get-ChildItem -Path $BackupRoot -Filter "config.xml.v1.backup.*" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestBackup) {
    try {
        [xml]$oldXml = Get-Content -Path $latestBackup.FullName -Raw
        if ($oldXml.Config -and $oldXml.Config.AdminPassword) {
            $candidate = [string]$oldXml.Config.AdminPassword
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $preservedAdminPassword = $candidate
                Write-Host "  Preserved AdminPassword from previous config." -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  Could not parse previous config; using default AdminPassword." -ForegroundColor Yellow
    }
}

$xmlAdminPassword = [System.Security.SecurityElement]::Escape($preservedAdminPassword)
$xmlLiteLlmBaseUrl = [System.Security.SecurityElement]::Escape($LiteLlmBaseUrl)
$xmlLiteLlmModel = [System.Security.SecurityElement]::Escape($LiteLlmModel)
$xmlLiteLlmVoiceModel = [System.Security.SecurityElement]::Escape($LiteLlmVoiceModel)

$v3Config = @"
<Config>
  <AdminPassword>$xmlAdminPassword</AdminPassword>
  <LiteLlmBaseUrl>$xmlLiteLlmBaseUrl</LiteLlmBaseUrl>
  <Model>$xmlLiteLlmModel</Model>
  <VoiceModel>$xmlLiteLlmVoiceModel</VoiceModel>
  <Temperature>$Temperature</Temperature>
  <MaxTokens>$MaxTokens</MaxTokens>
  <MaxBulkExportRows>$MaxBulkExportRows</MaxBulkExportRows>
</Config>
"@

Set-Content -Path $ConfigFilePath -Value $v3Config -Encoding UTF8
Write-Host "  Wrote $ConfigFilePath" -ForegroundColor Gray

# 32-bit Outlook runs the add-in in a 32-bit .NET process where
# Environment.SpecialFolder.ProgramFiles resolves to Program Files (x86).
# Keep a tiny config mirror there so that bitness cannot make the add-in
# fall back to compiled LiteLLM defaults.
if (-not [string]::IsNullOrWhiteSpace($ConfigMirrorFilePath) -and
    -not [string]::Equals($ConfigMirrorFilePath, $ConfigFilePath, [System.StringComparison]::OrdinalIgnoreCase)) {
    $mirrorDir = Split-Path -Parent $ConfigMirrorFilePath
    if (!(Test-Path $mirrorDir)) {
        New-Item -ItemType Directory -Path $mirrorDir -Force | Out-Null
    }
    Set-Content -Path $ConfigMirrorFilePath -Value $v3Config -Encoding UTF8
    Write-Host "  Wrote $ConfigMirrorFilePath" -ForegroundColor Gray
}
# v2.1+ release packages ship a version.json alongside Install-OutlookAI.ps1.
# Copy it into the install dir so the in-app updater knows what is installed.
$stagedVersionJson = Join-Path $SourcePath "version.json"
if (Test-Path $stagedVersionJson) {
    $installedVersionJson = Join-Path $InstallPath "version.json"
    Copy-Item -LiteralPath $stagedVersionJson -Destination $installedVersionJson -Force
    Write-Host "  Wrote $installedVersionJson" -ForegroundColor Gray
} else {
    # Backwards-compatible: older deploy ZIPs do not have version.json. The
    # in-app updater shows "Current: (dev build)" in this case and still
    # allows updates.
    Write-Host "  (no version.json in source path; skipping)" -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

# --- 6. Provision ProgramData folder -------------------------------------
Write-Host "[6/10] Provisioning ProgramData folder..." -ForegroundColor Yellow
if (!(Test-Path $ProgramDataPath)) {
    New-Item -ItemType Directory -Path $ProgramDataPath -Force | Out-Null
}
& icacls.exe $ProgramDataPath /grant "Authenticated Users:(OI)(CI)RX" /T | Out-Null
Write-Host "  Granted Authenticated Users: Read/Execute on $ProgramDataPath" -ForegroundColor Gray
Write-Host "  Done." -ForegroundColor Green

# --- 7. Per-user v1 AppData cleanup --------------------------------------
Write-Host "[7/10] Renaming per-user v1 AppData configs..." -ForegroundColor Yellow
$userProfiles = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("Public", "Default", "Default User", "All Users") }
$renamed = 0
foreach ($profile in $userProfiles) {
    $userConfig = Join-Path $profile.FullName "AppData\Roaming\OutlookAI\config.xml"
    if (Test-Path $userConfig) {
        try {
            [xml]$xml = Get-Content -Path $userConfig -Raw
            if ($xml.Config -and -not $xml.Config.LiteLlmBaseUrl) {
                $renamed++
                $renamedTarget = "$userConfig.v1.backup.$Timestamp"
                Move-Item -Path $userConfig -Destination $renamedTarget -Force
                Write-Host "  Renamed $userConfig -> $renamedTarget" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Skipped unreadable per-user config: $userConfig" -ForegroundColor Yellow
        }
    }
}
Write-Host ("  Done ({0} per-user v1 configs renamed)." -f $renamed) -ForegroundColor Green

# --- 8. Microsoft Edge WebView2 Runtime --------------------------------------
# Phase 2 chat surface uses WebView2 (Evergreen). On most modern Windows the
# runtime ships with the OS / is auto-installed by Edge, but on stripped-down
# server images it can be missing. Detect first; only run the bootstrapper
# (vendored alongside the installer as MicrosoftEdgeWebView2Setup.exe) if
# we don't see a version registered. If the bootstrapper isn't shipped with
# this install, the add-in's Chat tab falls back to a friendly "install
# WebView2 runtime" panel - the rest of the add-in (Actions tab, voice,
# Variants tab) keeps working.
Write-Host "[8/10] Ensuring Microsoft Edge WebView2 Runtime present..." -ForegroundColor Yellow
$wv2KeyA = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$wv2KeyB = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$wv2Installed = $false
foreach ($k in @($wv2KeyA, $wv2KeyB)) {
    if (Test-Path $k) {
        $pv = (Get-ItemProperty $k -ErrorAction SilentlyContinue).pv
        if (-not [string]::IsNullOrWhiteSpace($pv)) {
            $wv2Installed = $true
            Write-Host "  WebView2 Runtime detected (version $pv)." -ForegroundColor Gray
            break
        }
    }
}
if (-not $wv2Installed) {
    $bootstrap = Join-Path $SourcePath "MicrosoftEdgeWebView2Setup.exe"
    if (Test-Path $bootstrap) {
        Write-Host "  Installing WebView2 Runtime (silent)..." -ForegroundColor Gray
        try {
            Start-Process -FilePath $bootstrap -ArgumentList "/silent","/install" -Wait
            Write-Host "  WebView2 Runtime installed." -ForegroundColor Gray
        } catch {
            Write-Host "  WARN: WebView2 bootstrapper failed: $_" -ForegroundColor Yellow
            Write-Host "  Chat tab will show the runtime-missing fallback panel." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  WARN: MicrosoftEdgeWebView2Setup.exe missing in $SourcePath." -ForegroundColor Yellow
        Write-Host "  Chat tab will show the runtime-missing fallback panel until WebView2 is installed manually." -ForegroundColor Yellow
        Write-Host "  Download: https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Yellow
    }
}
Write-Host "  Done." -ForegroundColor Green

# --- 9. VSTO trust + Inclusion List + add-in registration ----------------
Write-Host "[9/10] Configuring VSTO trust & registering add-in..." -ForegroundColor Yellow

$trustPath   = "HKLM:\SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
$trustPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
foreach ($path in @($trustPath, $trustPath32)) {
    if (!(Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "MyComputer"    -Value "Enabled"
    Set-ItemProperty -Path $path -Name "LocalIntranet" -Value "Enabled"
}

$inclusionPath   = "HKLM:\SOFTWARE\Microsoft\VSTO\Security\Inclusion"
$inclusionPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO\Security\Inclusion"
foreach ($base in @($inclusionPath, $inclusionPath32)) {
    if (!(Test-Path $base)) { New-Item -Path $base -Force | Out-Null }
    $guid    = [System.Guid]::NewGuid().ToString("B")
    $entry   = Join-Path $base $guid
    New-Item -Path $entry -Force | Out-Null
    Set-ItemProperty -Path $entry -Name "Url"       -Value "file:///C:/Program Files/OutlookAI/OutlookAI.vsto"
    Set-ItemProperty -Path $entry -Name "PublicKey" -Value ""
}

$manifestPath  = "file:///C:/Program Files/OutlookAI/OutlookAI.vsto|vstolocal"
$addinPath     = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI"
$addinPath32   = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
foreach ($path in @($addinPath, $addinPath32)) {
    if (Test-Path $path) { Remove-Item -Path $path -Force }
    New-Item -Path $path -Force | Out-Null
    Set-ItemProperty -Path $path -Name "Description"  -Value "AI Writing Assistant for Outlook"
    Set-ItemProperty -Path $path -Name "FriendlyName" -Value "OutlookAI"
    Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $path -Name "Manifest"     -Value $manifestPath
}

# Prevent Outlook's resiliency detector from auto-disabling OutlookAI
# when a single long operation (e.g. outlook_count_messages on a
# 200-folder mailbox, or a long-running search) takes too much UI-thread
# time. Outlook normally watches add-in event handlers and ribbon click
# response times, and if cumulative latency crosses a threshold it pops
# the "this add-in is slow" prompt + can move the add-in to disabled.
# Setting a DWORD value with the add-in's name under
# Outlook\Resiliency\DoNotDisableAddinList = 1 tells Outlook this add-in
# is sanctioned and must remain enabled regardless of measured slowness.
# Reference:
#   https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-prevent-the-disabling-of-an-office-solution
$resiliencyPath = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"
try {
    if (!(Test-Path $resiliencyPath)) {
        New-Item -Path $resiliencyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $resiliencyPath -Name "OutlookAI" -Value 1 -Type DWord
    Write-Host "  Registered OutlookAI in DoNotDisableAddinList (auto-disable protected)." -ForegroundColor Gray
} catch {
    Write-Host ("  Could not set DoNotDisableAddinList: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
}

Write-Host "  Done." -ForegroundColor Green

# --- 10. Default user profile auto-load ----------------------------------
Write-Host "[10/10] Configuring auto-load for new user profiles..." -ForegroundColor Yellow
$defaultUserPath = "C:\Users\Default\NTUSER.DAT"
$tempKey         = "HKU\DefaultUser"
try {
    reg load $tempKey $defaultUserPath 2>$null
    $defaultAddinPath = "Registry::$tempKey\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI"
    if (!(Test-Path $defaultAddinPath)) {
        New-Item -Path $defaultAddinPath -Force | Out-Null
    }
    Set-ItemProperty -Path $defaultAddinPath -Name "LoadBehavior" -Value 3 -Type DWord
    reg unload $tempKey 2>$null
    Write-Host "  Configured default user profile." -ForegroundColor Gray
} catch {
    Write-Host "  Note: could not modify default user profile (may need reboot)." -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "LiteLLM connector defaults:" -ForegroundColor Yellow
Write-Host "  Base URL : $LiteLlmBaseUrl" -ForegroundColor Yellow
Write-Host "  Model    : $LiteLlmModel" -ForegroundColor Yellow
Write-Host "  Voice    : $LiteLlmVoiceModel" -ForegroundColor Yellow
Write-Host "  API keys are not installed globally. Each user enters their own" -ForegroundColor Yellow
Write-Host "  LiteLLM API key in OutlookAI Settings." -ForegroundColor Yellow
Write-Host ""
Write-Host "Users will see 'AI Assistant' on the compose-window ribbon." -ForegroundColor White
Write-Host "First action prompts the user to enter their LiteLLM API key." -ForegroundColor White
Write-Host ""
Write-Host "If add-in doesn't auto-load for existing users:" -ForegroundColor Yellow
Write-Host "  - Have user run Enable-OutlookAI-User.ps1, OR"             -ForegroundColor Gray
Write-Host "  - Push it via logon script / GPO."                          -ForegroundColor Gray
Write-Host ""
