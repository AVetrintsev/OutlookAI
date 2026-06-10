<#
.SYNOPSIS
    Builds a Release publish of OutlookAI, generates version.json, packages
    the deploy ZIP, computes the SHA256 sidecar. Output is suitable for
    `gh release create`.

.PARAMETER Tag
    Semver release tag, e.g. v2.1.0. Required.

.PARAMETER OutDir
    Directory where the final .zip and .zip.sha256 should land. Defaults to
    .\out next to the repo root.

.PARAMETER CertThumbprint
    Optional override for <ManifestCertificateThumbprint> in the csproj.
    Used by the CI release workflow to point MSBuild at an ephemeral
    self-signed cert generated on the runner (the dev box's cert is not
    available there, and VSTO's Office targets require manifest signing).
    When empty, MSBuild uses whatever the csproj specifies — that's the
    correct default for local dev builds.

.EXAMPLE
    .\Deploy\Make-ReleaseZip.ps1 -Tag v2.1.0
#>
param(
    [Parameter(Mandatory=$true)][string]$Tag,
    [string]$OutDir = "out",
    [string]$CertThumbprint = "",
    [switch]$KeepIntermediate
)

$ErrorActionPreference = "Stop"

if ($Tag -notmatch '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
    throw "Tag '$Tag' is not a valid semver tag (e.g. v2.1.0 or v2.1.0-beta.1)."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

function Get-VstoTargetsForMSBuild {
    param([string]$MSBuildPath)

    if ([string]::IsNullOrWhiteSpace($MSBuildPath) -or -not (Test-Path -LiteralPath $MSBuildPath)) {
        return $null
    }

    $directory = Split-Path -Parent ([System.IO.Path]::GetFullPath($MSBuildPath))
    while (-not [string]::IsNullOrWhiteSpace($directory) -and (Split-Path -Leaf $directory) -ne "MSBuild") {
        $parent = Split-Path -Parent $directory
        if ($parent -eq $directory) { break }
        $directory = $parent
    }

    if ([string]::IsNullOrWhiteSpace($directory) -or (Split-Path -Leaf $directory) -ne "MSBuild") {
        return $null
    }

    $vsRoot = Split-Path -Parent $directory
    $targetsRoot = Join-Path $vsRoot "MSBuild\Microsoft\VisualStudio"
    if (-not (Test-Path -LiteralPath $targetsRoot)) {
        return $null
    }

    return Get-ChildItem -Path (Join-Path $targetsRoot "v*\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets") -ErrorAction SilentlyContinue |
        Sort-Object -Property FullName -Descending |
        Select-Object -ExpandProperty FullName -First 1
}

function Add-MSBuildCandidate {
    param(
        [System.Collections.Generic.List[object]]$Candidates,
        [string]$Path,
        [string]$Source
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    foreach ($candidate in $Candidates) {
        if ([string]::Equals($candidate.Path, $fullPath, [StringComparison]::OrdinalIgnoreCase)) {
            return
        }
    }

    $Candidates.Add([pscustomobject]@{
        Path = $fullPath
        Source = $Source
        VstoTargets = Get-VstoTargetsForMSBuild $fullPath
    }) | Out-Null
}

function Get-NuGetExe {
    param([string]$RepoRoot)

    $nuget = Get-Command nuget.exe -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Source -First 1
    if ($nuget) {
        return $nuget
    }

    $toolsDir = Join-Path $RepoRoot "out\tools"
    if (-not (Test-Path -LiteralPath $toolsDir)) {
        New-Item -ItemType Directory -Path $toolsDir | Out-Null
    }

    $nuget = Join-Path $toolsDir "nuget.exe"
    if (-not (Test-Path -LiteralPath $nuget)) {
        Write-Host "Downloading nuget.exe..." -ForegroundColor Gray
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nuget -UseBasicParsing
    }

    return $nuget
}

function Restore-NuGetPackages {
    param(
        [string]$RepoRoot,
        [string]$NuGetExe
    )

    $packagesConfig = Join-Path $RepoRoot "VSTO2\OutlookAI\packages.config"
    $packagesDir = Join-Path $RepoRoot "VSTO2\packages"
    if (-not (Test-Path -LiteralPath $packagesConfig)) {
        throw "Missing packages.config: $packagesConfig"
    }

    Write-Host "Restoring NuGet packages..." -ForegroundColor Gray
    & $NuGetExe restore $packagesConfig -PackagesDirectory $packagesDir -NonInteractive
    if ($LASTEXITCODE -ne 0) {
        throw "NuGet restore failed with exit code $LASTEXITCODE."
    }
}

function Test-CodeSigningCertificate {
    param([string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return $false
    }

    return Test-Path -LiteralPath ("Cert:\CurrentUser\My\" + $Thumbprint)
}

function New-TemporaryManifestCertificate {
    if (-not (Get-Command New-SelfSignedCertificate -ErrorAction SilentlyContinue)) {
        throw "No manifest certificate thumbprint was supplied, and New-SelfSignedCertificate is unavailable. Supply -CertThumbprint or run on Windows with the PKI module."
    }

    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject "CN=OutlookAI Local Build" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -NotAfter (Get-Date).AddDays(7)

    if (-not $cert -or [string]::IsNullOrWhiteSpace($cert.Thumbprint)) {
        throw "Failed to create a temporary manifest signing certificate."
    }

    return $cert.Thumbprint
}

$temporaryCertThumbprint = $null

Push-Location $repoRoot
try {
    $staging = Join-Path $repoRoot "out\staging-$Tag"
    if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
    New-Item -ItemType Directory -Path $staging | Out-Null
    $nuget = Get-NuGetExe -RepoRoot $repoRoot
    Restore-NuGetPackages -RepoRoot $repoRoot -NuGetExe $nuget

    # Locate MSBuild.exe with Office/VSTO targets. A plain MSBuild install is
    # not enough for VSTO projects because OutlookAI.csproj imports
    # Microsoft.VisualStudio.Tools.Office.targets.
    $candidates = New-Object "System.Collections.Generic.List[object]"

    Get-Command MSBuild.exe -All -ErrorAction SilentlyContinue |
        ForEach-Object { Add-MSBuildCandidate -Candidates $candidates -Path $_.Source -Source "PATH" }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswhere) {
        & $vswhere -all -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null |
            ForEach-Object { Add-MSBuildCandidate -Candidates $candidates -Path $_ -Source "vswhere" }
    }

    foreach ($fallback in @(
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    )) {
        Add-MSBuildCandidate -Candidates $candidates -Path $fallback -Source "fallback"
    }

    $selected = $candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_.VstoTargets) } | Select-Object -First 1
    if (-not $selected) {
        $details = if ($candidates.Count -gt 0) {
            ($candidates | ForEach-Object { "  - $($_.Path) ($($_.Source)): missing Office/VSTO targets" }) -join [Environment]::NewLine
        } else {
            "  - no MSBuild.exe candidates found"
        }

        throw @"
MSBuild with Office/VSTO build targets was not found.

Install Visual Studio or Visual Studio Build Tools with the Office/SharePoint development workload, then rerun this script.

Detected MSBuild candidates:
$details
"@
    }

    $msbuild = $selected.Path
    $vstoToolsPath = Split-Path -Parent (Split-Path -Parent $selected.VstoTargets)
    Write-Host "Using MSBuild: $msbuild" -ForegroundColor Gray
    Write-Host "Using VSTO targets: $($selected.VstoTargets)" -ForegroundColor Gray

    $manifestCertificateThumbprint = $CertThumbprint
    if ([string]::IsNullOrWhiteSpace($manifestCertificateThumbprint)) {
        Write-Host "No manifest certificate thumbprint supplied; creating a temporary local signing certificate." -ForegroundColor Gray
        $temporaryCertThumbprint = New-TemporaryManifestCertificate
        $manifestCertificateThumbprint = $temporaryCertThumbprint
    } elseif (-not (Test-CodeSigningCertificate $manifestCertificateThumbprint)) {
        throw "Manifest signing certificate '$manifestCertificateThumbprint' was not found in Cert:\CurrentUser\My."
    }

    # Publish only the main project, not the whole solution, so the staging dir
    # does not pick up OutlookAI.Tests artifacts (Moq, xunit, Castle.Core,
    # Microsoft.CodeCoverage.*, etc.). The csproj's Platform is "AnyCPU"
    # (no space) even though the sln uses "Any CPU".
    #
    # The csproj has a hardcoded <ManifestCertificateThumbprint> pointing at
    # the dev box's local cert store. On the dev box this Just Works; on a
    # CI runner the cert isn't there, so MSBuild dies with MSB3323. VSTO's
    # Office targets force SignManifests=true (can't bypass), so the
    # workaround is to OVERRIDE the thumbprint with an ephemeral runner-local
    # cert that the workflow generates per-run. The resulting signature is
    # never re-checked at install time: Install-OutlookAI.ps1 registers the
    # add-in with the `|vstolocal` trust modifier, which makes Outlook trust
    # the manifest by location rather than by publisher.
    $msbuildArgs = @(
        "VSTO2\OutlookAI\OutlookAI.csproj"
        "/target:Publish"
        "/p:Configuration=Release"
        "/p:Platform=AnyCPU"
        "/p:PublishDir=$staging\"
        "/p:VSToolsPath=$vstoToolsPath"
        "/v:minimal"
        "/nologo"
    )
    $msbuildArgs += "/p:ManifestCertificateThumbprint=$manifestCertificateThumbprint"
    & $msbuild @msbuildArgs
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild publish failed with exit code $LASTEXITCODE."
    }

    # Copy install assets next to OutlookAI.vsto / Application Files\
    foreach ($f in @(
        "Install-OutlookAI.ps1",
        "Uninstall-OutlookAI.ps1",
        "Fetch-WebView2Bootstrapper.ps1",
        "MicrosoftEdgeWebView2Setup.exe",
        "README.txt"
    )) {
        Copy-Item -LiteralPath (Join-Path "Deploy" $f) -Destination (Join-Path $staging $f) -Force
    }

    # Convenience copy of the actual DLL at the staging root for hash-verify
    Copy-Item -LiteralPath "VSTO2\OutlookAI\bin\Release\OutlookAI.dll" -Destination (Join-Path $staging "OutlookAI.dll") -Force

    # version.json
    $commit = (git rev-parse --short HEAD).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($commit)) {
        throw "git rev-parse failed (exit=$LASTEXITCODE, output='$commit'). Run from a git working tree."
    }
    $buildDate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $versionJson = [ordered]@{
        tag = $Tag
        commit = $commit
        build_date = $buildDate
        repo = "kirklandsig/OutlookAI"
    } | ConvertTo-Json
    # Write UTF-8 *without* BOM. The in-app updater and any downstream JSON
    # reader should not see a leading 0xEF 0xBB 0xBF that some parsers reject.
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $staging "version.json"), $versionJson, $utf8NoBom)

    # Defensive: fail loudly if something we expect downstream isn't there,
    # rather than shipping a broken ZIP.
    foreach ($required in @("OutlookAI.vsto","Application Files","Install-OutlookAI.ps1","MicrosoftEdgeWebView2Setup.exe","version.json")) {
        if (-not (Test-Path -LiteralPath (Join-Path $staging $required))) {
            throw "Staging is missing required artifact: $required"
        }
    }

    # Zip
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
    $zipName = "OutlookAI-$Tag-RDS-Deploy.zip"
    $zipPath = Join-Path $OutDir $zipName
    if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal

    # SHA256 sidecar (single line of lowercase hex, no trailing newline beyond what Set-Content adds)
    $sha = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -LiteralPath ($zipPath + ".sha256") -Encoding ASCII -Value $sha -NoNewline

    Write-Host "Built $zipPath" -ForegroundColor Green
    Write-Host "SHA256 $sha" -ForegroundColor Green

    if (-not $KeepIntermediate) {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($temporaryCertThumbprint)) {
        Remove-Item -LiteralPath ("Cert:\CurrentUser\My\" + $temporaryCertThumbprint) -Force -ErrorAction SilentlyContinue
    }
    Pop-Location
}
