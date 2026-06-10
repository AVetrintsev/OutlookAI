<#
.SYNOPSIS
    Builds a single-file OutlookAI installer EXE.

.DESCRIPTION
    This script first builds the existing release ZIP by calling
    Deploy\Make-ReleaseZip.ps1, then wraps that ZIP in a Windows IExpress
    self-extracting executable.

    The generated EXE can be launched by double-clicking. It extracts the
    release payload to a temporary directory, requests elevation through UAC
    when needed, and runs Install-OutlookAI.ps1 with the LiteLLM base
    configuration supplied here.

    User API keys are never embedded in the installer. Each user still enters
    their own LiteLLM API key in OutlookAI Settings.

.PARAMETER Tag
    Semver release tag, e.g. v2.1.0. Required.

.PARAMETER OutDir
    Directory where the release ZIP and installer EXE should land. Defaults to
    .\out next to the repo root.

.PARAMETER LiteLlmBaseUrl
    Base URL for the OpenAI-compatible LiteLLM API.

.PARAMETER LiteLlmModel
    Chat/completions model name to write into the machine config.

.PARAMETER LiteLlmVoiceModel
    Optional audio transcription model name to write into the machine config.
    Use an empty value, "null", "none", "off", or "-" to disable transcription.

.PARAMETER Temperature
    Default chat temperature.

.PARAMETER MaxTokens
    Default maximum completion tokens.

.PARAMETER MaxBulkExportRows
    Default row cap for bulk exports.

.PARAMETER CertThumbprint
    Optional manifest signing certificate thumbprint forwarded to
    Make-ReleaseZip.ps1.

.PARAMETER ExeOnly
    Deletes the intermediate release ZIP and SHA256 sidecars after the EXE is
    created. Use this for local user-facing builds where only one file should
    remain in the output directory.

.EXAMPLE
    .\Deploy\Make-InstallerExe.ps1 `
      -Tag v3.0.0 `
      -LiteLlmBaseUrl "https://litellm.company.example/v1" `
      -LiteLlmModel "company/outlook-chat" `
      -LiteLlmVoiceModel ""
#>

param(
    [Parameter(Mandatory=$true)][string]$Tag,
    [string]$OutDir = "out",
    [string]$LiteLlmBaseUrl = "https://litellm.example.com/v1",
    [string]$LiteLlmModel = "gpt-4.1-mini",
    [AllowEmptyString()][string]$LiteLlmVoiceModel = "",
    [double]$Temperature = 0.2,
    [int]$MaxTokens = 4096,
    [int]$MaxBulkExportRows = 2000,
    [string]$CertThumbprint = "",
    [switch]$KeepIntermediate,
    [switch]$ExeOnly
)

$ErrorActionPreference = "Stop"

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Content
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Write-JsonUtf8NoBom {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$Value
    )

    Write-Utf8NoBom -Path $Path -Content ($Value | ConvertTo-Json -Depth 5)
}

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

function Test-LikelyOllamaEndpoint {
    param([string]$BaseUrl)

    $value = if ($null -eq $BaseUrl) { "" } else { $BaseUrl }
    return $value -match '://(localhost|127\.0\.0\.1|\[::1\]):11434(/|$)'
}

function Get-AvailableInstallerExePath {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$Tag
    )

    $preferredPath = Join-Path $Directory "OutlookAI-$Tag-Setup.exe"
    if (-not (Test-Path -LiteralPath $preferredPath)) {
        return $preferredPath
    }

    try {
        Remove-Item -LiteralPath $preferredPath -Force -ErrorAction Stop
        return $preferredPath
    } catch {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $fallbackPath = Join-Path $Directory "OutlookAI-$Tag-Setup-$timestamp.exe"
        Write-Warning "Existing installer EXE is locked and cannot be replaced: $preferredPath"
        Write-Warning "Building a new installer instead: $fallbackPath"
        return $fallbackPath
    }
}

function Wait-FileReadable {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastError = $null
    while ((Get-Date) -lt $deadline) {
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            $stream.Dispose()
            return
        } catch {
            $lastError = $_.Exception.Message
            Start-Sleep -Milliseconds 500
        }
    }

    throw "File is still locked after $TimeoutSeconds seconds: $Path. Last error: $lastError"
}

if ($Tag -notmatch '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
    throw "Tag '$Tag' is not a valid semver tag (e.g. v3.0.0 or v3.0.0-beta.1)."
}

if ((Test-LikelyOllamaEndpoint -BaseUrl $LiteLlmBaseUrl) -and $LiteLlmModel -notmatch '^ollama/') {
    Write-Warning "LiteLlmBaseUrl '$LiteLlmBaseUrl' looks like Ollama, not LiteLLM proxy. LiteLLM aliases like '$LiteLlmModel' require the LiteLLM proxy URL, usually http://localhost:4000/v1."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$releaseScript = Join-Path $repoRoot "Deploy\Make-ReleaseZip.ps1"
$setupTemplate = Join-Path $repoRoot "Deploy\OutlookAI-Setup.ps1"
$iexpress = Join-Path $env:WINDIR "System32\iexpress.exe"

if (-not (Test-Path -LiteralPath $releaseScript)) {
    throw "Missing release packaging script: $releaseScript"
}
if (-not (Test-Path -LiteralPath $setupTemplate)) {
    throw "Missing installer bootstrap script: $setupTemplate"
}
if (-not (Test-Path -LiteralPath $iexpress)) {
    throw "iexpress.exe was not found. This script must run on Windows with IExpress available."
}

Push-Location $repoRoot
try {
    $resolvedOutDir = if ([System.IO.Path]::IsPathRooted($OutDir)) {
        $OutDir
    } else {
        Join-Path $repoRoot $OutDir
    }
    if (-not (Test-Path -LiteralPath $resolvedOutDir)) {
        New-Item -ItemType Directory -Path $resolvedOutDir | Out-Null
    }

    $releaseArgs = @{
        Tag = $Tag
        OutDir = $resolvedOutDir
    }
    if (-not [string]::IsNullOrWhiteSpace($CertThumbprint)) {
        $releaseArgs.CertThumbprint = $CertThumbprint
    }
    & $releaseScript @releaseArgs

    $zipPath = Join-Path $resolvedOutDir "OutlookAI-$Tag-RDS-Deploy.zip"
    if (-not (Test-Path -LiteralPath $zipPath)) {
        throw "Expected release ZIP was not created: $zipPath"
    }

    $packageRoot = Join-Path $repoRoot "out\installer-$Tag"
    if (Test-Path -LiteralPath $packageRoot) {
        Remove-Item -LiteralPath $packageRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $packageRoot | Out-Null

    $payloadZip = Join-Path $packageRoot "OutlookAI-Payload.zip"
    Copy-Item -LiteralPath $zipPath -Destination $payloadZip -Force

    $setupPs1 = Join-Path $packageRoot "OutlookAI-Setup.ps1"
    $LiteLlmVoiceModel = Normalize-OptionalValue $LiteLlmVoiceModel
    Copy-Item -LiteralPath $setupTemplate -Destination $setupPs1 -Force

    $installerConfigPath = Join-Path $packageRoot "OutlookAI-InstallerConfig.json"
    Write-JsonUtf8NoBom -Path $installerConfigPath -Value ([ordered]@{
        LiteLlmBaseUrl = $LiteLlmBaseUrl
        LiteLlmModel = $LiteLlmModel
        LiteLlmVoiceModel = $LiteLlmVoiceModel
        Temperature = $Temperature
        MaxTokens = $MaxTokens
        MaxBulkExportRows = $MaxBulkExportRows
    })

    $sedPath = Join-Path $packageRoot "OutlookAI-Installer.sed"
    $exePath = Get-AvailableInstallerExePath -Directory $resolvedOutDir -Tag $Tag

    $sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=1
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=OutlookAI installation finished.
TargetName=$exePath
FriendlyName=OutlookAI Setup
AppLaunched=powershell.exe -NoProfile -ExecutionPolicy Bypass -File OutlookAI-Setup.ps1
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
SourceFiles=SourceFiles

[Strings]
FILE0=OutlookAI-Setup.ps1
FILE1=OutlookAI-Payload.zip
FILE2=OutlookAI-InstallerConfig.json

[SourceFiles]
SourceFiles0=$packageRoot

[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
"@
    Write-Utf8NoBom -Path $sedPath -Content $sedContent

    & $iexpress /N /Q $sedPath
    $iexpressExitCode = $LASTEXITCODE
    $deadline = (Get-Date).AddSeconds(15)
    while (-not (Test-Path -LiteralPath $exePath) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-Path -LiteralPath $exePath)) {
        throw "IExpress failed with exit code $iexpressExitCode and did not create installer EXE: $exePath"
    }
    if ($iexpressExitCode -ne 0) {
        Write-Host "WARN: IExpress returned exit code $iexpressExitCode, but installer EXE was created." -ForegroundColor Yellow
    }

    Wait-FileReadable -Path $exePath
    $sha = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -LiteralPath ($exePath + ".sha256") -Encoding ASCII -Value $sha -NoNewline

    Write-Host "Built $exePath" -ForegroundColor Green
    Write-Host "SHA256 $sha" -ForegroundColor Green

    if ($ExeOnly) {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath ($zipPath + ".sha256") -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath ($exePath + ".sha256") -Force -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $resolvedOutDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^(RCX.*\.tmp|~OutlookAI-.*-Setup\.CAB)$' } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    if (-not $KeepIntermediate) {
        Remove-Item -LiteralPath $packageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
finally {
    Pop-Location
}
