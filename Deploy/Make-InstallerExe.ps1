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

function ConvertTo-PowerShellLiteral {
    param([AllowEmptyString()][string]$Value = "")
    return "'" + $Value.Replace("'", "''") + "'"
}

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Content
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
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

if ($Tag -notmatch '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
    throw "Tag '$Tag' is not a valid semver tag (e.g. v3.0.0 or v3.0.0-beta.1)."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$releaseScript = Join-Path $repoRoot "Deploy\Make-ReleaseZip.ps1"
$iexpress = Join-Path $env:WINDIR "System32\iexpress.exe"

if (-not (Test-Path -LiteralPath $releaseScript)) {
    throw "Missing release packaging script: $releaseScript"
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
    $baseUrlLiteral = ConvertTo-PowerShellLiteral $LiteLlmBaseUrl
    $modelLiteral = ConvertTo-PowerShellLiteral $LiteLlmModel
    $LiteLlmVoiceModel = Normalize-OptionalValue $LiteLlmVoiceModel
    $voiceModelLiteral = ConvertTo-PowerShellLiteral $LiteLlmVoiceModel
    $setupContent = @"
`$ErrorActionPreference = "Stop"

function Test-IsAdministrator {
    `$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    `$principal = New-Object Security.Principal.WindowsPrincipal(`$identity)
    return `$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

`$payload = Join-Path `$PSScriptRoot "OutlookAI-Payload.zip"
if (-not (Test-Path -LiteralPath `$payload)) {
    throw "Installer payload is missing: `$payload"
}

`$extractRoot = Join-Path `$env:TEMP ("OutlookAI-Setup-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path `$extractRoot | Out-Null

try {
    Expand-Archive -LiteralPath `$payload -DestinationPath `$extractRoot -Force
    `$installer = Join-Path `$extractRoot "Install-OutlookAI.ps1"
    if (-not (Test-Path -LiteralPath `$installer)) {
        throw "Install-OutlookAI.ps1 was not found in the installer payload."
    }

    `$installArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", `$installer,
        "-SourcePath", `$extractRoot,
        "-LiteLlmBaseUrl", $baseUrlLiteral,
        "-LiteLlmModel", $modelLiteral,
        "-LiteLlmVoiceModel", $voiceModelLiteral,
        "-Temperature", "$Temperature",
        "-MaxTokens", "$MaxTokens",
        "-MaxBulkExportRows", "$MaxBulkExportRows"
    )

    if (-not (Test-IsAdministrator)) {
        `$process = Start-Process -FilePath "powershell.exe" -ArgumentList `$installArgs -Verb RunAs -Wait -PassThru
        exit `$process.ExitCode
    }

    & powershell.exe @installArgs
    exit `$LASTEXITCODE
}
finally {
    if (Test-Path -LiteralPath `$extractRoot) {
        Remove-Item -LiteralPath `$extractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
"@
    Write-Utf8NoBom -Path $setupPs1 -Content $setupContent

    $sedPath = Join-Path $packageRoot "OutlookAI-Installer.sed"
    $exePath = Join-Path $resolvedOutDir "OutlookAI-$Tag-Setup.exe"
    if (Test-Path -LiteralPath $exePath) {
        Remove-Item -LiteralPath $exePath -Force
    }

    $sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=0
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

[SourceFiles]
SourceFiles0=$packageRoot

[SourceFiles0]
%FILE0%=
%FILE1%=
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

    $sha = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -LiteralPath ($exePath + ".sha256") -Encoding ASCII -Value $sha -NoNewline

    Write-Host "Built $exePath" -ForegroundColor Green
    Write-Host "SHA256 $sha" -ForegroundColor Green

    if ($ExeOnly) {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath ($zipPath + ".sha256") -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath ($exePath + ".sha256") -Force -ErrorAction SilentlyContinue
    }

    if (-not $KeepIntermediate) {
        Remove-Item -LiteralPath $packageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
finally {
    Pop-Location
}
