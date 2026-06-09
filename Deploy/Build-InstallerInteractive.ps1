<#
.SYNOPSIS
    Interactively builds a single-file OutlookAI installer EXE.

.DESCRIPTION
    Prompts for the LiteLLM installer settings, calls
    Deploy\Make-InstallerExe.ps1, and opens the output folder with the
    generated EXE selected.

    The LiteLLM API key is intentionally not requested or embedded.
#>

$ErrorActionPreference = "Stop"

function Read-Value {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)][string]$DefaultValue
    )

    $value = Read-Host "$Prompt [$DefaultValue]"
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }

    return $value.Trim()
}

function Read-DoubleValue {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)][double]$DefaultValue
    )

    while ($true) {
        $text = Read-Value -Prompt $Prompt -DefaultValue ([string]::Format([Globalization.CultureInfo]::InvariantCulture, "{0}", $DefaultValue))
        $parsed = 0.0
        if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            return $parsed
        }

        Write-Host "Enter a number, for example 0.2." -ForegroundColor Yellow
    }
}

function Read-IntValue {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)][int]$DefaultValue
    )

    while ($true) {
        $text = Read-Value -Prompt $Prompt -DefaultValue ([string]$DefaultValue)
        $parsed = 0
        if ([int]::TryParse($text, [ref]$parsed) -and $parsed -gt 0) {
            return $parsed
        }

        Write-Host "Enter a positive integer." -ForegroundColor Yellow
    }
}

function Read-Tag {
    while ($true) {
        $tag = Read-Value -Prompt "Release tag" -DefaultValue "v3.0.0"
        if ($tag -match '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
            return $tag
        }

        Write-Host "Tag must look like v3.0.0 or v3.0.0-beta.1." -ForegroundColor Yellow
    }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$builder = Join-Path $repoRoot "Deploy\Make-InstallerExe.ps1"

if (-not (Test-Path -LiteralPath $builder)) {
    throw "Build script was not found: $builder"
}

Write-Host ""
Write-Host "OutlookAI single-file installer builder" -ForegroundColor Cyan
Write-Host "Press Enter to accept the value in brackets." -ForegroundColor Gray
Write-Host "LiteLLM API keys are not requested and will not be embedded." -ForegroundColor Gray
Write-Host ""

$tag = Read-Tag
$outDir = Read-Value -Prompt "Output directory" -DefaultValue "out"
$baseUrl = Read-Value -Prompt "LiteLLM base URL" -DefaultValue "https://litellm.company.example/v1"
$model = Read-Value -Prompt "LiteLLM chat model" -DefaultValue "company/outlook-chat"
$voiceModel = Read-Value -Prompt "LiteLLM voice model" -DefaultValue "company/outlook-transcribe"
$temperature = Read-DoubleValue -Prompt "Temperature" -DefaultValue 0.2
$maxTokens = Read-IntValue -Prompt "Max tokens" -DefaultValue 4096
$maxBulkExportRows = Read-IntValue -Prompt "Max bulk export rows" -DefaultValue 2000
$certThumbprint = Read-Value -Prompt "Manifest certificate thumbprint (optional)" -DefaultValue ""

$arguments = @{
    Tag = $tag
    OutDir = $outDir
    LiteLlmBaseUrl = $baseUrl
    LiteLlmModel = $model
    LiteLlmVoiceModel = $voiceModel
    Temperature = $temperature
    MaxTokens = $maxTokens
    MaxBulkExportRows = $maxBulkExportRows
}

if (-not [string]::IsNullOrWhiteSpace($certThumbprint)) {
    $arguments.CertThumbprint = $certThumbprint
}

Write-Host ""
Write-Host "Building installer..." -ForegroundColor Cyan
& $builder @arguments

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$resolvedOutDir = if ([System.IO.Path]::IsPathRooted($outDir)) {
    $outDir
} else {
    Join-Path $repoRoot $outDir
}
$exePath = Join-Path $resolvedOutDir "OutlookAI-$tag-Setup.exe"

if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Expected installer EXE was not found: $exePath"
}

Write-Host ""
Write-Host "Built: $exePath" -ForegroundColor Green
Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$exePath`""
