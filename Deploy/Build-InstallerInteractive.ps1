<#
.SYNOPSIS
    Interactively builds a single-file OutlookAI installer EXE.

.DESCRIPTION
    Prompts for the LiteLLM installer settings, calls
    Deploy\Make-InstallerExe.ps1, and opens the output folder with the
    generated EXE selected.

    The LiteLLM API key is intentionally not requested or embedded.
#>

param(
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"

function Read-Value {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [AllowEmptyString()][string]$DefaultValue = ""
    )

    $displayPrompt = if ([string]::IsNullOrEmpty($DefaultValue)) {
        $Prompt
    } else {
        "$Prompt [$DefaultValue]"
    }

    $value = Read-Host $displayPrompt
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

function Read-BoolValue {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)][bool]$DefaultValue
    )

    $defaultText = if ($DefaultValue) { "yes" } else { "no" }
    while ($true) {
        $text = (Read-Value -Prompt $Prompt -DefaultValue $defaultText).ToLowerInvariant()
        if ($text -in @("yes", "y", "true", "1", "да", "д")) {
            return $true
        }
        if ($text -in @("no", "n", "false", "0", "нет", "н")) {
            return $false
        }

        Write-Host "Enter yes or no." -ForegroundColor Yellow
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

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Read-LocalBuildConfig {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    } catch {
        Write-Host "Could not read local build config; defaults will be used: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Get-ConfigValue {
    param(
        $Config,
        [string]$Name,
        [AllowEmptyString()]$DefaultValue
    )

    if ($null -eq $Config) {
        return $DefaultValue
    }

    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }

    $text = [string]$property.Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $DefaultValue
    }

    return $text
}

function Save-LocalBuildConfig {
    param(
        [string]$Path,
        [Parameter(Mandatory=$true)]$Value
    )

    $json = $Value | ConvertTo-Json -Depth 5
    Write-Utf8NoBom -Path $Path -Content $json
}

function Test-LikelyOllamaEndpoint {
    param([string]$BaseUrl)

    $value = if ($null -eq $BaseUrl) { "" } else { $BaseUrl }
    return $value -match '://(localhost|127\.0\.0\.1|\[::1\]):11434(/|$)'
}

function Confirm-LiteLlmEndpoint {
    param(
        [string]$BaseUrl,
        [string]$Model
    )

    if (-not (Test-LikelyOllamaEndpoint -BaseUrl $BaseUrl)) {
        return
    }

    Write-Host ""
    Write-Host "WARNING: $BaseUrl looks like the Ollama API endpoint, not the LiteLLM proxy endpoint." -ForegroundColor Yellow
    Write-Host "LiteLLM model aliases such as '$Model' are available only through the LiteLLM proxy." -ForegroundColor Yellow
    Write-Host "Use the LiteLLM proxy URL, usually http://localhost:4000/v1, with model local-model." -ForegroundColor Yellow
    Write-Host "If you intentionally connect directly to Ollama, use an Ollama model name such as qwen2.5:3b." -ForegroundColor Yellow
    Write-Host ""
    $answer = Read-Host "Continue with this endpoint anyway? Type YES to continue"
    if ($answer -ne "YES") {
        throw "Build cancelled. Re-run the script with the LiteLLM proxy base URL."
    }
}

function Wait-BeforeExit {
    param([int]$ExitCode)

    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Press Enter to close this window"
    }

    exit $ExitCode
}

try {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
    $builder = Join-Path $repoRoot "Deploy\Make-InstallerExe.ps1"
    $localConfigPath = Join-Path $PSScriptRoot "Build-InstallerInteractive.local.json"
    $localConfig = Read-LocalBuildConfig -Path $localConfigPath

    if (-not (Test-Path -LiteralPath $builder)) {
        throw "Build script was not found: $builder"
    }

    Write-Host ""
    Write-Host "OutlookAI single-file installer builder" -ForegroundColor Cyan
    Write-Host "Press Enter to accept the value in brackets." -ForegroundColor Gray
    Write-Host "LiteLLM API keys are not requested and will not be embedded." -ForegroundColor Gray
    if ($localConfig) {
        Write-Host "Defaults loaded from $localConfigPath" -ForegroundColor Gray
    }
    Write-Host ""

    $tag = Read-Value -Prompt "Release tag" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "Tag" -DefaultValue "v3.0.0")
    while ($tag -notmatch '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
        Write-Host "Tag must look like v3.0.0 or v3.0.0-beta.1." -ForegroundColor Yellow
        $tag = Read-Value -Prompt "Release tag" -DefaultValue "v3.0.0"
    }
    $outDir = Read-Value -Prompt "Output directory" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "OutDir" -DefaultValue "out")
    $baseUrl = Read-Value -Prompt "LiteLLM base URL" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "LiteLlmBaseUrl" -DefaultValue "https://litellm.company.example/v1")
    $model = Read-Value -Prompt "LiteLLM chat model" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "LiteLlmModel" -DefaultValue "company/outlook-chat")
    Confirm-LiteLlmEndpoint -BaseUrl $baseUrl -Model $model
    $voiceModel = Normalize-OptionalValue (Read-Value -Prompt "LiteLLM voice model (optional, empty/null disables transcription)" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "LiteLlmVoiceModel" -DefaultValue ""))
    $temperature = Read-DoubleValue -Prompt "Temperature" -DefaultValue ([double](Get-ConfigValue -Config $localConfig -Name "Temperature" -DefaultValue 0.2))
    $maxTokens = Read-IntValue -Prompt "Max tokens" -DefaultValue ([int](Get-ConfigValue -Config $localConfig -Name "MaxTokens" -DefaultValue 4096))
    $maxBulkExportRows = Read-IntValue -Prompt "Max bulk export rows" -DefaultValue ([int](Get-ConfigValue -Config $localConfig -Name "MaxBulkExportRows" -DefaultValue 2000))
    $recommendationsEnabled = Read-BoolValue -Prompt "Enable AI recommendations" -DefaultValue ([System.Convert]::ToBoolean(
        (Get-ConfigValue -Config $localConfig -Name "RecommendationsEnabled" -DefaultValue $false)))
    $certThumbprint = Read-Value -Prompt "Manifest certificate thumbprint (optional)" -DefaultValue (Get-ConfigValue -Config $localConfig -Name "CertThumbprint" -DefaultValue "")

    Save-LocalBuildConfig -Path $localConfigPath -Value ([ordered]@{
        Tag = $tag
        OutDir = $outDir
        LiteLlmBaseUrl = $baseUrl
        LiteLlmModel = $model
        LiteLlmVoiceModel = $voiceModel
        Temperature = $temperature
        MaxTokens = $maxTokens
        MaxBulkExportRows = $maxBulkExportRows
        RecommendationsEnabled = $recommendationsEnabled
        CertThumbprint = $certThumbprint
    })
    Write-Host "Saved defaults to $localConfigPath" -ForegroundColor Gray

    $arguments = @{
        Tag = $tag
        OutDir = $outDir
        LiteLlmBaseUrl = $baseUrl
        LiteLlmModel = $model
        LiteLlmVoiceModel = $voiceModel
        Temperature = $temperature
        MaxTokens = $maxTokens
        MaxBulkExportRows = $maxBulkExportRows
        RecommendationsEnabled = $recommendationsEnabled
        ExeOnly = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($certThumbprint)) {
        $arguments.CertThumbprint = $certThumbprint
    }

    Write-Host ""
    Write-Host "Building installer..." -ForegroundColor Cyan
    & $builder @arguments

    if ($LASTEXITCODE -ne 0) {
        throw "Make-InstallerExe.ps1 failed with exit code $LASTEXITCODE."
    }

    $resolvedOutDir = if ([System.IO.Path]::IsPathRooted($outDir)) {
        $outDir
    } else {
        Join-Path $repoRoot $outDir
    }
    $exePath = Join-Path $resolvedOutDir "OutlookAI-$tag-Setup.exe"

    if (-not (Test-Path -LiteralPath $exePath)) {
        $exePath = Get-ChildItem -LiteralPath $resolvedOutDir -Filter "OutlookAI-$tag-Setup*.exe" |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if ([string]::IsNullOrWhiteSpace($exePath) -or -not (Test-Path -LiteralPath $exePath)) {
        throw "Expected installer EXE was not found in $resolvedOutDir."
    }

    Write-Host ""
    Write-Host "Built: $exePath" -ForegroundColor Green
    if (-not $NoPause) {
        try {
            Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$exePath`""
        } catch {
            Write-Host "Could not open Explorer: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host ""
    Write-Host "OutlookAI installer build failed." -ForegroundColor Red
    Write-Host ""
    Write-Host "Error:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red

    if ($_.InvocationInfo) {
        Write-Host ""
        Write-Host "Location:" -ForegroundColor Yellow
        Write-Host $_.InvocationInfo.PositionMessage
    }

    if ($_.ScriptStackTrace) {
        Write-Host ""
        Write-Host "Stack trace:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace
    }

    Wait-BeforeExit -ExitCode 1
}
