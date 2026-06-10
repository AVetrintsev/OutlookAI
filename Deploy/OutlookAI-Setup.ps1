param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "OutlookAI-InstallerConfig.json")
)

$ErrorActionPreference = "Stop"

$script:LogPath = Join-Path $env:TEMP ("OutlookAI-Installer-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
$script:ExtractRoot = $null

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-LogLine {
    param(
        [AllowEmptyString()][string]$Message = "",
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::Gray
    )

    Write-Host $Message -ForegroundColor $ForegroundColor
    try {
        Add-Content -LiteralPath $script:LogPath -Value $Message -Encoding UTF8
    } catch {
        # Logging must not break installation.
    }
}

function Wait-BeforeExit {
    param([int]$ExitCode)

    Write-LogLine ""
    Write-LogLine ("Installer log: " + $script:LogPath) ([ConsoleColor]::Yellow)
    Write-Host ""
    Read-Host "Press Enter to close this window"
    exit $ExitCode
}

function Read-InstallerConfig {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Installer config is missing: $Path"
    }

    return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Add-RequiredInstallArgument {
    param(
        [System.Collections.Generic.List[string]]$Arguments,
        [string]$Name,
        [object]$Value
    )

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "Installer config value '$Name' is required."
    }

    $Arguments.Add($Name)
    $Arguments.Add($text)
}

function Add-OptionalInstallArgument {
    param(
        [System.Collections.Generic.List[string]]$Arguments,
        [string]$Name,
        [object]$Value
    )

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $Arguments.Add($Name)
    $Arguments.Add($text)
}

function Write-ProcessOutput {
    param([string[]]$Paths)

    foreach ($outputPath in $Paths) {
        if (-not (Test-Path -LiteralPath $outputPath)) {
            continue
        }

        Get-Content -LiteralPath $outputPath -ErrorAction SilentlyContinue | ForEach-Object {
            $line = if ($null -eq $_) { "" } else { $_.ToString() }
            Write-Host $line
            try {
                Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
            } catch {
                # Logging must not break installation.
            }
        }
        Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
    }
}

try {
    Write-LogLine "OutlookAI installer started." ([ConsoleColor]::Cyan)
    Write-LogLine ("Log file: " + $script:LogPath) ([ConsoleColor]::Gray)

    $config = Read-InstallerConfig -Path $ConfigPath

    if (-not (Test-IsAdministrator)) {
        Write-LogLine "Requesting administrator privileges..." ([ConsoleColor]::Yellow)
        $elevatedArgs = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $PSCommandPath,
            "-ConfigPath", $ConfigPath
        )
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList $elevatedArgs -Verb RunAs -Wait -PassThru
        exit $process.ExitCode
    }

    $payload = Join-Path $PSScriptRoot "OutlookAI-Payload.zip"
    if (-not (Test-Path -LiteralPath $payload)) {
        throw "Installer payload is missing: $payload"
    }

    $script:ExtractRoot = Join-Path $env:TEMP ("OutlookAI-Setup-" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $script:ExtractRoot | Out-Null

    Expand-Archive -LiteralPath $payload -DestinationPath $script:ExtractRoot -Force
    $installer = Join-Path $script:ExtractRoot "Install-OutlookAI.ps1"
    if (-not (Test-Path -LiteralPath $installer)) {
        throw "Install-OutlookAI.ps1 was not found in the installer payload."
    }

    $installArgs = [System.Collections.Generic.List[string]]::new()
    $installArgs.Add("-NoProfile")
    $installArgs.Add("-ExecutionPolicy")
    $installArgs.Add("Bypass")
    $installArgs.Add("-File")
    $installArgs.Add($installer)
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-SourcePath" -Value $script:ExtractRoot
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-LiteLlmBaseUrl" -Value $config.LiteLlmBaseUrl
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-LiteLlmModel" -Value $config.LiteLlmModel
    Add-OptionalInstallArgument -Arguments $installArgs -Name "-LiteLlmVoiceModel" -Value $config.LiteLlmVoiceModel
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-Temperature" -Value $config.Temperature
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-MaxTokens" -Value $config.MaxTokens
    Add-RequiredInstallArgument -Arguments $installArgs -Name "-MaxBulkExportRows" -Value $config.MaxBulkExportRows

    Write-LogLine "Installer parameters:" ([ConsoleColor]::Cyan)
    Write-LogLine ("  LiteLLM base URL : " + $config.LiteLlmBaseUrl)
    Write-LogLine ("  LiteLLM model    : " + $config.LiteLlmModel)
    Write-LogLine ("  LiteLLM voice    : " + $config.LiteLlmVoiceModel)
    Write-LogLine ""

    $stdoutPath = Join-Path $env:TEMP ("OutlookAI-Installer-" + [Guid]::NewGuid().ToString("N") + ".stdout.log")
    $stderrPath = Join-Path $env:TEMP ("OutlookAI-Installer-" + [Guid]::NewGuid().ToString("N") + ".stderr.log")
    $startProcessArgs = @{
        FilePath = "powershell.exe"
        ArgumentList = $installArgs.ToArray()
        Wait = $true
        PassThru = $true
        WindowStyle = "Hidden"
        RedirectStandardOutput = $stdoutPath
        RedirectStandardError = $stderrPath
    }
    $process = Start-Process @startProcessArgs
    Write-ProcessOutput -Paths @($stdoutPath, $stderrPath)

    if ($process.ExitCode -eq 0) {
        Write-LogLine ""
        Write-LogLine "OutlookAI installation completed." ([ConsoleColor]::Green)
    } else {
        Write-LogLine ""
        Write-LogLine ("OutlookAI installation failed with exit code " + $process.ExitCode + ".") ([ConsoleColor]::Red)
    }

    Wait-BeforeExit -ExitCode $process.ExitCode
}
catch {
    Write-LogLine ""
    Write-LogLine "OutlookAI installer failed." ([ConsoleColor]::Red)
    Write-LogLine ""
    Write-LogLine "Error:" ([ConsoleColor]::Yellow)
    Write-LogLine $_.Exception.Message ([ConsoleColor]::Red)

    if ($_.InvocationInfo) {
        Write-LogLine ""
        Write-LogLine "Location:" ([ConsoleColor]::Yellow)
        Write-LogLine $_.InvocationInfo.PositionMessage
    }

    if ($_.ScriptStackTrace) {
        Write-LogLine ""
        Write-LogLine "Stack trace:" ([ConsoleColor]::Yellow)
        Write-LogLine $_.ScriptStackTrace
    }

    Wait-BeforeExit -ExitCode 1
}
finally {
    if ($script:ExtractRoot -and (Test-Path -LiteralPath $script:ExtractRoot)) {
        Remove-Item -LiteralPath $script:ExtractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
