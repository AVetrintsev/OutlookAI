<#
.SYNOPSIS
    Builds and tests OutlookAI without registering the add-in in Outlook.
.DESCRIPTION
    Requires Windows, Visual Studio/Build Tools with Office/VSTO targets,
    .NET SDK, .NET Framework 4.7.2 targeting pack and classic Office.
    Uses full-framework MSBuild for COM references. Test binaries are not
    signed deployment artifacts: publishing and Office registration are skipped.
.PARAMETER WordCom
    Opts into integration tests using a new Word instance and unsaved test
    documents. Never opens the user's documents or mailbox.
#>
[CmdletBinding()]
param(
    [string]$Filter,
    [switch]$WordCom,
    [switch]$NoRestore
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$testProject = Join-Path $repoRoot 'VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'Install Visual Studio/Build Tools with Office/VSTO build tools first.'
}
$installation = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild Microsoft.VisualStudio.Component.TeamOffice.BuildTools -property installationPath
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($installation)) {
    throw 'An MSBuild installation with Office/VSTO build tools was not found.'
}
$msbuild = Join-Path $installation 'MSBuild\Current\Bin\MSBuild.exe'
$vstoDirectory = Get-ChildItem -LiteralPath (Join-Path $installation 'MSBuild\Microsoft\VisualStudio') -Directory |
    Where-Object { $_.Name -match '^v\d+\.\d+$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'OfficeTools\Microsoft.VisualStudio.Tools.Office.targets')) } |
    Sort-Object { [version]$_.Name.Substring(1) } -Descending | Select-Object -First 1
if (-not $vstoDirectory) { throw 'Office/VSTO targets were not found.' }
Get-Command dotnet -ErrorAction Stop | Out-Null

$environment = @{
    DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    DOTNET_ADD_GLOBAL_TOOLS_TO_PATH = 'false'
    DOTNET_GENERATE_ASPNET_CERTIFICATE = 'false'
    OUTLOOKAI_RUN_WORD_COM_TESTS = $(if ($WordCom) { '1' } else { $null })
}
$previousEnvironment = @{}
try {
    foreach ($settingName in $environment.Keys) {
        $previousEnvironment[$settingName] = [Environment]::GetEnvironmentVariable($settingName, 'Process')
        [Environment]::SetEnvironmentVariable($settingName, $environment[$settingName], 'Process')
    }
    $properties = @("/p:VSToolsPath=$($vstoDirectory.FullName)", "/p:VisualStudioVersion=$($vstoDirectory.Name.Substring(1))")
    if (-not $NoRestore) {
        & dotnet restore $testProject @properties --verbosity minimal
        if ($LASTEXITCODE -ne 0) { throw "Restore failed: $LASTEXITCODE" }
    }

    # VSTO's normal PrepareForRun registers the add-in and touches certificates.
    # Unit tests need the real compiled DLL and resources, not registration.
    & $msbuild $testProject @properties /t:Build /p:Configuration=Debug /p:SignManifests=false /p:GenerateManifests=false /p:PrepareForRunDependsOn=CopyFilesToOutputDirectory /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $LASTEXITCODE" }

    $testArguments = @('test', $testProject, '--no-build', '--no-restore', '--verbosity', 'minimal', '--logger', 'trx;LogFileName=tests.trx', '--results-directory', (Join-Path $repoRoot 'out\tests')) + $properties
    if (-not [string]::IsNullOrWhiteSpace($Filter)) { $testArguments += @('--filter', $Filter) }
    & dotnet @testArguments
    if ($LASTEXITCODE -ne 0) { throw "Tests failed: $LASTEXITCODE" }
}
finally {
    foreach ($settingName in $previousEnvironment.Keys) {
        [Environment]::SetEnvironmentVariable($settingName, $previousEnvironment[$settingName], 'Process')
    }
}
