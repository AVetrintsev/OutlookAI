#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Uninstalls OutlookAI v3 from this server.

.DESCRIPTION
    Removes:
      - HKLM Outlook add-in registration (64-bit + WOW6432Node)
      - C:\Program Files\OutlookAI install directory
      - C:\ProgramData\OutlookAI runtime directory except Backups

    Preserves:
      - C:\ProgramData\OutlookAI\Backups (v1 config rollback artifacts)

    Per-user LiteLLM API keys live under each user's AppData config. This
    script does not enumerate and delete user profiles.

    Run as Administrator.

.EXAMPLE
    .\Uninstall-OutlookAI.ps1
#>

param()

$ErrorActionPreference = "Stop"
$InstallPath     = "C:\Program Files\OutlookAI"
$ProgramDataPath = "C:\ProgramData\OutlookAI"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI v3 Uninstaller" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

# --- 1. Registry ---------------------------------------------------------
Write-Host "[1/3] Removing add-in registry entries..." -ForegroundColor Yellow
foreach ($path in @(
    "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
)) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        Write-Host "  Removed: $path" -ForegroundColor Gray
    }
}
Write-Host "  Done." -ForegroundColor Green

# --- 2. Install directory ------------------------------------------------
Write-Host "[2/3] Removing install files..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
    Write-Host "  Removed: $InstallPath" -ForegroundColor Gray
} else {
    Write-Host "  Already absent: $InstallPath" -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

# --- 3. ProgramData runtime files (preserve Backups) ---------------------
Write-Host "[3/3] Removing ProgramData runtime files (preserving Backups)..." -ForegroundColor Yellow
if (Test-Path $ProgramDataPath) {
    Get-ChildItem -Path $ProgramDataPath -Force | Where-Object { $_.Name -ne "Backups" } | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
        Write-Host "  Removed: $($_.FullName)" -ForegroundColor Gray
    }
}
Write-Host "  Preserved backups under $ProgramDataPath\Backups" -ForegroundColor Gray
Write-Host "  Done." -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Uninstall Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Users will need to restart Outlook for changes to take effect." -ForegroundColor White
Write-Host ""
Write-Host "NOTE: Per-user LiteLLM API keys are stored in each user's AppData config." -ForegroundColor Yellow
Write-Host "      Delete %APPDATA%\\OutlookAI\\config.xml for a user to remove it." -ForegroundColor Yellow
Write-Host ""
