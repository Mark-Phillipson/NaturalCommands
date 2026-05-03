<#
Send a test Windows Toast notification for smoke-testing the WindowsNotificationListenerService.

Usage:
  powershell -ExecutionPolicy Bypass -File .\scripts\send-test-toast.ps1 -Title "Smoke Test" -Body "smoke_nonce:2026-05-03T12:00:00Z"

This script will install the BurntToast module for the current user if it is not present.
#>
Param(
    [string]$Title = "NaturalCommands Test Toast",
    [string]$Body = "This is a test notification for WindowsNotificationListenerService."
)

try {
    if (-not (Get-Module -ListAvailable -Name BurntToast)) {
        Write-Host "BurntToast module not found. Installing to CurrentUser scope..." -ForegroundColor Yellow
        Install-Module -Name BurntToast -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
    }
    Import-Module BurntToast -ErrorAction Stop
}
catch {
    Write-Error "Failed to ensure BurntToast is available: $($_.Exception.Message)"
    exit 1
}

$nonce = [Guid]::NewGuid().ToString()
$bodyWithNonce = "$Body`nnonce:$nonce"

Write-Host "Sending toast: $Title - $bodyWithNonce"
New-BurntToastNotification -Text $Title, $bodyWithNonce

Write-Host "Toast sent. nonce=$nonce"
