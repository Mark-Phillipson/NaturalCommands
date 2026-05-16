<#
Sends test Windows toast notifications so the running app can pick them up.
Usage:
  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\send-test-notifications.ps1 -Count 5 -DelayMs 700 -AppId "NaturalCommands-Test"

Options:
  -Count    Number of notifications to send (default 5)
  -DelayMs  Milliseconds between notifications (default 700)
  -AppId    Toast notifier AppId used for CreateToastNotifier (default "NaturalCommands-Test")
#>

param(
    [int]$Count = 5,
    [int]$DelayMs = 700,
    [string]$AppId = "NaturalCommands-Test",
    [switch]$Verbose
)

function Send-ToastWinRT {
    param($Title, $Body, $AppId)
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
        $safeTitle = [System.Security.SecurityElement]::Escape($Title)
        $safeBody = [System.Security.SecurityElement]::Escape($Body)
        $xml = "<toast><visual><binding template='ToastGeneric'><text>$safeTitle</text><text>$safeBody</text></binding></visual></toast>"
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($xml)
        $toast = [Windows.UI.Notifications.ToastNotification]::new($doc)
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)
        $notifier.Show($toast)
        if ($Verbose) { Write-Host "[WinRT] Sent toast: $Title" }
        return $true
    } catch {
        if ($Verbose) { Write-Host "[WinRT] Failed to send toast: $_" }
        return $false
    }
}

function Send-ToastBurnt {
    param($Title, $Body)
    try {
        if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
            New-BurntToastNotification -Text $Title, $Body
            if ($Verbose) { Write-Host "[BurntToast] Sent toast: $Title" }
            return $true
        }
    } catch {
        if ($Verbose) { Write-Host "[BurntToast] Failed: $_" }
    }
    return $false
}

function Send-ToastFallback {
    param($Title, $Body)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.Visible = $true
        $notify.ShowBalloonTip(4000, $Title, $Body, [System.Windows.Forms.ToolTipIcon]::Info)
        Start-Sleep -Seconds 4
        $notify.Dispose()
        if ($Verbose) { Write-Host "[Fallback] Showed balloon: $Title" }
        return $true
    } catch {
        if ($Verbose) { Write-Host "[Fallback] Failed: $_" }
        return $false
    }
}

Write-Host "Sending $Count test notifications (AppId='$AppId')..."
for ($i = 1; $i -le $Count; $i++) {
    $title = "Test Notification #$i"
    $body = "Payload generated at $(Get-Date -Format s)"

    $sent = Send-ToastWinRT -Title $title -Body $body -AppId $AppId
    if (-not $sent) {
        $sent = Send-ToastBurnt -Title $title -Body $body
    }
    if (-not $sent) {
        Send-ToastFallback -Title $title -Body $body | Out-Null
    }

    Start-Sleep -Milliseconds $DelayMs
}
Write-Host "Done sending notifications."
