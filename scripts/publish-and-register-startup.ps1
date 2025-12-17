<#
.SYNOPSIS
    Publish the NaturalCommands project to the repository publish folder and optionally create a per-user Startup shortcut.

.DESCRIPTION
    This script runs `dotnet publish` for the project at the repository root (default ./NaturalCommands.csproj) and by default
    publishes a self-contained single-file build for win-x64 to the standard repo publish path:
      ./bin/Release/net10.0-windows/win-x64/publish

    After publishing it will create a Startup shortcut that runs the exe with the `-- listen` argument, so the app starts when the user logs in.

.EXAMPLE
    # Publish self-contained single-file and create startup shortcut
    powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -SelfContained -CreateStartupShortcut

.EXAMPLE
    # Publish framework-dependent build (smaller) and skip creating a shortcut
    powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -NoSelfContained -CreateStartupShortcut:$false

#>

param(
    [string]$ProjectPath = "./NaturalCommands.csproj",
    [string]$Configuration = "Release",
    [string]$Framework = "net10.0-windows",
    [string]$Runtime = "win-x64",
    [switch]$SelfContained = $true,
    [switch]$CreateStartupShortcut = $true
)

function Write-Info($m){ Write-Host "[INFO] $m" -ForegroundColor Green }
function Write-Err($m){ Write-Host "[ERROR] $m" -ForegroundColor Red }

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Err "dotnet CLI not found in PATH. Install .NET SDK or ensure dotnet is on PATH."
    exit 1
}

$publishPath = if ($SelfContained) { "./bin/$Configuration/$Framework/$Runtime/publish" } else { "./bin/$Configuration/$Framework/publish" }
$publishPath = Resolve-Path -Path $publishPath -ErrorAction Ignore | ForEach-Object { $_.ProviderPath }
if (-not $publishPath) { $publishPath = (Join-Path -Path (Get-Location) -ChildPath (if ($SelfContained) { "bin/$Configuration/$Framework/$Runtime/publish" } else { "bin/$Configuration/$Framework/publish" })) }


$publishArgs = @(
    $ProjectPath,
    "-c", $Configuration,
    "-f", $Framework,
    "-o", $publishPath
)

if ($SelfContained) {
    $publishArgs += @("-r", $Runtime, "--self-contained", "true", "-p:PublishSingleFile=true")
}

Write-Info "Publishing to: $publishPath"
Write-Info "Running: dotnet publish $($publishArgs -join ' ')"

$pub = dotnet publish @publishArgs 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Err "dotnet publish failed:`n$pub"
    exit $LASTEXITCODE
}

$exePath = Join-Path $publishPath "NaturalCommands.exe"
if (-not (Test-Path $exePath)) {
    Write-Err "Expected exe not found at: $exePath"
    exit 1
}

Write-Info "Publish succeeded. Exe: $exePath"

if ($CreateStartupShortcut) {
    try {
        $startupFolder = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
        if (-not (Test-Path $startupFolder)) { New-Item -Path $startupFolder -ItemType Directory -Force | Out-Null }

        $shortcutPath = Join-Path $startupFolder "NaturalCommands.lnk"

        $WshShell = New-Object -ComObject WScript.Shell
        $shortcut = $WshShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $exePath
        $shortcut.Arguments = "listen"
        $shortcut.WorkingDirectory = Split-Path $exePath -Parent
        $shortcut.Save()

        Write-Info "Startup shortcut created at: $shortcutPath"
    }
    catch {
        Write-Err "Failed to create Startup shortcut: $_"
        exit 1
    }
}

Write-Info "Done. To test: `"$exePath`" -- listen"
