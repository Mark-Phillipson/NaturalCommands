<#
fix-log-encoding.ps1

Usage:
  pwsh -File .\scripts\fix-log-encoding.ps1 -Path .\bin\app.log -Out .\bin\app.fixed.log -AutoFix -ReplaceEmoji

What it does:
 - Scans the log for common mojibake sequences (e.g. â€” instead of —)
 - Attempts to recover text by re-encoding from Windows-1252 -> UTF8
 - Optionally scans for emoji names from emoji_mappings.json and suggests replacements
 - When -AutoFix is provided, writes a fixed UTF-8 output file to -Out
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$Path = ".\bin\app.log",

    [Parameter(Mandatory=$false)]
    [string]$Out = ".\bin\app.fixed.log",

    [switch]$AutoFix,

    [switch]$ReplaceEmoji
)

try {
    $inPath = (Resolve-Path $Path -ErrorAction Stop).Path
} catch {
    Write-Error "Input path not found: $Path"
    exit 2
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$repoRoot = Resolve-Path (Join-Path $scriptDir "..")
$emojiCandidates = @(Join-Path $repoRoot "emoji_mappings.json", Join-Path (Get-Location) "emoji_mappings.json")
$emojiMap = @{}
foreach ($p in $emojiCandidates) {
    if (Test-Path $p) {
        try {
            $emojiMap = Get-Content $p -Raw | ConvertFrom-Json
            Write-Host "Loaded emoji mappings from: $p"
        } catch {
            Write-Warning "Could not parse emoji mappings at $p"
            $emojiMap = @{}
        }
        break
    }
}

# Read bytes and decode as UTF-8 (typical file encoding)
[byte[]]$bytes = [System.IO.File]::ReadAllBytes($inPath)
[string]$text = [System.Text.Encoding]::UTF8.GetString($bytes)
[string[]]$lines = $text -split "`r?`n"

# Regex to catch common mojibake patterns (adjust as needed)
$mojibakeRegex = '(â|Ã|â€“|â€”|â€œ|â€�|â€™|â€¦|Ã©|Ã¤|Ã¶|Ã¼)'
$changed = $false
$resultLines = New-Object System.Collections.Generic.List[string]

foreach ($line in $lines) {
    $fixedLine = $line
    $found = $false

    if ($line -match $mojibakeRegex) {
        # Try to recover: treat current text as CP1252 bytes then decode as UTF8
        try {
            $candidate = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding(1252).GetBytes($line))
            if ($candidate -ne $line) {
                Write-Host "Mojibake detected:" -ForegroundColor Yellow
                Write-Host "  Original: $line"
                Write-Host "  Proposed: $candidate"
                $found = $true
                if ($AutoFix) {
                    $fixedLine = $candidate
                    $changed = $true
                }
            }
        } catch {
            # ignore
        }
    }

    if ($ReplaceEmoji -and $emojiMap -ne $null -and $emojiMap.PSObject.Properties.Count -gt 0) {
        foreach ($prop in $emojiMap.PSObject.Properties) {
            $key = $prop.Name
            $val = $prop.Value
            if ($fixedLine -match [regex]::Escape($key)) {
                Write-Host "Possible emoji name '$key' found in line:" -ForegroundColor Cyan
                Write-Host "  $fixedLine"
                Write-Host "  Suggest replace with: $val"
                $found = $true
                if ($AutoFix) {
                    $fixedLine = $fixedLine -replace [regex]::Escape($key), $val
                    $changed = $true
                }
            }
        }
    }

    $resultLines.Add($fixedLine)
}

if ($AutoFix -and $changed) {
    try {
        $outDir = Split-Path -Parent $Out
        if (-not [string]::IsNullOrEmpty($outDir) -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        [System.IO.File]::WriteAllText($Out, ($resultLines -join "`n"), [System.Text.Encoding]::UTF8)
        Write-Host "Wrote fixed file to $Out" -ForegroundColor Green
    } catch {
        Write-Error "Failed to write output file: $_"
    }
} else {
    if (-not $AutoFix) {
        Write-Host "Preview only. Use -AutoFix to write a fixed copy to -Out." -ForegroundColor Gray
    } elseif (-not $changed) {
        Write-Host "No mojibake or emoji candidates found; no changes made." -ForegroundColor Gray
    }
}
