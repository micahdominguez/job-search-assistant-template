$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $root

Write-Host ""
Write-Host "Candidate-specific fallback hotspots"
Write-Host "==================================="
Write-Host ""

$patterns = @(
    "Your Name",
    "you@example\.com",
    "linkedin\.com/in/your-linkedin",
    "@yourhandle",
    "555-555-5555",
    "Messari",
    "Mercer",
    "FIO"
) -join "|"

rg -n $patterns .\job_search_assistant.py .\profile.json
