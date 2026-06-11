param(
    [Parameter(Mandatory = $true)]
    [int]$JobId,
    [string]$Output,
    [switch]$PrintPacket,
    [switch]$LocalOnly
)

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$packetArgs = @(
    ".\job_search_assistant.py",
    "application-packet",
    "--id",
    $JobId,
    "--skip-google-sync"
)

if ($Output) {
    $packetArgs += @("--output", $Output)
}

if ($PrintPacket) {
    $packetArgs += "--print"
}

Push-Location $repoRoot
try {
    & python @packetArgs
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    if ($LocalOnly) {
        Write-Host "Skipped Google Drive sync by request."
        exit 0
    }

    & python .\job_search_assistant.py sync-google-drive-docs --auth-mode auto --id $JobId
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}
