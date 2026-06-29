[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$DailyRunArgs
)

$ErrorActionPreference = "Stop"

function Get-ArgValue {
    param(
        [string[]]$ArgsList,
        [string]$Name
    )
    for ($i = 0; $i -lt $ArgsList.Count; $i++) {
        if ($ArgsList[$i] -eq $Name -and $i + 1 -lt $ArgsList.Count) {
            return $ArgsList[$i + 1]
        }
    }
    return $null
}

function Invoke-PythonCommand {
    param(
        [string[]]$Arguments
    )

    & python @Arguments
    return $LASTEXITCODE
}

$dailyArgs = @(".\job_search_assistant.py", "daily-run", "--skip-sheet-sync") + $DailyRunArgs
$sheetArgs = @(".\job_search_assistant.py", "sync-google-sheets-workbook")

$spreadsheetUrl = Get-ArgValue -ArgsList $DailyRunArgs -Name "--spreadsheet-url"
if ($spreadsheetUrl) {
    $sheetArgs += @("--spreadsheet-url", $spreadsheetUrl)
}
$sheetAuthMode = Get-ArgValue -ArgsList $DailyRunArgs -Name "--sheet-auth-mode"
if ($sheetAuthMode) {
    $sheetArgs += @("--auth-mode", $sheetAuthMode)
}
$serviceAccountJson = Get-ArgValue -ArgsList $DailyRunArgs -Name "--service-account-json"
if ($serviceAccountJson) {
    $sheetArgs += @("--service-account-json", $serviceAccountJson)
}
$oauthClientJson = Get-ArgValue -ArgsList $DailyRunArgs -Name "--oauth-client-json"
if ($oauthClientJson) {
    $sheetArgs += @("--oauth-client-json", $oauthClientJson)
}
$oauthTokenJson = Get-ArgValue -ArgsList $DailyRunArgs -Name "--oauth-token-json"
if ($oauthTokenJson) {
    $sheetArgs += @("--oauth-token-json", $oauthTokenJson)
}

Write-Host "Running daily job-search pipeline with sourcing, evaluation, packets, and local exports..."
$dailyExitCode = Invoke-PythonCommand -Arguments $dailyArgs
if ($dailyExitCode -ne 0) {
    exit $dailyExitCode
}

Write-Host ""
Write-Host "Running live Google Sheet sync as a separate verified step..."
$sheetExitCode = Invoke-PythonCommand -Arguments $sheetArgs
if ($sheetExitCode -ne 0) {
    exit $sheetExitCode
}

Write-Host ""
Write-Host "Checking whether browser follow-up is still pending for this run..."
$browserStatusArgs = @(".\job_search_assistant.py", "browser-follow-up-status", "--latest", "--fail-if-pending")
$browserExitCode = Invoke-PythonCommand -Arguments $browserStatusArgs
if ($browserExitCode -ne 0) {
    Write-Host ""
    Write-Host "Daily workflow is not fully complete yet: CLI search and Sheet sync finished, but browser follow-up is still pending."
    exit $browserExitCode
}

Write-Host ""
Write-Host "Daily workflow completed: CLI search, browser follow-up state, and live Google Sheet sync are all complete."
exit 0
