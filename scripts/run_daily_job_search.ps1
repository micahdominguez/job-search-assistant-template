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

Write-Host "Running daily job-search refresh with explicit local export step..."
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
Write-Host "Daily workflow completed: local refresh succeeded and live Google Sheet sync completed."
exit 0
