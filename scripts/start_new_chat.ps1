param(
    [string]$Task = "Describe the specific task here.",
    [string[]]$JobIds = @(),
    [string[]]$Files = @()
)

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$jobLine = ""
if ($JobIds.Count -gt 0) {
    $jobLine = "Focus job ids: " + ($JobIds -join ", ") + "."
}

$fileLine = ""
if ($Files.Count -gt 0) {
    $fileLine = "Focus files: " + ($Files -join ", ") + "."
}

$prompt = @"
Read PROJECT_RUNBOOK.md, README.md, JOB_PACKET_GENERATION_INSTRUCTIONS.md, profile.md, voice_application_style.md, and GOOGLE_SHEETS.md first.

Workspace:
$repoRoot

Task:
$Task

$jobLine
$fileLine

Only inspect files directly relevant to this task.
Do not reread unrelated old packet folders unless needed.
Keep the work focused and efficient.
"@.Trim()

try {
    Set-Clipboard -Value $prompt
    $clipboardMessage = "Prompt copied to clipboard."
}
catch {
    $clipboardMessage = "Could not copy to clipboard automatically, but the prompt is printed below."
}

Write-Host ""
Write-Host "New chat starter"
Write-Host "================"
Write-Host $clipboardMessage
Write-Host ""
Write-Host $prompt
Write-Host ""
Write-Host "Example:"
Write-Host '.\scripts\start_new_chat.ps1 -Task "Review job 56 packet quality" -JobIds 56'
