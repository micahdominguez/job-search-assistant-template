# Google Sheets CRM

The local SQLite database remains the durable source of truth. The `Jobs` tab is the command center for manual review.

If you connect Google Sheets, the workbook is meant to recreate the day-to-day working view of the app in spreadsheet form. It is not just a raw export. It is the tab-based workspace where you review job leads, check fit scores, open job links, track follow-ups, and see packet progress in one place.

## Workbook Tabs

- `Jobs` - master list, one row per opportunity
- `Top Today` - the best active jobs to apply to first
- `Applied` - jobs already applied to or later in the pipeline
- `Follow Ups` - rows with follow-up dates due soon
- `Sector Summary` - sector-level counts and best open role
- `Packets` - condensed packet review view
- `Applications` - application records from SQLite
- `Contacts` - recruiter, hiring manager, referral, and network contacts
- `Correspondence` - outreach and response log
- `Notes` - job notes

## Core Rule

Manual `Status` edits belong in `Jobs`.

The live Google Sheet can be configured so:

- changing `Jobs!Status` updates row color
- `Top Today`, `Applied`, and `Follow Ups` become formula-driven views of `Jobs`
- `Packets!Status` looks up the current job status from `Jobs`

## Create Your Own Workbook

The easiest setup is:

1. Open [sheets.new](https://sheets.new).
2. Name the blank spreadsheet `Job Search CRM` or whatever you prefer.
3. Copy the spreadsheet URL.
4. Open [Google Cloud Console](https://console.cloud.google.com/).
5. Create a Google Cloud project or use one you already control.
6. Enable the Google Sheets API.
7. Open `APIs & Services` -> `OAuth consent screen`.
8. Choose `External` unless you know you need another option.
9. Fill in the required basic app details and save.
10. Open `APIs & Services` -> `Credentials`.
11. Create an `OAuth client ID`.
12. Choose application type `Desktop app`.
13. Download the OAuth client JSON.
14. Put that file at `secrets/google-oauth-client.json`.
15. Install the Google sync dependencies.
16. Run:

```powershell
python .\job_search_assistant.py sync-google-sheets-workbook --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --auth-mode oauth
```

On the first run, the tool will:

- open a browser for Google consent
- write `secrets/google-oauth-token.json`
- create any missing workbook tabs automatically
- push the current local workbook data into the sheet

You do **not** need to pre-create the tabs manually.

If the browser does not open, stop and verify that:

- `secrets/google-oauth-client.json` exists
- the JSON came from a `Desktop app` OAuth client
- the Google Sheets API is enabled in the same project

## Useful Commands

Export the master sheet CSV:

```powershell
python .\job_search_assistant.py export-sheets-csv
```

Export every workbook tab:

```powershell
python .\job_search_assistant.py export-sheets-workbook
```

Sync to your own live workbook:

```powershell
python .\job_search_assistant.py sync-google-sheets-workbook --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --auth-mode oauth
```

Run the local refresh flow without live sync:

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
```

Once your own workbook URL and credentials are working, use the wrapper for the full local-plus-live flow:

```powershell
.\scripts\run_daily_job_search.ps1 --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --sheet-auth-mode oauth
```

## Manual Job Add Workflow

When you find a job yourself, paste it into your analysis workflow first, then import the structured output:

```powershell
Get-Clipboard | python .\job_search_assistant.py import-chatgpt-job
```

After import, tighten the row with your own judgment:

```powershell
python .\job_search_assistant.py update-job --id <JOB_ID> --priority "P1 Apply Today" --sector "Your Sector" --role-lane "Strategic Account Management" --next-action "Generate packet and apply"
```

For strong jobs:

```powershell
.\scripts\generate_packet_and_mirror.ps1 -JobId <JOB_ID>
```

## Google Credentials

For the simplest public-template setup, start with OAuth only.

Store personal credentials in:

- `secrets/google-oauth-client.json`
- `secrets/google-oauth-token.json`
- `secrets/google-service-account.json`

For most new users:

- start with `google-oauth-client.json`
- let the app create `google-oauth-token.json` on first login
- ignore `google-service-account.json` unless you have a specific shared-drive or automation reason

Install Google sync dependencies only if you need Sheets or Drive sync:

```powershell
python -m pip install -r .\requirements-google-drive-sync.txt
```

Official references:

- [Google Sheets API Python quickstart](https://developers.google.com/workspace/sheets/api/quickstart/python)
- [Google Drive API Python quickstart](https://developers.google.com/workspace/drive/api/quickstart/python)
