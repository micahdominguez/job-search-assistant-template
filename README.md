# Personal Job Search Assistant Template

This is a sanitized starter repo for running the local-first job search workflow with your own profile, packet voice, CRM data, and Google setup.

## What This Template Includes

- Python CLI for job evaluation, CRM tracking, exports, and optional Google sync
- Starter profile and voice files that you should replace with your own data
- Google Sheets command-center workflow
- Packet-writing instructions and sample scaffolding
- A personalization checklist so you can safely adapt it before real usage

## What You Should Customize First

Before using this seriously, replace:

- `profile.md`
- `profile.json`
- `voice_application_style.md`
- `voice_application_style.json`
- `packet_writing_training_samples.md`
- `JOB_FINDER_SYSTEM_PROMPT.md`

If you want to use application packet generation heavily, also run:

```powershell
.\scripts\find_personalization_hotspots.ps1
```

That will show remaining candidate-tuned fallback text inside `job_search_assistant.py` that should be reviewed for your background.

## Recommended Sharing Model

Use this as a **GitHub template repo**.

Recommended flow:

1. Push this folder to a GitHub repo.
2. Mark the repo as a GitHub template.
3. Have each person create their own repo from the template.
4. Each person fills in their own profile, voice, samples, and secrets locally.

Public is fine once the template is sanitized. If you are still checking for personal traces, keep it private until that review is done.

Do not share:

- `secrets/`
- `data/`
- `exports/`
- `job_packets/`

## First-Time Setup

1. Install Python 3.11+.
2. Clone the repo.
3. Open PowerShell in the repo folder.
4. Install the Google sync dependencies if you plan to use Google Sheets or Google Drive:

```powershell
python -m pip install -r .\requirements-google-drive-sync.txt
```

5. Open the folder in Codex.
6. Confirm the CLI works:

```powershell
python .\job_search_assistant.py -h
```

7. Fill in the profile and voice files with your real background.
8. Read the checklist in `PERSONALIZATION_CHECKLIST.md`.

## Basic Usage

Evaluate pasted text:

```powershell
python .\job_search_assistant.py evaluate --text "Strategic Account Manager role owning renewals, expansion, executive stakeholders, and enterprise accounts."
```

Import a job analysis you prepared elsewhere:

```powershell
Get-Clipboard | python .\job_search_assistant.py import-chatgpt-job
```

Show the pipeline:

```powershell
python .\job_search_assistant.py pipeline --min-score 70
```

Generate a packet after you finish personalization:

```powershell
.\scripts\generate_packet_and_mirror.ps1 -JobId 1
```

Generate a local-only packet without Google sync:

```powershell
python .\job_search_assistant.py application-packet --id 1 --skip-google-sync
```

## Google Setup

The template does not point at any live Google Sheet or Drive folder by default.

If you want the fastest first success, set up Google Sheets first and leave Google Drive for later.

### Google Sheets OAuth Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or choose an existing one you control.
3. Enable the Google Sheets API for that project.
4. If you want Google Drive packet sync later, enable the Google Drive API too.
5. Open `APIs & Services` -> `OAuth consent screen`.
6. Choose `External` unless you have a specific reason to use `Internal`.
7. Fill in the basic app name and email fields, then save.
8. Open `APIs & Services` -> `Credentials`.
9. Choose `Create Credentials` -> `OAuth client ID`.
10. For application type, choose `Desktop app`.
11. Download the JSON file.
12. Put that file at `secrets/google-oauth-client.json`.

The repo expects this exact path:

```text
secrets/google-oauth-client.json
```

If `secrets\` does not exist yet, create it first.

### First Live Sheet Sync

1. Create a blank Google Sheet at [sheets.new](https://sheets.new).
2. Name it something like `Job Search CRM`.
3. Copy the spreadsheet URL from the browser.
4. Run this command with your own sheet URL:

```powershell
python .\job_search_assistant.py sync-google-sheets-workbook --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --auth-mode oauth
```

What should happen on the first run:

- a browser window opens for Google login and consent
- Google asks you to approve access
- the repo writes `secrets/google-oauth-token.json`
- the tool creates the workbook tabs for you
- the tool pushes the current local workbook data into the sheet

You do not need to pre-create tabs manually.

### After Google Sheets Works

Once that first sync succeeds, these files should exist:

- `secrets/google-oauth-client.json`
- `secrets/google-oauth-token.json`

You can then use the daily wrapper:

```powershell
.\scripts\run_daily_job_search.ps1 --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --sheet-auth-mode oauth
```

If you only want local tracking, keep using:

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
```

### Optional Google Drive Packet Sync

Only do this after Sheets OAuth is already working.

1. Create or choose one Google Drive folder for packet docs.
2. Create or choose one Google Drive folder for cover-letter docs.
3. Copy both folder URLs from the browser.
4. Run:

```powershell
python .\job_search_assistant.py sync-google-drive-docs --packet-folder-url "https://drive.google.com/drive/folders/<YOUR_PACKET_FOLDER_ID>" --cover-letter-folder-url "https://drive.google.com/drive/folders/<YOUR_COVER_LETTER_FOLDER_ID>" --auth-mode oauth
```

After that succeeds, use:

```powershell
.\scripts\generate_packet_and_mirror.ps1 -JobId <JOB_ID>
```

### Common New-User Checks

- `python .\job_search_assistant.py -h` works
- `secrets/google-oauth-client.json` exists
- the first sheet sync opens a browser
- `secrets/google-oauth-token.json` appears after consent
- the target Google Sheet gets tabs created automatically

### Local-Only Mode

You do not need Google at all to use the core workflow.

Use these when you want to stay local:

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
python .\job_search_assistant.py application-packet --id 1 --skip-google-sync
```

Official Google setup references:

- [Google Sheets API Python quickstart](https://developers.google.com/workspace/sheets/api/quickstart/python)
- [Google Drive API Python quickstart](https://developers.google.com/workspace/drive/api/quickstart/python)

## GitHub Push

If you are starting fresh:

```powershell
git init
git add .
git commit -m "Initial template commit"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
```
