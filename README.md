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
3. Open the folder in Codex.
4. Confirm the CLI works:

```powershell
python .\job_search_assistant.py -h
```

5. Fill in the profile and voice files with your real background.
6. Read the checklist in `PERSONALIZATION_CHECKLIST.md`.

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

If you want a live Google Sheet command center:

1. Create a blank Google Sheet at [sheets.new](https://sheets.new) and name it something like `Job Search CRM`.
2. Copy the spreadsheet URL.
3. In your own Google Cloud project, enable the Google Sheets API.
4. If you also want Google Drive packet sync later, enable the Google Drive API too.
5. Configure OAuth consent and create a Desktop app OAuth client.
6. Save the downloaded OAuth client JSON to `secrets/google-oauth-client.json`.
7. Install dependencies:

```powershell
python -m pip install -r .\requirements-google-drive-sync.txt
```

8. Run the first live sheet sync with your own spreadsheet URL:

```powershell
python .\job_search_assistant.py sync-google-sheets-workbook --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --auth-mode oauth
```

The first OAuth run opens a browser, asks you to approve access, writes `secrets/google-oauth-token.json`, and creates the missing workbook tabs automatically inside the blank spreadsheet.

9. If you also want Google Drive packet and cover-letter sync, pass your own Drive folder URLs explicitly:

```powershell
python .\job_search_assistant.py sync-google-drive-docs --packet-folder-url "https://drive.google.com/drive/folders/<YOUR_PACKET_FOLDER_ID>" --cover-letter-folder-url "https://drive.google.com/drive/folders/<YOUR_COVER_LETTER_FOLDER_ID>" --auth-mode oauth
```

10. After that is working, you can use the wrapper for the full daily flow:

```powershell
.\scripts\run_daily_job_search.ps1 --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --sheet-auth-mode oauth
```

If you are staying local-only, keep using `python .\job_search_assistant.py daily-run --skip-sheet-sync`.

For packet creation after Drive sync is configured, use:

```powershell
.\scripts\generate_packet_and_mirror.ps1 -JobId <JOB_ID>
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
