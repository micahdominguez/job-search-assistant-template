# Job Search Automation CRM Template

This template is for building a personal job search automation CRM that helps find jobs for you and take the busy work out of applying.

In plain English, the system is built to search for job leads that match your background, score them for fit, save them in a spreadsheet-style CRM, and help generate the follow-up materials you need to apply.

You do not need to be a developer to use it. There are still a few technical setup steps, but the goal is to give you a working system for your job search, not to make you learn engineering first.

## What This Template Does

- searches for new job leads tailored to your background
- scores jobs with a fit score so you can focus on stronger matches first
- saves job leads into a spreadsheet-style CRM with links back to the original posting
- tracks your pipeline like a personal CRM
- stores notes, follow-ups, and application history
- generates job packets with job details, application notes, and question support
- generates cover letters
- can sync your tracker to Google Sheets with the same tab-based working view
- can sync packet documents to Google Drive

## What The Main Workflow Looks Like

The main job of this system is:

1. find jobs that fit your background
2. score those jobs so the best ones rise to the top
3. place them into a spreadsheet-style CRM with links and workflow tabs
4. help create packet notes, question answers, and a cover letter
5. reduce the repetitive admin work around job hunting

If you use Google Sheets, the workbook tabs are meant to recreate the working experience of the app in a format that feels easy to review every day.

## Best Way To Use This README

If you are new, follow this order:

1. Get the repo running locally.
2. Add your own basic background information.
3. Test one simple local command.
4. Set up Google Sheets if you want a live spreadsheet.
5. Set up Google Drive only after Sheets works.

That order keeps the setup from getting overwhelming.

## What The Main Files Mean

You do not need to memorize these right away. This is just the simple version:

- `profile.md` and `profile.json`
  These store your background, goals, and contact information.
- `voice_application_style.md` and `voice_application_style.json`
  These help the tool learn how your application writing should sound.
- `packet_writing_training_samples.md`
  This is where you can store good past writing samples.
- `JOB_FINDER_SYSTEM_PROMPT.md`
  This helps shape how the system thinks about good job matches.

If you want the lightest possible start, focus on `profile.md` and `profile.json` first.

## First-Time Setup

1. Install Python 3.11+.
2. Clone the repo.
3. Open PowerShell in the repo folder.
4. If you plan to use Google Sheets or Google Drive, install the Google sync dependencies:

```powershell
python -m pip install -r .\requirements-google-drive-sync.txt
```

5. Open the folder in Codex.
6. Confirm the CLI works:

```powershell
python .\job_search_assistant.py -h
```

## First Personalization Pass

Before worrying about advanced features, do this:

1. Open `profile.md`.
2. Replace the placeholder background with your real background.
3. Open `profile.json`.
4. Replace the placeholder contact info and summary with your own.
5. If you want better writing quality later, update `voice_application_style.md` and `voice_application_style.json`.

That is enough to get started.

You do not need to fully understand every file before your first test.

If you plan to use packet generation heavily later, run this after your first setup pass:

```powershell
.\scripts\find_personalization_hotspots.ps1
```

That script points out remaining example text inside `job_search_assistant.py` that should be adjusted to better match your own background.

## First Local Test

Try these in order:

Show the help menu:

```powershell
python .\job_search_assistant.py -h
```

Evaluate a simple example:

```powershell
python .\job_search_assistant.py evaluate --text "Strategic Account Manager role owning renewals, expansion, executive stakeholders, and enterprise accounts."
```

Show the pipeline:

```powershell
python .\job_search_assistant.py pipeline --min-score 70
```

Run a simple local-only refresh:

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
```

If those work, the local setup is in good shape.

## Basic Usage

Import a job analysis you prepared elsewhere:

```powershell
Get-Clipboard | python .\job_search_assistant.py import-chatgpt-job
```

Generate a packet after you finish personalization:

```powershell
.\scripts\generate_packet_and_mirror.ps1 -JobId 1
```

Generate a local-only packet without Google sync:

```powershell
python .\job_search_assistant.py application-packet --id 1 --skip-google-sync
```

## What You Get In Google Sheets

If you connect Google Sheets, the app can recreate the working tracker experience in a spreadsheet with tabs like:

- `Jobs`
- `Top Today`
- `Applied`
- `Follow Ups`
- `Packets`
- `Applications`
- `Contacts`
- `Correspondence`

That sheet is meant to give you a familiar daily operating view: new leads, fit scores, job links, statuses, follow-ups, and packet progress all in one place.

## Google Setup

The template does not point at any live Google Sheet or Drive folder by default.

If you want the easiest first success, set up Google Sheets first and leave Google Drive for later.

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

## Local-Only Mode

You do not need Google at all to use the core workflow.

Use these when you want to stay local:

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
python .\job_search_assistant.py application-packet --id 1 --skip-google-sync
```

## Recommended Sharing Model

Use this as a **GitHub template repo** if you want other people to start from it.

Recommended flow:

1. Push this folder to a GitHub repo.
2. Mark the repo as a GitHub template.
3. Have each person create their own repo from the template.
4. Each person fills in their own profile, writing style, samples, and secrets locally.

Public is fine once the template is sanitized. If you are still checking for personal traces, keep it private until that review is done.

Do not share:

- `secrets/`
- `data/`
- `exports/`
- `job_packets/`

## Official Google References

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
