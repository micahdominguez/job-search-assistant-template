# Project Runbook

Use this file to start fresh Codex chats quickly and keep the work focused.

## Goal

This repo exists to:

- track and prioritize jobs
- generate stronger application packets
- sync useful outputs into Google Sheets and Google Drive
- improve packet voice over time from real examples

## Source Of Truth

Use this order:

1. `data/job_results.sqlite`
2. `profile.json` and `profile.md`
3. `voice_application_style.json` and `voice_application_style.md`
4. `packet_writing_training_samples.md`
5. `JOB_PACKET_GENERATION_INSTRUCTIONS.md`
6. `application_question_overrides.json`
7. `job_packets/job-*/`
8. `exports/`
9. Google Sheets / Google Docs

## Read Order For A New Chat

Tell a new Codex chat to read only these files first:

1. `PROJECT_RUNBOOK.md`
2. `README.md`
3. `JOB_PACKET_GENERATION_INSTRUCTIONS.md`
4. `profile.md`
5. `voice_application_style.md`
6. `GOOGLE_SHEETS.md`

Only after that should it open:

- `job_search_assistant.py`
- `application_question_overrides.json`
- specific `job_packets/job-*/` folders relevant to the task

## When To Start A New Chat

Start a new chat when:

- you switch topics
- you want to review a specific batch of jobs
- you want to train on one letter lane or answer pattern
- the current chat is long and repeating history

## What Should Be Saved To Files

Save to `profile.*` when it is a true fact about the candidate.

Save to `voice_application_style.*` when it is a reusable writing or tone rule.

Save to `packet_writing_training_samples.md` when it is a raw answer or cover letter worth learning from.

Save to `JOB_PACKET_GENERATION_INSTRUCTIONS.md` when it is a cross-job packet rule.

Save to `application_question_overrides.json` when you have real application questions for a specific job.

Save to `GOOGLE_SHEETS.md` when it affects statuses, tabs, sync behavior, or workflow rules.

Save to `README.md` when a new machine or new collaborator would need it.

## Daily Commands

```powershell
python .\job_search_assistant.py daily-run --skip-sheet-sync
.\scripts\run_daily_job_search.ps1 --spreadsheet-url "https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit" --sheet-auth-mode oauth
python .\job_search_assistant.py pipeline --min-score 90 --limit 20
python .\job_search_assistant.py followups --days 7
Get-Clipboard | python .\job_search_assistant.py import-chatgpt-job
.\scripts\generate_packet_and_mirror.ps1 -JobId <JOB_ID>
```

## New Chat Helper

```powershell
.\scripts\start_new_chat.ps1 -Task "Describe the task here"
```
