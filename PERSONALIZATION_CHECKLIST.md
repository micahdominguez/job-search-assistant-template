# Personalization Checklist

## Before Real Use

1. Replace `profile.md` and `profile.json` with the candidate's real background.
2. Replace `voice_application_style.md` and `voice_application_style.json` with the candidate's real voice.
3. Add real writing samples to `packet_writing_training_samples.md`.
4. Review `JOB_FINDER_SYSTEM_PROMPT.md` and `JOB_PACKET_GENERATION_INSTRUCTIONS.md`.
5. Keep the repo private unless the candidate explicitly wants it public.

## Packet Engine Audit

This template started from a real working repo and some packet fallback text is still candidate-tuned.

Run:

```powershell
.\scripts\find_personalization_hotspots.ps1
```

Then refit the remaining references before trusting packet generation for real applications.

## Google Setup

Only if needed:

1. Put your own credential files in `secrets/`.
2. Install Google sync dependencies.
3. Pass your own spreadsheet and folder URLs at the command line.
4. Use `.\scripts\run_daily_job_search.ps1` only after those values work for your environment.

## Safe Git Hygiene

Do not commit:

- `secrets/`
- `data/`
- `exports/`
- `job_packets/`
