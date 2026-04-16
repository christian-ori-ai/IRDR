# IRDR Mobile Count

A small GitHub Pages app for weekly IRDR counting on mobile devices.

## What It Does

- Lets the user choose a weekly sample and facility
- Shows count locations in a flashcard-style mobile workflow
- Tracks cases incorrect or missing with `+` and `-`
- Saves notes and progress in the browser with `localStorage`
- Exports the entered results to CSV from the phone
- Supports install-style behavior with a web manifest and service worker

## Files

- `index.html`: app shell for GitHub Pages
- `assets/styles.css`: mobile-first styling
- `assets/app.js`: app logic and local progress handling
- `data/samples.json`: week and facility sample data used by the app
- `scripts/export_irdr_sample.py`: converts an IRDR workbook into `samples.json`

## Update The Weekly Sample

From the repo root:

```powershell
python scripts/export_irdr_sample.py "..\IRDR Sample 2026-04-16 Excluding 9300 9181.xlsx"
```

That command updates `data/samples.json` so the site uses the newest workbook data.

## Publish To GitHub Pages

1. Push the repo to GitHub.
2. In GitHub, open `Settings` -> `Pages`.
3. Set the source to `Deploy from a branch`.
4. Choose the `main` branch and `/ (root)`.
5. Save.

After GitHub Pages publishes, open the site on a phone and optionally install it to the home screen.
