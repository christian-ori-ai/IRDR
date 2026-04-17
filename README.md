# IRDR Mobile Count

A small GitHub Pages app for weekly IRDR counting on mobile devices.

## What It Does

- Lets the user choose a weekly sample and facility
- Shows count locations in a flashcard-style mobile workflow
- Tracks cases incorrect or missing with `+` and `-`
- Saves notes and progress in the browser with `localStorage`
- Supports shared count coordination through the selected `Results` folder so one device can pause and another can resume
- Saves results directly into a user-selected local folder on compatible Chromium/Android devices
- Can hand the generated CSV to Android's native share sheet so the counter can send it to OneDrive or another installed app
- Falls back to downloading the results CSV when direct folder save is unavailable
- Supports install-style behavior with a web manifest and service worker

## Files

- `index.html`: app shell for GitHub Pages
- `assets/styles.css`: mobile-first styling
- `assets/app.js`: app logic, shared count coordination, and local progress handling
- `data/samples.json`: week and facility sample data used by the app
- `Results/`: place exported count-result CSV files here after they are collected from phones
- `scripts/export_irdr_sample.py`: converts an IRDR workbook into `samples.json`

## Update The Weekly Sample

From the repo root:

```powershell
python scripts/export_irdr_sample.py "..\IRDR Sample 2026-04-16 Excluding 9300 9181.xlsx"
```

That command updates `data/samples.json` so the site uses the newest workbook data.

## Count Results

- While a count is in progress, results are stored on the device in the browser.
- On compatible Chromium browsers, the user can tap `Choose Results Folder` and select the device's local `IRDR/Results` folder.
- Selecting that folder also turns on shared coordination for the chosen week and facility by storing lightweight runtime files in `Results/.irdr-runtime/`.
- Leaving the count screen pauses the shared claim, keeps the latest progress in the shared runtime state, and lets another counter resume or take over if needed.
- If another device already has the count open, the app will show that status and ask before taking over the shared lock.
- If that folder has been granted write access, tapping `Finish Count` will try to save the CSV directly there.
- On devices that support file sharing from the browser, `Finish & Share` opens the native Android share sheet first so the counter can send the CSV to OneDrive.
- If folder access is not supported or permission is unavailable, the app falls back to downloading the CSV.
- You can also use `Save or Download Results` for a local export or `Share Results` for a share-sheet snapshot during the count.
- The app stores the chosen folder handle in IndexedDB when the browser allows it, but write permission may still need to be re-granted later.

## Publish To GitHub Pages

1. Push the repo to GitHub.
2. In GitHub, open `Settings` -> `Pages`.
3. Set the source to `Deploy from a branch`.
4. Choose the `main` branch and `/ (root)`.
5. Save.

After GitHub Pages publishes, open the site on a phone and optionally install it to the home screen.
