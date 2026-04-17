# IRDR Mobile Count

A small GitHub Pages app for weekly IRDR counting on mobile devices.

## What It Does

- Lets the user choose a weekly sample and facility
- Shows count locations in a flashcard-style mobile workflow
- Tracks cases incorrect or missing with `+` and `-`
- Saves notes and progress in the browser with `localStorage`
- Keeps progress local to the device/browser until the count is exported or reset
- Saves results directly into a user-selected local folder on compatible Chromium/Android devices
- Supports Microsoft sign-in plus direct OneDrive upload into a chosen folder path in the signed-in user's OneDrive
- Can hand the generated CSV to Android's native share sheet so the counter can send it to OneDrive or another installed app
- Falls back to downloading the results CSV when direct folder save is unavailable
- Supports install-style behavior with a web manifest and service worker

## Files

- `index.html`: app shell for GitHub Pages
- `assets/styles.css`: mobile-first styling
- `assets/app.js`: app logic and local progress handling
- `assets/onedrive-config.js`: OneDrive upload settings and Microsoft app registration values
- `assets/vendor/msal-browser.min.js`: Microsoft browser auth library used for sign-in
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
- Progress stays on that same device/browser until the count is exported, reset, or local browser data is cleared.
- If that folder has been granted write access, tapping `Finish Count` will try to save the CSV directly there.
- If OneDrive is connected on the setup screen, `Finish & Upload` also sends the CSV to the configured OneDrive folder path.
- On devices that support file sharing from the browser, `Finish & Share` opens the native Android share sheet first so the counter can send the CSV to OneDrive.
- If folder access is not supported or permission is unavailable, the app falls back to downloading the CSV.
- You can also use `Save or Download Results` for a local export or `Share Results` for a share-sheet snapshot during the count.
- The app stores the chosen folder handle in IndexedDB when the browser allows it, but write permission may still need to be re-granted later.

## Enable OneDrive Upload

1. Register a Microsoft Entra application for a single-page app.
2. Add this GitHub Pages URL as a SPA redirect URI:
   `https://christian-ori-ai.github.io/IRDR/`
3. Grant delegated Microsoft Graph permission `Files.ReadWrite`.
4. Open [assets/onedrive-config.js](/c:/Users/T065714/OneDrive - US Foods/CC/Dev/IRDR/IRDR/assets/onedrive-config.js) and:
   set `enabled` to `true`
   replace `clientId` with the app registration's Application (client) ID
   keep the redirect/logout URIs aligned with the GitHub Pages URL unless you intentionally change the published URL
   set `uploadPath` to the OneDrive folder path you want, such as `IRDR/Results`
5. Republish the site.

With the default config, uploaded CSVs land in:
`OneDrive/IRDR/Results`

This uses delegated access in the signed-in user's own drive, which fits OneDrive for Business better than the app-folder scope.

## Publish To GitHub Pages

1. Push the repo to GitHub.
2. In GitHub, open `Settings` -> `Pages`.
3. Set the source to `Deploy from a branch`.
4. Choose the `main` branch and `/ (root)`.
5. Save.

After GitHub Pages publishes, open the site on a phone and optionally install it to the home screen.
