# Results Folder

Place exported IRDR results files in this folder.

## Current Workflow

1. On a compatible Chromium-based device, the user taps `Choose Results Folder`.
2. The user selects this local `Results` folder from the device filesystem.
3. When the counter taps `Finish Count` or `Finish & Share`, the app tries to write the CSV directly here.
4. On share-capable Android browsers, the app can also open the native share sheet so the same CSV can be sent to OneDrive or another installed app.
5. If folder save is unavailable or permission is denied, the app falls back to downloading the CSV instead.

## File Format

Each results CSV includes:

- company
- week
- facility
- source sample file
- export timestamp
- facility population
- sample size
- completed locations
- defect locations
- total variance cases
- per-location count results and notes
- exporting counter name and last-updated details for traceability

The direct-save path depends on the device browser supporting the File System Access API and the user granting write permission to this folder.
