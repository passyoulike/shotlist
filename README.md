# CAHS Shot List

Google Apps Script web app for the CAHS promotional shoot. It reads and writes the
`CAHS_Batch1 _Shot_List` Google Sheet.

## Files

| File | In the Apps Script project | What it does |
|---|---|---|
| `Code.gs` | `Code.gs` | Server side: reads/writes the sheet, saves uploaded photos to Drive |
| `Index.html` | `Index.html.html` (loaded as `Index.html`) | The web page |
| `guide/parts.html` | – | Labelled illustration of the app's parts |

## Sheets used

- **Shot List**: headings in row 4 (Time, Location, Faculty / CI, Status, Notes, CATEGORY, then one column per shot type). Row 1 photos live in the shot-type cells.
- **PHOTOS**: photos 2–5 for each shot type, plus the **Video Done** checkbox for every photo. Created automatically.
- **TEAM** and **SL 1**: the Team and SL 1 tabs.

## Deploying a change

1. Paste `Code.gs` and `Index.html` into the Apps Script project and save.
2. **Deploy → Manage deployments →** edit (pencil) **→ Version: New version → Deploy**.
   The `/exec` link stays the same.

Photo upload needs Google Drive access: run `doGet` once from the editor and approve the prompt.
