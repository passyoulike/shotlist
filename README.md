# Radi Production Shot List Web App

Google Apps Script web app for the CAHS promotional shoot (Radi Production). It reads and writes the
`CAHS_Batch1 _Shot_List` Google Sheet.

**Open the app:** https://passyoulike.github.io/shotlist/

The GitHub Pages site (`docs/index.html`) shows the live Apps Script app full-screen.
The app itself runs on Google Apps Script, so changes to `Code.gs` / `Index.html` still have to be
deployed there (see below).

## Files

| File | In the Apps Script project | What it does |
|---|---|---|
| `Code.gs` | `Code.gs` | Server side: reads/writes the sheet, saves uploaded photos to Drive |
| `Index.html` | `Index.html.html` (loaded as `Index.html`) | The web page |
| `guide/parts.html` | – | Labelled illustration of the app's parts |
| `docs/index.html` | – | GitHub Pages page that opens the live app |

## Sheets used

- **Shot List**: headings in row 4 (Time, Location, Faculty / CI, Status, Notes, CATEGORY, then one column per shot type). Row 1 photos live in the shot-type cells.
- **PHOTOS**: photos 2–5 for each shot type, plus the **Video Done** checkbox for every photo. Created automatically.
- **NOTES**: the Notes tab (created automatically). Anyone can read; editing needs an admin password from the **Admin** sheet (any cell from row 2 down; row 1 is for labels). Passwords are never stored in this repo.
- **MEDIA**: uploads from the Audio tab (files go to the Drive folder "Radi Production Shot List Media"). Anyone can upload audio/video up to 35 MB; only an unlocked admin can delete.
- **TEAM** and **SL 1** are no longer shown in the app.

## Deploying a change

1. Paste `Code.gs` and `Index.html` into the Apps Script project and save.
2. **Deploy → Manage deployments →** edit (pencil) **→ Version: New version → Deploy**.
   The `/exec` link stays the same.

Photo upload needs Google Drive access: run `doGet` once from the editor and approve the prompt.
