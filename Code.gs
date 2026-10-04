/**
 * Radi Production Shot List Web App (CAHS promotional shoot)
 * Tabs:  Shot List  ->  sheet "Shot List"  (header row 4: Time | Location | Faculty / CI)
 *        Team       ->  sheet "TEAM"
 *        SL 1       ->  sheet "SL 1"
 * Empty header cells are filled in automatically the first time the app loads.
 */
const SPREADSHEET_ID = '1jxyMIOusVC4mzjhnEOCKgB9Rx_6Dmsv3GrgoDExRfug';

const TABS = {
  shot: { sheet: 'Shot List', headerRow: 4,
          headers: ['Time', 'Location', 'Faculty / CI', 'Status', 'Notes'] },
  team: { sheet: 'TEAM', headerRow: 1,
          headers: ['Name', 'Role', 'Contact No.', 'Email', 'Assignment', 'Notes'] },
  sl1:  { sheet: 'SL 1', headerRow: 1,
          headers: ['Shot #', 'Time', 'Location', 'Shot Description', 'Shot Type',
                    'Camera / Angle', 'Subject / Talent', 'Assigned To', 'Status', 'Notes'] }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index.html') // project file is named "Index.html"
    .setTitle('Radi Production Shot List Web App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- API called from Index.html ---------- */

function getData() {
  return { shot: read_('shot'), team: read_('team'), sl1: read_('sl1') };
}

function updateCell(key, row, col, value) {
  return withLock_(() => {
    const cfg = cfg_(key), sh = sheet_(key);
    if (row <= cfg.headerRow) throw new Error('Cannot edit header rows.');
    let cell = sh.getRange(row, col);
    if (cell.isPartOfMerge()) cell = cell.getMergedRanges()[0].getCell(1, 1);
    cell.setValue(value);
    return true;
  });
}

function setDate(value) {
  return withLock_(() => { sheet_('shot').getRange('B2').setValue(value); return true; });
}

function addRow(key, values) {
  return withLock_(() => {
    const cfg = cfg_(key), sh = sheet_(key);
    const row = Math.max(sh.getLastRow(), cfg.headerRow) + 1;
    const vals = cfg.headers.map((_, i) => (values && values[i] != null ? values[i] : ''));
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    return read_(key);
  });
}

/**
 * Saves an uploaded/pasted photo to Drive.
 * Take 1 goes in the Shot List cell itself; takes 2-5 go in the PHOTOS sheet.
 * Both use =IMAGE(...) so the pictures show in the sheet too.
 */
function uploadPhoto(row, col, take, fileName, base64, mime) {
  const cfg = cfg_('shot');
  if (row <= cfg.headerRow) throw new Error('Cannot add a photo to a header row.');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime || 'image/jpeg', fileName || 'photo.jpg');
  const file = photoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
  const formula = '=IMAGE("' + url + '")';
  return withLock_(() => {
    const sh = sheet_('shot');
    if (take <= 1) {
      sh.getRange(row, col).setFormula(formula);
      return url;
    }
    const ph = photosSheet_(), target = photoRow_(ph, sh, row, col, take);
    ph.getRange(target, PHOTO_COL).setFormula(formula);
    return url;
  });
}

/** Ticks / unticks "Video Done" for one picture (stored in the PHOTOS sheet for every take, including take 1). */
function setVideoDone(row, col, take, done) {
  return withLock_(() => {
    if (row <= cfg_('shot').headerRow) throw new Error('Cannot mark a header row.');
    const ph = photosSheet_(), target = photoRow_(ph, sheet_('shot'), row, col, take);
    ph.getRange(target, DONE_COL).insertCheckboxes().setValue(!!done);
    return true;
  });
}

function clearPhoto(row, col, take) {
  return withLock_(() => {
    if (row <= cfg_('shot').headerRow) throw new Error('Cannot clear a header row.');
    if (take <= 1) sheet_('shot').getRange(row, col).clearContent();
    // Removes the photo for takes 2-5, and the Video Done tick for every take.
    const ph = photosSheet_(), at = findPhoto_(ph, row, col, take);
    if (at) ph.deleteRow(at);
    return true;
  });
}

function deleteRow(key, row) {
  return withLock_(() => {
    const cfg = cfg_(key), sh = sheet_(key);
    if (row <= cfg.headerRow || row > sh.getLastRow()) throw new Error('Invalid row.');
    sh.deleteRow(row);
    if (key === 'shot') shiftPhotos_(row);
    return read_(key);
  });
}

/* ---------- PHOTOS sheet (photo rows 2-5 under each time slot) ---------- */

const PHOTO_HEADERS = ['Slot Row', 'Time', 'Location', 'Shot Type', 'Column', 'Take', 'Photo', 'Video Done'];
const PHOTO_COL = 7, DONE_COL = 8;

function photosSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('PHOTOS');
  if (!sh) {
    sh = ss.insertSheet('PHOTOS');
    sh.setFrozenRows(1);
  }
  const head = sh.getRange(1, 1, 1, PHOTO_HEADERS.length);
  if (head.getDisplayValues()[0].join('|') !== PHOTO_HEADERS.join('|')) head.setValues([PHOTO_HEADERS]).setFontWeight('bold');
  return sh;
}

// Finds (or adds) the PHOTOS row for one picture and keeps its description columns up to date.
function photoRow_(ph, sh, row, col, take) {
  const cfg = cfg_('shot');
  const info = [row, sh.getRange(row, 1).getDisplayValue(), sh.getRange(row, 2).getDisplayValue(),
                sh.getRange(cfg.headerRow, col).getDisplayValue(), col, take];
  const target = findPhoto_(ph, row, col, take) || ph.getLastRow() + 1;
  ph.getRange(target, 1, 1, info.length).setValues([info]);
  return target;
}

function readPhotos_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName('PHOTOS');
  if (!sh || sh.getLastRow() < 2) return [];
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 1, n, DONE_COL).getValues(), formulas = sh.getRange(2, PHOTO_COL, n, 1).getFormulas();
  const out = [];
  vals.forEach((v, i) => {
    const p = v[PHOTO_COL - 1];
    const url = imageUrl_(p, formulas[i][0]) || (/^https?:\/\//.test(String(p)) ? String(p) : '');
    const done = v[DONE_COL - 1] === true;
    if (url || done) out.push({ slot: Number(v[0]), col: Number(v[4]), take: Number(v[5]), url: url, done: done });
  });
  return out;
}

function findPhoto_(ph, row, col, take) {
  if (ph.getLastRow() < 2) return 0;
  const vals = ph.getRange(2, 1, ph.getLastRow() - 1, 6).getValues();
  const i = vals.findIndex(v => Number(v[0]) === row && Number(v[4]) === col && Number(v[5]) === take);
  return i < 0 ? 0 : i + 2;
}

// Keep PHOTOS pointing at the right time slot after a Shot List row is deleted.
function shiftPhotos_(deletedRow) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PHOTOS');
  if (!sh || sh.getLastRow() < 2) return;
  for (let r = sh.getLastRow(); r >= 2; r--) {
    const slot = Number(sh.getRange(r, 1).getValue());
    if (slot === deletedRow) sh.deleteRow(r);
    else if (slot > deletedRow) sh.getRange(r, 1).setValue(slot - 1);
  }
}

/* ---------- helpers ---------- */

function cfg_(key) {
  const cfg = TABS[key];
  if (!cfg) throw new Error('Unknown tab: ' + key);
  return cfg;
}

function sheet_(key) {
  const cfg = cfg_(key), ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(cfg.sheet)
    || ss.getSheets().find(s => s.getName().trim().toLowerCase() === cfg.sheet.toLowerCase())
    || ss.insertSheet(cfg.sheet);
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* folder was deleted; make a new one */ }
  }
  const folder = DriveApp.createFolder('CAHS Shot List Photos');
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

/** Picture URL for a cell: =IMAGE("url") formula, or an image placed in the cell. */
function imageUrl_(raw, formula) {
  const m = /^=\s*IMAGE\(\s*"([^"]+)"/i.exec(formula || '');
  if (m) return m[1];
  if (raw && typeof raw === 'object' && raw.valueType === SpreadsheetApp.ValueType.IMAGE) {
    try { return raw.getContentUrl() || ''; } catch (e) { return ''; }
  }
  return '';
}

function ensureHeaders_(sh, cfg) {
  const range = sh.getRange(cfg.headerRow, 1, 1, cfg.headers.length);
  const cur = range.getDisplayValues()[0];
  let changed = false;
  const out = cur.map((v, i) => {
    if (String(v).trim()) return v;
    changed = true;
    return cfg.headers[i];
  });
  if (changed) range.setValues([out]).setFontWeight('bold');
  return out;
}

function read_(key) {
  const cfg = cfg_(key), sh = sheet_(key);
  ensureHeaders_(sh, cfg);
  // Include any extra columns added in the sheet (e.g. CATEGORY and the photo columns).
  const ncol = Math.max(cfg.headers.length, sh.getLastColumn());
  const headers = sh.getRange(cfg.headerRow, 1, 1, ncol).getDisplayValues()[0];
  const res = { key: key, sheet: sh.getName(), headers: headers, rows: [] };
  if (key === 'shot') {
    res.title = sh.getRange('A1').getDisplayValue();
    res.date = sh.getRange('B2').getDisplayValue();
    res.extra = readPhotos_();
  }

  const last = sh.getLastRow();
  if (last <= cfg.headerRow) return res;
  const n = last - cfg.headerRow;
  const rng = sh.getRange(cfg.headerRow + 1, 1, n, ncol);
  const vals = rng.getDisplayValues();
  const locked = vals.map(r => r.map(() => false));

  // Merged cells: show the merged value on every row, but only the first cell is editable.
  rng.getMergedRanges().forEach(m => {
    const r0 = m.getRow() - cfg.headerRow - 1, c0 = m.getColumn() - 1, v = m.getDisplayValue();
    for (let i = 0; i < m.getNumRows(); i++) {
      for (let j = 0; j < m.getNumColumns(); j++) {
        const r = r0 + i, c = c0 + j;
        if (r < 0 || r >= n || c < 0 || c >= ncol) continue;
        vals[r][c] = v;
        if (i || j) locked[r][c] = true;
      }
    }
  });

  const raw = rng.getValues(), formulas = rng.getFormulas();
  vals.forEach((cells, i) => {
    const img = cells.map((_, j) => imageUrl_(raw[i][j], formulas[i][j]));
    if (cells.every(v => !String(v).trim()) && img.every(u => !u)) return;
    res.rows.push({ row: cfg.headerRow + 1 + i, cells: cells, locked: locked[i], img: img });
  });
  return res;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
