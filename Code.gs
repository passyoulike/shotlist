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
  return { shot: read_('shot'), team: read_('team'), sl1: read_('sl1'), notes: readNotes_(), media: readMedia_() };
}

/* ---------- Audio tab: anyone can upload audio/video; only admins delete ---------- */

const MEDIA_HEADERS = ['ID', 'Title', 'File Name', 'Type', 'Size (MB)', 'Uploaded By', 'Uploaded', 'Drive ID'];
const MAX_MEDIA_MB = 35;

function uploadMedia(info, base64) {
  const mime = String(info && info.mime || '');
  if (!/^(audio|video)\//.test(mime)) throw new Error('Only audio (e.g. MP3) and video files can be uploaded.');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_MEDIA_MB * 1024 * 1024) throw new Error('File is larger than ' + MAX_MEDIA_MB + ' MB.');
  const name = String(info.name || 'upload').slice(0, 200);
  const file = mediaFolder_().createFile(Utilities.newBlob(bytes, mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return withLock_(() => {
    const now = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    mediaSheet_().appendRow([
      Utilities.getUuid().slice(0, 8),
      String(info.title || name).slice(0, 200),
      name,
      mime.indexOf('video/') === 0 ? 'Video' : 'Audio',
      Math.round(bytes.length / 1048576 * 10) / 10,
      String(info.by || '').slice(0, 80),
      now,
      file.getId()
    ]);
    return readMedia_();
  });
}

function deleteMedia(token, id) {
  checkAdmin_(token);
  return withLock_(() => {
    const sh = mediaSheet_();
    if (sh.getLastRow() < 2) return [];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, MEDIA_HEADERS.length).getDisplayValues();
    const i = rows.findIndex(r => r[0] === String(id));
    if (i >= 0) {
      try { DriveApp.getFileById(rows[i][7]).setTrashed(true); } catch (e) { /* already gone */ }
      sh.deleteRow(i + 2);
    }
    return readMedia_();
  });
}

function mediaSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('MEDIA');
  if (!sh) {
    sh = ss.insertSheet('MEDIA');
    sh.getRange(1, 1, 1, MEDIA_HEADERS.length).setValues([MEDIA_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function mediaFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('MEDIA_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* folder was deleted; make a new one */ }
  }
  const folder = DriveApp.createFolder('Radi Production Shot List Media');
  props.setProperty('MEDIA_FOLDER_ID', folder.getId());
  return folder;
}

function readMedia_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('MEDIA');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MEDIA_HEADERS.length).getDisplayValues()
    .filter(r => r[0] && r[7])
    .map(r => ({ id: r[0], title: r[1], name: r[2], kind: r[3], size: r[4], by: r[5], uploaded: r[6], driveId: r[7] }));
}

/* ---------- Notes tab: everyone reads, admins (password) write ---------- */
// The password lives in Project Settings > Script Properties as NOTES_PASSWORD (never in this code).

const NOTE_HEADERS = ['ID', 'Title', 'Note', 'Tag', 'Pinned', 'Updated', 'Created'];
const ADMIN_HOURS = 6;

function adminLogin(password) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adm_fails') || 0);
  if (fails >= 10) throw new Error('Too many wrong tries. Wait 10 minutes and try again.');
  const real = PropertiesService.getScriptProperties().getProperty('NOTES_PASSWORD');
  if (!real) throw new Error('The admin password has not been set up yet.');
  if (String(password || '').trim() !== real) {
    cache.put('adm_fails', String(fails + 1), 600);
    throw new Error('Wrong password.');
  }
  const token = Utilities.getUuid();
  cache.put('adm_' + token, '1', ADMIN_HOURS * 3600);
  return token;
}

function saveNote(token, note) {
  checkAdmin_(token);
  return withLock_(() => {
    const sh = notesSheet_();
    const now = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const title = String(note.title || '').slice(0, 200), body = String(note.body || '').slice(0, 20000);
    const tag = String(note.tag || '').trim().slice(0, 60), pinned = note.pinned ? 'TRUE' : '';
    if (!title.trim() && !body.trim()) throw new Error('Write a title or some text first.');
    const at = note.id ? findNote_(sh, note.id) : 0;
    if (at) {
      sh.getRange(at, 2, 1, 5).setValues([[title, body, tag, pinned, now]]);
    } else {
      sh.appendRow([Utilities.getUuid().slice(0, 8), title, body, tag, pinned, now, now]);
    }
    return readNotes_();
  });
}

function deleteNote(token, id) {
  checkAdmin_(token);
  return withLock_(() => {
    const sh = notesSheet_(), at = findNote_(sh, id);
    if (at) sh.deleteRow(at);
    return readNotes_();
  });
}

function checkAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('adm_' + token)) {
    throw new Error('Admin access has expired. Unlock again.');
  }
}

function notesSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('NOTES');
  if (!sh) {
    sh = ss.insertSheet('NOTES');
    sh.getRange(1, 1, 1, NOTE_HEADERS.length).setValues([NOTE_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findNote_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  const i = ids.findIndex(r => r[0] === String(id));
  return i < 0 ? 0 : i + 2;
}

function readNotes_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NOTES');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, NOTE_HEADERS.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ id: r[0], title: r[1], body: r[2], tag: r[3], pinned: r[4].toUpperCase() === 'TRUE', updated: r[5], created: r[6] }));
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
