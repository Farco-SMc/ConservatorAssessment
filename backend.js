/**
 * FARCO — Web App backend (Sheets + Drive)
 *
 * This script implements the server‑side logic for the FARCO conservator
 * assessment web app. It writes batches, items, selections and photos
 * into a Google Sheet database, stores photos in Drive, and exposes a
 * simple web interface via doGet().
 *
 * Features:
 *  - Self‑hosted font loading via Drive to embed fonts in the HTML.
 *  - Automatic creation and repair of database sheets with header alignment.
 *  - Support for replacement and reupholstery flags on selections.
 *  - Support for additional notes (incident and other notes) on items.
 */

const TZ = Session.getScriptTimeZone();

// —— DRIVE ROOT ————————————————————————————————
// Folder ID where client photo folders will be created.
const PHOTOS_ROOT_ID = '1OE4NogQ-I1fSFEETsqv_YuQT5qyC8yw1';

// Public/Drive logo file ID used to embed the brand logo as a data URL.
const LOGO_FILE_ID = '1Z3I__UwcTHAcmQ0tBz6cJbRzOPuYROy9';  // TODO: replace with your logo file ID

// —— SPREADSHEET TARGET ————————————————————————
// If you have a dedicated database spreadsheet, set its ID here. Leave as an
// empty string to use the active spreadsheet (for container‑bound scripts).
const SS_ID = ('17cxrzag_x5RhOULB6jF91Q6A0kcgrFyM5K8zYKsjhH0');

// —— SHEET NAMES ————————————————————————————————
const SH_BATCHES = 'Batches';
const SH_ITEMS   = 'Items';
const SH_SELECT  = 'Selections';
const SH_PHOTOS  = 'Photos';

// Canonical header definitions used when creating or repairing sheets.
// These arrays can be extended without breaking existing sheets; missing
// columns will be appended automatically by ensureSheet_().
const HEADERS = {
  [SH_BATCHES]: ['BatchID','Client','StartDate','Assessor','Status'],
  [SH_ITEMS]: [
    'ItemID','BatchID','Code','PerilType','ImpactLevel','Type','Title','Artist','Material','Date',
    'Dimensions','Features','HistoricIssuesPresent','HistoricCode','HistoricNotes',
    'TreatmentTimeMinutes','TreatmentTimeUnit','TreatmentTime',
    'AdditionalNotes','OtherNotes','Notes','CreatedAt'
  ],
  [SH_SELECT]: [
    'SelectionID','ItemID','OptionCode','Severity','ExtentPercent','Location',
    'ItemType','UseSeverity','SeverityWord','BasePhrase','IsLocalized',
    'LocPart','ExtPart','CondLine',
    'GlobalFrag','LocalText','LocalWhere',
    'NeedsReplacement','NeedsReupholstery'
  ],
  [SH_PHOTOS]: ['PhotoID','ItemID','Image','Caption','TakenAt','Lat','Lng','Uploader']
};

// ========== WEB ENDPOINT + FONTS ===========================================

/**
 * Entry point for GET requests. Renders the 'index' HTML template and
 * injects font data URLs for self‑hosted fonts stored in Drive.
 */
function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.fonts = getFontData_();
  t.logoDataUrl = getLogoDataUrl_();
  const page = t.evaluate();
  page.setTitle('FARCO — Conservator assessment');
  page.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return page;
}

/**
 * Helper to convert Drive font files into data: URLs for embedding.
 * Replace the values in the ids map with your actual font file IDs or
 * shared URLs. Only non‑empty entries will be processed.
 */
function getFontData_() {
  const ids = {
    gillLight:   '',   // e.g. '1YUu6YPSWh87doniA_jLeyXCY78thYr2r'
    gillRegular: '',   // e.g. '1eIuMV6mIUTFK3JBD6psJrDrkszJIGFw'
    gillMedium:  '',   // optional: weight 600
    gillBold:    '',   // optional: weight 700
    gillItalic:  '',   // optional: italic style
    begum:       ''    // e.g. '1gfEfjzozzvJWJgYP-Er-zG1ZJ8Nc0FpxDH'
  };
  function toId(val) {
    if (!val) return null;
    const s = String(val);
    const m = s.match(/[-\w]{25,}/);
    return m ? m[0] : null;
  }
  function mimeFor(name) {
    const n = String(name).toLowerCase();
    if (n.endsWith('.woff2')) return 'font/woff2';
    if (n.endsWith('.woff'))  return 'font/woff';
    if (n.endsWith('.otf'))   return 'font/otf';
    if (n.endsWith('.ttf'))   return 'font/ttf';
    return 'application/octet-stream';
  }
  const out = {};
  for (const key in ids) {
    const id = toId(ids[key]);
    if (!id) continue;
    const file = DriveApp.getFileById(id);
    const mime = mimeFor(file.getName());
    const b64  = Utilities.base64Encode(file.getBlob().getBytes());
    out[key] = 'data:' + mime + ';base64,' + b64;
  }
  return out;
}

/**
 * Build a data URL for the brand logo from Drive so the HTML can embed it
 * without relying on Drive sharing/permissions.
 * If you prefer, store the Drive file ID in LOGO_FILE_ID above.
 */
function getLogoDataUrl_() {
  try {
    if (!LOGO_FILE_ID) return '';
    var file = DriveApp.getFileById(LOGO_FILE_ID);
    var blob = file.getBlob();
    var mime = blob.getContentType();
    var b64  = Utilities.base64Encode(blob.getBytes());
    return 'data:' + mime + ';base64,' + b64;
  } catch (e) {
    console.error('getLogoDataUrl_ failed', e);
    return ''; // HTML will hide the image if empty/invalid
  }
}


// ========== SPREADSHEET HELPERS ============================================

/**
 * Obtain a Spreadsheet instance. If SS_ID is defined (non‑empty), the
 * corresponding spreadsheet will be opened by ID. Otherwise the active
 * spreadsheet (bound to this script) will be returned. We do not cache
 * the spreadsheet because the execution environment can vary.
 */
function getSpreadsheet_() {
  if (SS_ID && String(SS_ID).trim()) {
    return SpreadsheetApp.openById(SS_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Ensure a sheet with the given name exists and has the specified header.
 * If the sheet does not exist it will be created. Missing header columns
 * will be appended to the existing header row.
 *
 * @param {string} name The sheet name.
 * @param {string[]} header The desired header row. If provided, missing
 *        columns will be appended.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet instance.
 */
function ensureSheet_(name, header) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header && header.length) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    }
    return sh;
  }
  if (header && header.length) {
    const current = getHeaders_(sh);
    const missing = header.filter(h => current.indexOf(h) === -1);
    if (missing.length) {
      sh.insertColumnsAfter(current.length, missing.length);
      sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sh;
}

/**
 * Return the header row of a sheet as an array of strings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The sheet.
 * @returns {string[]} The header row.
 */
function getHeaders_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

/**
 * Find the next free row based on a key header. Works around ARRAYFORMULA rows
 * by looking at the display values of the key column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The sheet.
 * @param {string} keyHeader The header name whose column is used to find free rows.
 * @returns {number} The 1‑based row index for appending a new row.
 */
function nextFreeRowByHeader_(sh, keyHeader) {
  const headers = getHeaders_(sh);
  const idx = headers.indexOf(keyHeader);
  if (idx === -1) throw new Error('Key header not found: ' + keyHeader);
  const col = idx + 1;
  const totalRows = Math.max(0, sh.getMaxRows() - 1);
  if (totalRows === 0) return 2;
  const colVals = sh.getRange(2, col, totalRows, 1).getDisplayValues();
  for (let i = colVals.length - 1; i >= 0; i--) {
    if (colVals[i][0] !== '') {
      return i + 2 + 1;
    }
  }
  return 2;
}

/**
 * Append a row of values to the sheet identified by name. The row is built
 * according to the sheet's headers. Unknown keys in obj will be ignored and
 * missing keys will result in empty cells.
 *
 * @param {string} name The sheet name.
 * @param {string} keyHeader The header used to determine the next free row.
 * @param {Object} obj An object mapping header names to values.
 */
function appendRowAtKey_(name, keyHeader, obj) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  const headers = getHeaders_(sh);
  const r = nextFreeRowByHeader_(sh, keyHeader);
  // Only set provided keys to avoid overwriting formula columns
  Object.keys(obj || {}).forEach(function(h){
    const col = headers.indexOf(h);
    if (col !== -1) sh.getRange(r, col+1).setValue(obj[h]);
  });
}

/**
 * Read all data rows from a sheet into an array of objects keyed by headers.
 *
 * @param {string} name The sheet name.
 * @returns {Object[]} An array of row objects.
 */
function getAll_(name) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length <= 1) return [];
  const head = v[0];
  return v.slice(1).map(row => {
    const obj = {};
    head.forEach((key, i) => {
      obj[key] = row[i];
    });
    return obj;
  });
}

/**
 * Compute the next code for an item within a batch. Codes are of the form
 * I001, I002, etc., based on existing items for the given batch ID.
 *
 * @param {string} batchId The batch ID.
 * @returns {string} The next item code.
 */
function nextCodeForBatch_(batchId) {
  const items = getAll_(SH_ITEMS).filter(r => r.BatchID === batchId);
  const n = items.length + 1;
  return 'I' + String(n).padStart(3, '0');
}

// ========== NOTES BUILDER ====================================================

/**
 * Build automatic notes based on selected condition codes and other flags.
 *
 * @param {Object} payload The item payload.
 * @returns {string} The notes text.
 */
function buildNotes_(payload) {
  const codes = new Set((payload.selections || []).map(s => s.code));
  const out = [];
  if (payload.historicYes) {
    out.push('Historic issues, including age, wear and use issues will not be addressed as part of the claim');
    out.push('Historic issues can be addressed outside the claim on a private client basis should this be of interest to the client.');
  }
  if (codes.size) {
    out.push('Some minor visible difference may remain in impacted area following treatment');
  }
  const structural = [
    'STRUCT_DAMAGED','JOINTS_LOOSE','VENEER_LIFTING','VENEER_LOSS',
    'PNT_LIFTING','PNT_LOSS','PNT_TEAR','WOP_MEDIA_LOSS','WOP_HINGE_FAIL',
    'FRM_STRUCT','GILT_ORNAMENT_CRACK','OBJ_CORROSION','OBJ_OXIDATION',
    'CERAMIC_CRACK','CERAMIC_CHIP','CERAMIC_BREAK','TXT_TEAR',
    'UPH_FABRIC_DAMAGE','RUG_COLOUR_RUN'
  ];
  if (structural.some(c => codes.has(c))) {
    out.push('Underlying material instability will remain following treatment');
  }
  const flattening = ['WOP_COCKLING','PNT_DEFORMATION'];
  if (flattening.some(c => codes.has(c))) {
    out.push('Given the materials we may face limitation in flattening the artwork, we will take this to a safe level');
  }
  return out.join('\\n');
}

function getNoteFromNotesRules_(code){
  code = String(code || '').trim();
  if (!code) return '';
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('NotesRules');
  if (!sh) return '';
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return '';
  // Try to detect header: if first row contains 'Code'/'NoteText', skip it
  let startRow = 1;
  const h0 = v[0].map(String);
  const twoCol = v[0].length >= 2;
  // We will read first two columns as (Code, Body)
  for (let i = 1; i < v.length; i++){
    const a = String(v[i][0] || '').trim();
    const b = String(v[i][1] || '').trim();
    if (!a) continue;
    if (a.toLowerCase() === code.toLowerCase()) return b;
  }
  return '';
}

// ========== DRIVE HELPERS ====================================================

/**
 * Ensure a client folder exists under the root photo folder. If a folder
 * with the client's name already exists it will be returned; otherwise a
 * new folder will be created.
 *
 * @param {string} rootId The root folder ID.
 * @param {string} client The client name.
 * @returns {GoogleAppsScript.Drive.Folder} The client folder.
 */
function ensureClientFolder_(rootId, client) {
  const root = DriveApp.getFolderById(rootId);
  const name = (client || 'Unnamed Client').trim();
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

// ========== API: CREATE BATCH ===============================================

/**
 * Create a new batch and associated photo folder. Returns the batch ID and
 * folder ID to the client.
 *
 * @param {string} client The client name.
 * @param {string} assessorName The assessor's name.
 * @returns {Object} An object containing batchId and clientFolderId.
 */
function newBatch(client, assessorName) {
  if (!client) throw new Error('Client is required');
  if (!assessorName) throw new Error('Assessor is required');
  const now = new Date();
  const batchId = 'BATCH-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss-SSS');
  const clientFolder = ensureClientFolder_(PHOTOS_ROOT_ID, client);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CLIENT_' + batchId, client);
  props.setProperty('PHOTOS_' + batchId, clientFolder.getId());
  ensureSheet_(SH_BATCHES, HEADERS[SH_BATCHES]);
  appendRowAtKey_(SH_BATCHES, 'BatchID', {
    BatchID: batchId,
    Client: client,
    StartDate: Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
    Assessor: assessorName,
    Status: 'Open'
  });
  return { batchId: batchId, clientFolderId: clientFolder.getId() };
}

// ========== API: SAVE ONE ITEM ==============================================

/**
 * Save an item and its selections. The payload comes from the client form.
 *
 * @param {Object} payload The item data payload.
 * @returns {Object} An object containing the new itemId and generated code.
 */
function saveItem(payload) {
  try {
    if (!payload || !payload.batchId) throw new Error('Missing batchId');
    // Ensure headers for items and selections
    ensureSheet_(SH_ITEMS,  HEADERS[SH_ITEMS]);
    ensureSheet_(SH_SELECT, HEADERS[SH_SELECT]);
    const itemId    = 'ITM-' + Utilities.getUuid().slice(0, 8);
    const code      = nextCodeForBatch_(payload.batchId);
    const createdAt = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const notesAuto = buildNotes_(payload);
    // Flatten AdditionalNotes: combine incident notes
    const addNotes = payload.incidentNotes || '';
    // Write item row
    appendRowAtKey_(SH_ITEMS, 'ItemID', {
ItemID: itemId,
      BatchID: payload.batchId,
      Code: code,
      PerilType: payload.peril || '',
      ImpactLevel: payload.impact || '',
      Type: payload.type || '',
      Title: payload.title || '',
      Artist: payload.artist || '',
      Material: payload.material || '',
      Date: payload.date || '',
      Dimensions: payload.dimensions || '',
      Features: payload.features || '',
      HistoricIssuesPresent: String(!!payload.historicYes).toUpperCase(),
      HistoricCode: String(payload.historicCode || '').toLowerCase(),
      HistoricNotes: payload.historicNotes || '',
      TreatmentTimeMinutes: payload.TreatmentTimeMinutes || '',
      TreatmentTimeUnit: payload.TreatmentTimeUnit || '',
      TreatmentTime: payload.TreatmentTime || '',
      AdditionalNotes: addNotes,
      OtherNotes: payload.otherNotes || '',
      CreatedAt: createdAt,
      Notes: getNoteFromNotesRules_(String(payload.historicCode || ''))
    });
    // Write selection rows
    (payload.selections || []).forEach(s => {
      const isLocal = !!s.localized;
      const note    = String(s.note || '');
      const needsRepl = /Replacement needed/i.test(note);
      const needsReuph= /Reupholstery needed/i.test(note);
      appendRowAtKey_(SH_SELECT, 'SelectionID', {
        SelectionID: 'SEL-' + Utilities.getUuid().slice(0, 8),
        ItemID: itemId,
        OptionCode: s.code || '',
        Severity: '',
        ExtentPercent: s.extent || '',
        Location: s.location || '',
        ItemType: payload.type || '',
        UseSeverity: '',
        SeverityWord: '',
        BasePhrase: '',
        IsLocalized: String(isLocal).toUpperCase(),
        LocPart: '',
        ExtPart: '',
        CondLine: '',
        GlobalFrag: isLocal ? '' : note,
        LocalText:  isLocal ? note : '',
        LocalWhere: s.location || '',
        NeedsReplacement: needsRepl ? 'TRUE' : 'FALSE',
        NeedsReupholstery: needsReuph ? 'TRUE' : 'FALSE'
      });
    });
    return { itemId: itemId, code: code };
  } catch (err) {
    console.error('saveItem error', err);
    throw err;
  }
}

// ========== API: SAVE PHOTOS (data URLs) ====================================

/**
 * Save an array of photos (data URLs) to Drive and record them in the Photos sheet.
 *
 * @param {string} batchId The batch ID.
 * @param {string} itemId The item ID.
 * @param {Array} photos An array of objects with a dataUrl property.
 * @returns {Object} An object with the number of photos saved.
 */
function savePhotosDataUrls(batchId, itemId, photos) {
  if (!batchId || !itemId) throw new Error('Missing batchId or itemId');
  if (!photos || !photos.length) return { saved: 0 };
  ensureSheet_(SH_PHOTOS, HEADERS[SH_PHOTOS]);
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('PHOTOS_' + batchId);
  if (!folderId) throw new Error('Photo folder not configured for this batch.');
  const folder = DriveApp.getFolderById(folderId);
  const when = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  let saved = 0;
  photos.forEach((p, idx) => {
    try {
      const dataUrl = String(p.dataUrl || '');
      const m = dataUrl.match(/^data:(image\/[^;]+);base64,(.+)$/);
      if (!m) return;
      const contentType = m[1];
      const bytes = Utilities.base64Decode(m[2]);
      const ext = contentType === 'image/png' ? 'png' : 'jpg';
      const name = itemId + '-' + when + '-' + String(idx + 1).padStart(2, '0') + '.' + ext;
      const file = folder.createFile(Utilities.newBlob(bytes, contentType, name));
      appendRowAtKey_(SH_PHOTOS, 'PhotoID', {
        PhotoID: Utilities.getUuid(),
        ItemID: itemId,
        Image: file.getId(),
        Caption: '',
        TakenAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
        Lat: '',
        Lng: '',
        Uploader: ''
      });
      saved++;
    } catch (err) {
      console.error('photo save failed', err);
    }
  });
  return { saved: saved };
}

// ========== ONE‑CLICK SETUP / REPAIR =======================================

/**
 * Convenience function to create any missing sheets and header columns. It
 * is safe to run multiple times and returns the database URL and sheet URLs.
 *
 * @returns {Object} An object with spreadsheetUrl and sheet URL map.
 */
function initializeDatabase() {
  Object.keys(HEADERS).forEach(name => ensureSheet_(name, HEADERS[name]));
  const urls = {};
  const ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) urls[name] = ss.getUrl() + '#gid=' + sh.getSheetId();
  });
  return { spreadsheetUrl: ss.getUrl(), sheets: urls };
}
