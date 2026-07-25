// ═══════════════════════════════════════════════════
// Lectio Divina / Reading Tracker — Sheets Sync Backend
// Reads/writes A1 (JSON) with B1 = last-saved timestamp.
// After pasting: Save, then Deploy → Manage deployments
// → pencil → Version: New version → Deploy.
// ═══════════════════════════════════════════════════

const DATA_TAB = '_data';

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Use the _data tab if it exists, otherwise the first sheet
  // (your current data lives on the first sheet at A1).
  return ss.getSheetByName(DATA_TAB) || ss.getSheets()[0];
}

function readJson_() {
  const sheet = getDataSheet_();
  let val = sheet.getRange('A1').getValue();
  if (!val || String(val).trim().charAt(0) !== '{') {
    val = sheet.getRange('A2').getValue();  // legacy fallback
  }
  const s = val ? String(val).trim() : '';
  return (s.charAt(0) === '{') ? s : '{}';
}

// POST save — client sends form-encoded: data=<json>
function doPost(e) {
  try {
    const json = e && e.parameter ? e.parameter.data : null;
    if (json) {
      const sheet = getDataSheet_();
      sheet.getRange('A1').setValue(json);
      sheet.getRange('B1').setValue('Saved: ' + new Date().toISOString());
      populateReadableTabs_(json);
    }
  } catch (err) {
    // no-cors client can't read the response; fail silently
  }
  return ContentService.createTextOutput('ok');
}

// JSONP GET read-back — ?action=load&callback=xyz
function doGet(e) {
  const cb = ((e && e.parameter && e.parameter.callback) || '')
    .replace(/[^a-zA-Z0-9_$]/g, '');
  let payload = '{}';
  try {
    payload = readJson_();
  } catch (err) {
    payload = '{}';
  }
  const body = cb ? (cb + '(' + payload + ')') : payload;
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Optional readable mirror tabs
function populateReadableTabs_(jsonString) {
  let data;
  try { data = JSON.parse(jsonString); } catch (e) { return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeArrayTab_(ss, 'Reading Books', (data.reading || {}).books);
  writeArrayTab_(ss, 'Reading Logs', (data.reading || {}).logs);
  writeArrayTab_(ss, 'Study Journal', (data.study || {}).entries);
}

function writeArrayTab_(ss, tabName, arr) {
  if (!Array.isArray(arr) || arr.length === 0) return;
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  sheet.clear();
  const keys = Array.from(arr.reduce((set, o) => {
    Object.keys(o || {}).forEach(k => set.add(k));
    return set;
  }, new Set()));
  sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  const rows = arr.map(o => keys.map(k => {
    const v = o[k];
    return (v && typeof v === 'object') ? JSON.stringify(v) : v;
  }));
  if (rows.length) sheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
}
