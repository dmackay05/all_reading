// ═══════════════════════════════════════════════════
// Lectio Divina / Reading Tracker — Sheets Sync Backend
// Paste this whole file into Extensions → Apps Script,
// replacing all existing content. Then Deploy → Manage
// Deployments → edit (pencil) → Version: New → Deploy.
// ═══════════════════════════════════════════════════

const DATA_TAB = '_data';

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_TAB);
    sheet.getRange(1, 1).setValue('json');
  }
  return sheet;
}

// Handles the POST save (no-cors from the browser, form-encoded body: data=<json>)
function doPost(e) {
  try {
    const json = e.parameter.data;
    const sheet = getDataSheet_();
    sheet.getRange(2, 1).setValue(json);

    // Optional: populate human-readable tabs from the JSON.
    populateReadableTabs_(json);

  } catch (err) {
    // Swallow errors — the client uses no-cors and can't read the response anyway.
  }
  return ContentService.createTextOutput('ok');
}

// Handles the JSONP GET read-back (?action=load&callback=xyz)
function doGet(e) {
  const callback = (e.parameter.callback || '').replace(/[^a-zA-Z0-9_$]/g, '');
  let payload = '{}';

  try {
    if (e.parameter.action === 'load') {
      const sheet = getDataSheet_();
      const val = sheet.getRange(2, 1).getValue();
      payload = val ? val.toString() : '{}';
    }
  } catch (err) {
    payload = '{}';
  }

  const body = callback ? (callback + '(' + payload + ')') : payload;
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Optional — mirrors key arrays into readable tabs for browsing in Sheets.
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

  const keys = Array.from(
    arr.reduce((set, obj) => {
      Object.keys(obj || {}).forEach(k => set.add(k));
      return set;
    }, new Set())
  );

  sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  const rows = arr.map(obj => keys.map(k => {
    const v = obj[k];
    return (v && typeof v === 'object') ? JSON.stringify(v) : v;
  }));
  if (rows.length) sheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
}
