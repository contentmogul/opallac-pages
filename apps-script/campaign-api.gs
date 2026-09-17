/**
 * Campaign assets API — Apps Script Web App
 * ==========================================
 * ONE deployment of this script serves every campaign folder/sheet listed
 * in campaigns.json. You do NOT need a separate script per campaign.
 *
 * This replaces an earlier draft that scraped Google Drive's folder page
 * HTML for an anonymous file listing. That was unofficial (parsing an
 * internal, undocumented variable never meant for outside use) and fragile
 * (could break silently on any Drive UI change). This version uses only
 * Apps Script's own sanctioned services — DriveApp and SpreadsheetApp — so
 * every read and write goes through Google's real, supported surface,
 * authenticated as whoever deploys this script.
 *
 * SETUP (one-time, ~10 minutes):
 *   1. Go to https://script.google.com (NOT "Extensions > Apps Script" from
 *      inside one specific Sheet — that binds the script to just that sheet,
 *      which is not what we want; this one script serves all of them).
 *   2. Click "New project".
 *   3. Delete the placeholder code and paste this entire file in.
 *   4. Update MANIFEST_URL below if the repo/path ever changes.
 *   5. Deploy > New deployment > type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   6. Authorize it when prompted — it needs permission to read Drive and
 *      edit spreadsheets your account can access. That's expected: it's
 *      what lets one script serve every campaign folder/sheet.
 *   7. Copy the Web App URL (ends in /exec) into config.json's
 *      "apiEndpoint" field in this repo.
 *   8. Any time you edit this script: Deploy > Manage deployments > pencil
 *      icon > New version. Just clicking "Save" does not update the live
 *      Web App.
 *
 * SECURITY: every folderId/sheetId is checked against campaigns.json before
 * anything is read or written. This script can only touch a folder/sheet
 * that's actually listed there — not any other file your account happens
 * to have access to, even though it technically could.
 *
 * ── doGet — reads ──────────────────────────────────────────────────────
 *   ?action=folder&folderId=XXX
 *     -> { ok:true, images:[{id,name}], videos:[{id,name}], sheets:[{id,name}], docs:[{id,name}] }
 *
 *   ?action=sheet&sheetId=XXX&sheetName=YYY   (sheetName optional, defaults to first tab)
 *     -> { ok:true, headers:[...], rows:[ { row:3, values:{header: value, ...} }, ... ] }
 *
 * ── doPost — writes (Content-Type: text/plain, JSON string body — see
 *    the note in doPost for why not application/json) ──────────────────
 *   {
 *     "sheetId": "...",
 *     "sheetName": "Copywriting",     // optional
 *     "row": 3,                       // the row number from a prior ?action=sheet read
 *     "updates": { "CLIENT APPROVED?": true, "CLIENT FEEDBACK: ...": "text" }
 *   }
 *     -> { ok:true, row:3 }  or  { ok:false, error:"..." }
 */

var MANIFEST_URL = 'https://contentmogul.github.io/opallac-pages/campaigns.json';

function doGet(e) {
  var result;
  try {
    var action = e.parameter.action;
    if (action === 'folder') {
      result = listFolder(e.parameter.folderId);
    } else if (action === 'sheet') {
      result = readSheet(e.parameter.sheetId, e.parameter.sheetName);
    } else {
      result = { ok: false, error: 'Unknown or missing action.' };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return jsonOutput(result);
}

function doPost(e) {
  var result;
  try {
    // Apps Script Web Apps don't handle the CORS preflight (OPTIONS request)
    // that a JSON Content-Type triggers, so the page sends the JSON body as
    // Content-Type: text/plain instead — a CORS-"simple" request, no
    // preflight. We just parse the text as JSON here.
    var payload = JSON.parse(e.postData.contents);
    result = writeUpdate(payload);
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return jsonOutput(result);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function listFolder(folderId) {
  if (!folderId) return { ok: false, error: 'folderId is required.' };
  if (!isKnownFolder(folderId)) return { ok: false, error: 'folderId is not in campaigns.json.' };

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var images = [], videos = [], sheets = [], docs = [];
  while (files.hasNext()) {
    var f = files.next();
    var entry = { id: f.getId(), name: f.getName() };
    var mime = f.getMimeType();
    if (mime.indexOf('image/') === 0) images.push(entry);
    else if (mime.indexOf('video/') === 0) videos.push(entry);
    else if (mime === MimeType.GOOGLE_SHEETS) sheets.push(entry);
    else if (mime === MimeType.GOOGLE_DOCS) docs.push(entry);
  }

  return { ok: true, images: images, videos: videos, sheets: sheets, docs: docs };
}

function readSheet(sheetId, sheetName) {
  if (!sheetId) return { ok: false, error: 'sheetId is required.' };
  if (!isKnownSheet(sheetId)) return { ok: false, error: 'sheetId is not in campaigns.json.' };

  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getSheets()[0];
  if (!sheet) return { ok: false, error: 'Sheet/tab "' + sheetName + '" not found.' };

  var values = sheet.getDataRange().getValues();
  if (values.length === 0) return { ok: true, headers: [], rows: [] };

  var headers = values[0].map(function (h) { return String(h); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rowValues = values[r];
    var isBlank = rowValues.every(function (v) { return v === '' || v === null; });
    if (isBlank) continue;

    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = rowValues[c];
    }
    rows.push({ row: r + 1, values: obj }); // r+1 = 1-based actual sheet row number
  }

  return { ok: true, headers: headers, rows: rows };
}

function writeUpdate(payload) {
  var sheetId = payload.sheetId;
  var sheetName = payload.sheetName;
  var row = payload.row;
  var updates = payload.updates || {};

  if (!sheetId || !row) {
    return { ok: false, error: 'sheetId and row are required.' };
  }
  if (!isKnownSheet(sheetId)) {
    return { ok: false, error: 'sheetId is not in campaigns.json — refusing to write.' };
  }

  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getSheets()[0];
  if (!sheet) return { ok: false, error: 'Sheet/tab "' + sheetName + '" not found.' };

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerIndex = {};
  for (var c = 0; c < headerRow.length; c++) {
    headerIndex[String(headerRow[c])] = c;
  }

  for (var header in updates) {
    if (!updates.hasOwnProperty(header)) continue;
    var colIdx = headerIndex[header];
    if (colIdx === undefined) {
      return { ok: false, error: 'updates column "' + header + '" not found in header row.' };
    }
    sheet.getRange(row, colIdx + 1).setValue(updates[header]);
  }

  return { ok: true, row: row };
}

function getManifest_() {
  var res = UrlFetchApp.fetch(MANIFEST_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Could not read campaigns.json (HTTP ' + res.getResponseCode() + ').');
  }
  return JSON.parse(res.getContentText());
}

function isKnownFolder(folderId) {
  var manifest = getManifest_();
  for (var i = 0; i < manifest.length; i++) {
    if (manifest[i].folderId === folderId) return true;
  }
  return false;
}

function isKnownSheet(sheetId) {
  // A sheet is "known" if it belongs to a known folder — we don't need a
  // separate per-sheet list, listFolder() already tells the client which
  // sheet IDs live in a known folder, and that's the only way the client
  // learns a sheetId in the first place.
  var manifest = getManifest_();
  for (var i = 0; i < manifest.length; i++) {
    try {
      var files = DriveApp.getFolderById(manifest[i].folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        if (files.next().getId() === sheetId) return true;
      }
    } catch (e) {
      // skip a folder we can no longer access rather than failing the whole check
    }
  }
  return false;
}
