'use strict';

const express = require('express');
const { GoogleAuth } = require('google-auth-library');

// Same manifest the frontend (drive-assets.html) and the abandoned Apps Script
// draft (apps-script/campaign-api.gs) both read from — this is the single
// source of truth for which folders/sheets this API is allowed to touch.
const MANIFEST_URL = 'https://contentmogul.github.io/opallac-pages/campaigns.json';
const DRIVE_API = 'https://www.googleapis.com/drive/v3';
const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

// No key file: Cloud Run supplies Application Default Credentials for the
// service account it runs as, and that SA holds no project-wide IAM roles —
// its Drive/Sheets access comes entirely from folder/sheet sharing.
const auth = new GoogleAuth({
  scopes: [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
  ],
});

async function authHeaders() {
  const client = await auth.getClient();
  const token = await client.getAccessToken();
  return { Authorization: `Bearer ${token.token}` };
}

async function getManifest() {
  const res = await fetch(MANIFEST_URL, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Could not read campaigns.json (HTTP ${res.status}).`);
  return res.json();
}

async function isKnownFolder(folderId) {
  const manifest = await getManifest();
  return manifest.some((c) => c.folderId === folderId);
}

async function listSheetFilesInFolder(folderId, headers) {
  const q = encodeURIComponent(
    `'${folderId}' in parents and trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet'`
  );
  const url = `${DRIVE_API}/files?q=${q}&fields=files(id,name)&pageSize=1000`;
  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`Drive list failed for folder ${folderId} (HTTP ${res.status}).`);
  const body = await res.json();
  return body.files || [];
}

// A sheet is "known" if it lives inside a known folder — mirrors the abandoned
// Apps Script draft's isKnownSheet(): there is no separate per-sheet allowlist,
// listFolder() is the only way the client ever learns a sheetId in the first place.
async function isKnownSheet(sheetId, headers) {
  const manifest = await getManifest();
  for (const c of manifest) {
    try {
      const files = await listSheetFilesInFolder(c.folderId, headers);
      if (files.some((f) => f.id === sheetId)) return true;
    } catch (err) {
      // Skip a folder we can no longer read rather than failing the whole check.
    }
  }
  return false;
}

async function listFolder(folderId) {
  if (!folderId) return { ok: false, error: 'folderId is required.' };
  if (!(await isKnownFolder(folderId))) return { ok: false, error: 'folderId is not in campaigns.json.' };

  const headers = await authHeaders();
  const images = [];
  const videos = [];
  const sheets = [];
  const docs = [];

  let pageToken;
  do {
    const q = encodeURIComponent(`'${folderId}' in parents and trashed = false`);
    const pageParam = pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : '';
    const url = `${DRIVE_API}/files?q=${q}&fields=files(id,name,mimeType),nextPageToken&pageSize=1000${pageParam}`;
    const res = await fetch(url, { headers });
    if (!res.ok) throw new Error(`Drive list failed for folder ${folderId} (HTTP ${res.status}).`);
    const body = await res.json();
    for (const f of body.files || []) {
      const entry = { id: f.id, name: f.name };
      if (f.mimeType.startsWith('image/')) images.push(entry);
      else if (f.mimeType.startsWith('video/')) videos.push(entry);
      else if (f.mimeType === 'application/vnd.google-apps.spreadsheet') sheets.push(entry);
      else if (f.mimeType === 'application/vnd.google-apps.document') docs.push(entry);
    }
    pageToken = body.nextPageToken;
  } while (pageToken);

  return { ok: true, images, videos, sheets, docs };
}

// 0-based column index -> A1 column letters (0 -> A, 25 -> Z, 26 -> AA, ...).
function colToA1(idx) {
  let n = idx + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function resolveFirstSheetName(sheetId, headers) {
  const res = await fetch(`${SHEETS_API}/${sheetId}?fields=sheets.properties.title`, { headers });
  if (!res.ok) throw new Error(`Could not read spreadsheet metadata (HTTP ${res.status}).`);
  const body = await res.json();
  const first = body.sheets && body.sheets[0] && body.sheets[0].properties && body.sheets[0].properties.title;
  if (!first) throw new Error('Spreadsheet has no tabs.');
  return first;
}

async function readSheet(sheetId, sheetName) {
  if (!sheetId) return { ok: false, error: 'sheetId is required.' };
  const headers = await authHeaders();
  if (!(await isKnownSheet(sheetId, headers))) return { ok: false, error: 'sheetId is not in campaigns.json.' };

  const tab = sheetName || (await resolveFirstSheetName(sheetId, headers));
  const res = await fetch(`${SHEETS_API}/${sheetId}/values/${encodeURIComponent(tab)}`, { headers });
  if (!res.ok) return { ok: false, error: `Sheet/tab "${tab}" not found or could not be read (HTTP ${res.status}).` };
  const body = await res.json();
  const values = body.values || [];
  if (values.length === 0) return { ok: true, headers: [], rows: [] };

  const headerRow = values[0].map((h) => String(h));
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const rowValues = values[r] || [];
    const isBlank = headerRow.every((_, c) => rowValues[c] === undefined || rowValues[c] === '');
    if (isBlank) continue;
    const obj = {};
    headerRow.forEach((h, c) => {
      obj[h] = rowValues[c] !== undefined ? rowValues[c] : '';
    });
    rows.push({ row: r + 1, values: obj }); // r+1 = 1-based actual sheet row number
  }

  return { ok: true, headers: headerRow, rows };
}

async function writeUpdate(payload) {
  const { sheetId, sheetName, row, updates } = payload || {};
  if (!sheetId || !row) return { ok: false, error: 'sheetId and row are required.' };

  const headers = await authHeaders();
  if (!(await isKnownSheet(sheetId, headers))) {
    return { ok: false, error: 'sheetId is not in campaigns.json — refusing to write.' };
  }

  const tab = sheetName || (await resolveFirstSheetName(sheetId, headers));
  const headerRes = await fetch(`${SHEETS_API}/${sheetId}/values/${encodeURIComponent(`${tab}!1:1`)}`, { headers });
  if (!headerRes.ok) return { ok: false, error: `Sheet/tab "${tab}" not found.` };
  const headerBody = await headerRes.json();
  const headerRow = (headerBody.values && headerBody.values[0]) || [];
  const headerIndex = {};
  headerRow.forEach((h, i) => {
    headerIndex[String(h)] = i;
  });

  const data = [];
  for (const header of Object.keys(updates || {})) {
    const colIdx = headerIndex[header];
    if (colIdx === undefined) return { ok: false, error: `updates column "${header}" not found in header row.` };
    data.push({ range: `${tab}!${colToA1(colIdx)}${row}`, values: [[updates[header]]] });
  }
  if (data.length === 0) return { ok: true, row };

  const res = await fetch(`${SHEETS_API}/${sheetId}/values:batchUpdate`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({ valueInputOption: 'RAW', data }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    return { ok: false, error: `Sheets write failed (HTTP ${res.status}): ${text}` };
  }

  return { ok: true, row };
}

const app = express();

app.use((req, res, next) => {
  res.set('Access-Control-Allow-Origin', '*');
  res.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.set('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }
  next();
});

// drive-assets.html posts JSON with Content-Type: text/plain on purpose — a
// CORS-"simple" request avoids the preflight an application/json body would
// trigger. Accept any content type as raw text and JSON.parse it ourselves.
app.use(express.text({ type: '*/*' }));

app.get('/', async (req, res) => {
  try {
    const { action, folderId, sheetId, sheetName } = req.query;
    let result;
    if (action === 'folder') result = await listFolder(folderId);
    else if (action === 'sheet') result = await readSheet(sheetId, sheetName);
    else result = { ok: false, error: 'Unknown or missing action.' };
    res.json(result);
  } catch (err) {
    res.json({ ok: false, error: String((err && err.message) || err) });
  }
});

app.post('/', async (req, res) => {
  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : req.body;
    const result = await writeUpdate(payload);
    res.json(result);
  } catch (err) {
    res.json({ ok: false, error: String((err && err.message) || err) });
  }
});

const port = process.env.PORT || 8080;
app.listen(port, () => {
  console.log(`opallac-drive-sheets-api listening on ${port}`);
});
