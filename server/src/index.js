'use strict';

const express = require('express');
const { GoogleAuth } = require('google-auth-library');

// clients.json maps each client to their own registry Sheet (columns Label,
// Drive Folder URL, Active). A folder/sheet is authorized only if it appears
// as an active row in one of these registries — this replaces the old static
// campaigns.json manifest so day-to-day campaign additions for an existing
// client are just a new Sheet row, no repo/Actions involvement.
const CLIENTS_URL = 'https://contentmogul.github.io/opallac-pages/clients.json';
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

async function getClients() {
  const res = await fetch(CLIENTS_URL, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Could not read clients.json (HTTP ${res.status}).`);
  return res.json();
}

// Mirrors the /folders/<id> extraction the add-campaign-folder.yml workflow
// (and the retired manual campaigns.json flow) used, so a pasted Drive share
// URL in the registry Sheet resolves the same way everywhere.
function extractFolderId(url) {
  const m = String(url || '').match(/\/folders\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : null;
}

async function isKnownRegistrySheet(sheetId) {
  const clients = await getClients();
  return clients.some((c) => c.registrySheetId === sheetId);
}

// Reads a client's registry Sheet (Label, Drive Folder URL, Active columns,
// matched by header name so column order doesn't matter) and returns the
// active rows as {label, folderId}.
async function readRegistryRows(sheetId, headers) {
  const tab = await resolveFirstSheetName(sheetId, headers);
  const res = await fetch(`${SHEETS_API}/${sheetId}/values/${encodeURIComponent(tab)}`, { headers });
  if (!res.ok) throw new Error(`Registry sheet read failed for ${sheetId} (HTTP ${res.status}).`);
  const body = await res.json();
  const values = body.values || [];
  if (values.length === 0) return [];

  const headerRow = values[0].map((h) => String(h).trim().toLowerCase());
  const labelIdx = headerRow.indexOf('label');
  const urlIdx = headerRow.indexOf('drive folder url');
  const activeIdx = headerRow.indexOf('active');

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const rowValues = values[r] || [];
    const activeRaw = activeIdx >= 0 ? rowValues[activeIdx] : '';
    if (String(activeRaw).trim().toUpperCase() !== 'TRUE') continue;
    const folderId = urlIdx >= 0 ? extractFolderId(rowValues[urlIdx]) : null;
    if (!folderId) continue;
    rows.push({ label: (labelIdx >= 0 && rowValues[labelIdx]) || '', folderId });
  }
  return rows;
}

// Every active row across every client's registry — the live replacement for
// the old static campaigns.json manifest.
async function getKnownFolders(headers) {
  const clients = await getClients();
  const known = [];
  for (const c of clients) {
    try {
      known.push(...(await readRegistryRows(c.registrySheetId, headers)));
    } catch (err) {
      // Skip a registry we can no longer read rather than failing the whole check.
    }
  }
  return known;
}

async function isKnownFolder(folderId, headers) {
  const known = await getKnownFolders(headers);
  return known.some((c) => c.folderId === folderId);
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

// A sheet is "known" if it lives inside a known (registry-listed) folder —
// mirrors the abandoned Apps Script draft's isKnownSheet(): there is no
// separate per-sheet allowlist, listFolder() is the only way the client ever
// learns a campaign sheetId in the first place. The registry sheet itself is
// authorized separately via isKnownRegistrySheet().
async function isKnownSheet(sheetId, headers) {
  const known = await getKnownFolders(headers);
  for (const c of known) {
    try {
      const files = await listSheetFilesInFolder(c.folderId, headers);
      if (files.some((f) => f.id === sheetId)) return true;
    } catch (err) {
      // Skip a folder we can no longer read rather than failing the whole check.
    }
  }
  return false;
}

async function readRegistry(sheetId) {
  if (!sheetId) return { ok: false, error: 'sheetId is required.' };
  if (!(await isKnownRegistrySheet(sheetId))) return { ok: false, error: 'sheetId is not in clients.json.' };
  const headers = await authHeaders();
  const rows = await readRegistryRows(sheetId, headers);
  return { ok: true, rows };
}

async function listFolder(folderId) {
  if (!folderId) return { ok: false, error: 'folderId is required.' };

  const headers = await authHeaders();
  if (!(await isKnownFolder(folderId, headers))) {
    return { ok: false, error: 'folderId is not an active row in any client registry.' };
  }

  const images = [];
  const videos = [];
  const sheets = [];
  const docs = [];

  let pageToken;
  do {
    const q = encodeURIComponent(`'${folderId}' in parents and trashed = false`);
    const pageParam = pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : '';
    const url = `${DRIVE_API}/files?q=${q}&fields=files(id,name,mimeType,videoMediaMetadata(width,height)),nextPageToken&pageSize=1000${pageParam}`;
    const res = await fetch(url, { headers });
    if (!res.ok) throw new Error(`Drive list failed for folder ${folderId} (HTTP ${res.status}).`);
    const body = await res.json();
    for (const f of body.files || []) {
      const entry = { id: f.id, name: f.name };
      if (f.mimeType.startsWith('image/')) images.push(entry);
      else if (f.mimeType.startsWith('video/')) {
        // Real dimensions from Drive so the frontend can size the preview
        // to the video's actual aspect ratio instead of assuming 16:9.
        const meta = f.videoMediaMetadata;
        if (meta && meta.width && meta.height) {
          entry.width = meta.width;
          entry.height = meta.height;
        }
        videos.push(entry);
      }
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
  if (!(await isKnownSheet(sheetId, headers))) return { ok: false, error: 'sheetId is not in an active client registry.' };

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

// Per-ASSET (image/video) feedback lives in its own "Asset Feedback" tab
// inside the same campaign spreadsheet, keyed by filename — separate from
// the copy table's rows entirely, so per-copy-variant data and per-creative
// approval never share a row space. The tab is created lazily, only the
// first time someone actually comments on an asset in that campaign, so a
// client's sheet gains no clutter until there's something to record.
const ASSET_FEEDBACK_TAB = 'Asset Feedback';
const ASSET_FEEDBACK_HEADERS = ['File', 'Approved', 'Feedback'];

async function ensureAssetFeedbackTab(sheetId, headers) {
  const metaRes = await fetch(`${SHEETS_API}/${sheetId}?fields=sheets.properties`, { headers });
  if (!metaRes.ok) throw new Error(`Could not read spreadsheet metadata (HTTP ${metaRes.status}).`);
  const meta = await metaRes.json();
  const exists = (meta.sheets || []).some((s) => s.properties.title === ASSET_FEEDBACK_TAB);
  if (exists) return;

  const addRes = await fetch(`${SHEETS_API}/${sheetId}:batchUpdate`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests: [{ addSheet: { properties: { title: ASSET_FEEDBACK_TAB } } }] }),
  });
  if (!addRes.ok) {
    const text = await addRes.text().catch(() => '');
    throw new Error(`Could not create "${ASSET_FEEDBACK_TAB}" tab (HTTP ${addRes.status}): ${text}`);
  }

  const headerRes = await fetch(
    `${SHEETS_API}/${sheetId}/values/${encodeURIComponent(ASSET_FEEDBACK_TAB + '!A1')}?valueInputOption=RAW`,
    {
      method: 'PUT',
      headers: { ...headers, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [ASSET_FEEDBACK_HEADERS] }),
    }
  );
  if (!headerRes.ok) {
    const text = await headerRes.text().catch(() => '');
    throw new Error(`Could not write header row for "${ASSET_FEEDBACK_TAB}" (HTTP ${headerRes.status}): ${text}`);
  }
}

async function readAssetFeedbackValues(sheetId, headers) {
  await ensureAssetFeedbackTab(sheetId, headers);
  const res = await fetch(`${SHEETS_API}/${sheetId}/values/${encodeURIComponent(ASSET_FEEDBACK_TAB)}`, { headers });
  if (!res.ok) throw new Error(`Could not read "${ASSET_FEEDBACK_TAB}" (HTTP ${res.status}).`);
  const body = await res.json();
  return body.values && body.values.length ? body.values : [ASSET_FEEDBACK_HEADERS];
}

// Returns { ok, byFile: { "<filename>": { File, Approved, Feedback } } } so
// the frontend can prefill previously-saved feedback on load, not just show
// blank boxes every visit.
async function readAssetFeedback(sheetId) {
  if (!sheetId) return { ok: false, error: 'sheetId is required.' };
  const headers = await authHeaders();
  if (!(await isKnownSheet(sheetId, headers))) return { ok: false, error: 'sheetId is not in an active client registry.' };

  const values = await readAssetFeedbackValues(sheetId, headers);
  const headerRow = values[0];
  const fileIdx = headerRow.indexOf('File');
  const byFile = {};
  for (let r = 1; r < values.length; r++) {
    const rowValues = values[r] || [];
    const file = rowValues[fileIdx];
    if (!file) continue;
    const obj = {};
    headerRow.forEach((h, c) => {
      obj[h] = rowValues[c] !== undefined ? rowValues[c] : '';
    });
    byFile[file] = obj;
  }
  return { ok: true, byFile };
}

// Writes {Approved, Feedback} for one filename, creating that row (and the
// tab itself, on the very first call for a campaign) if it doesn't exist yet.
async function writeAssetFeedback(payload) {
  const { sheetId, fileName, updates } = payload || {};
  if (!sheetId || !fileName) return { ok: false, error: 'sheetId and fileName are required.' };

  const headers = await authHeaders();
  if (!(await isKnownSheet(sheetId, headers))) {
    return { ok: false, error: 'sheetId is not in an active client registry — refusing to write.' };
  }

  const values = await readAssetFeedbackValues(sheetId, headers);
  const headerRow = values[0];
  const fileIdx = headerRow.indexOf('File');

  let targetRow = -1; // 1-based sheet row
  for (let r = 1; r < values.length; r++) {
    if ((values[r][fileIdx] || '') === fileName) {
      targetRow = r + 1;
      break;
    }
  }

  if (targetRow === -1) {
    const rowValues = headerRow.map((h) => {
      if (h === 'File') return fileName;
      if (Object.prototype.hasOwnProperty.call(updates || {}, h)) return updates[h];
      return '';
    });
    const appendRes = await fetch(
      `${SHEETS_API}/${sheetId}/values/${encodeURIComponent(ASSET_FEEDBACK_TAB)}:append?valueInputOption=RAW`,
      {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: [rowValues] }),
      }
    );
    if (!appendRes.ok) {
      const text = await appendRes.text().catch(() => '');
      return { ok: false, error: `Could not create asset feedback row (HTTP ${appendRes.status}): ${text}` };
    }
    return { ok: true, created: true };
  }

  const data = [];
  for (const header of Object.keys(updates || {})) {
    const colIdx = headerRow.indexOf(header);
    if (colIdx === -1) return { ok: false, error: `updates column "${header}" not found in "${ASSET_FEEDBACK_TAB}".` };
    data.push({ range: `${ASSET_FEEDBACK_TAB}!${colToA1(colIdx)}${targetRow}`, values: [[updates[header]]] });
  }
  if (data.length === 0) return { ok: true, row: targetRow };

  const res = await fetch(`${SHEETS_API}/${sheetId}/values:batchUpdate`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({ valueInputOption: 'RAW', data }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    return { ok: false, error: `Sheets write failed (HTTP ${res.status}): ${text}` };
  }
  return { ok: true, row: targetRow };
}

async function writeUpdate(payload) {
  const { sheetId, sheetName, row, updates } = payload || {};
  if (!sheetId || !row) return { ok: false, error: 'sheetId and row are required.' };

  const headers = await authHeaders();
  if (!(await isKnownSheet(sheetId, headers))) {
    return { ok: false, error: 'sheetId is not in an active client registry — refusing to write.' };
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
    else if (action === 'registry') result = await readRegistry(sheetId);
    else if (action === 'assetFeedback') result = await readAssetFeedback(sheetId);
    else result = { ok: false, error: 'Unknown or missing action.' };
    res.json(result);
  } catch (err) {
    res.json({ ok: false, error: String((err && err.message) || err) });
  }
});

app.post('/', async (req, res) => {
  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : req.body;
    const result = payload && payload.fileName ? await writeAssetFeedback(payload) : await writeUpdate(payload);
    res.json(result);
  } catch (err) {
    res.json({ ok: false, error: String((err && err.message) || err) });
  }
});

const port = process.env.PORT || 8080;
app.listen(port, () => {
  console.log(`opallac-drive-sheets-api listening on ${port}`);
});
