// api/sheets.js
// Proxies Google Sheets & Drive API calls using the signed-in user's access token.
// The frontend never touches Google APIs directly — this keeps auth centralized.

import { getSession } from './auth/me.js';

const SHEET_NAME = 'Job Search Tracker';
const HEADERS = ['ID','Title','Company','Location','Salary','URL','Posted','Search Term','Source','Status','Fit Score','Fit Notes','Date Added'];

async function gFetch(url, token, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Google API ${res.status}: ${text}`);
  }
  return res.json();
}

// ── Find or create the tracker spreadsheet in the user's Drive ────────────────
async function findOrCreateSheet(token) {
  // Search Drive for existing sheet
  const query = encodeURIComponent(`name='${SHEET_NAME}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`);
  const list = await gFetch(`https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name)`, token);

  if (list.files && list.files.length > 0) {
    return list.files[0].id;
  }

  // Create new spreadsheet
  const created = await gFetch('https://sheets.googleapis.com/v4/spreadsheets', token, {
    method: 'POST',
    body: JSON.stringify({
      properties: { title: SHEET_NAME },
      sheets: [{
        properties: { title: 'Jobs' },
        data: [{
          startRow: 0, startColumn: 0,
          rowData: [{ values: HEADERS.map(h => ({ userEnteredValue: { stringValue: h } })) }],
        }],
      }],
    }),
  });

  return created.spreadsheetId;
}

// ── Read all rows ─────────────────────────────────────────────────────────────
async function readRows(token, sheetId) {
  const data = await gFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Jobs!A2:M`,
    token
  );
  const rows = data.values || [];
  return rows.map(row => ({
    id:          row[0]  || '',
    title:       row[1]  || '',
    company:     row[2]  || '',
    location:    row[3]  || '',
    salary:      row[4]  || '',
    url:         row[5]  || '',
    posted:      row[6]  || '',
    searchTitle: row[7]  || '',
    source:      row[8]  || '',
    status:      row[9]  || 'Saved',
    fitScore:    row[10] || '',
    fitNotes:    row[11] || '',
    dateAdded:   row[12] || '',
  }));
}

// ── Append a new row ──────────────────────────────────────────────────────────
async function appendRow(token, sheetId, job) {
  await gFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Jobs!A:M:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    token,
    {
      method: 'POST',
      body: JSON.stringify({
        values: [[
          job.id, job.title, job.company, job.location || '',
          job.salary || 'Not specified', job.url || '', job.posted || '',
          job.searchTitle || '', job.source || '', job.status || 'Saved',
          job.fitScore || '', job.fitNotes || '',
          job.dateAdded || new Date().toLocaleDateString(),
        ]],
      }),
    }
  );
}

// ── Find a row by job ID and return its row number (1-based, including header) ─
async function findRowNumber(token, sheetId, jobId) {
  const data = await gFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Jobs!A:A`,
    token
  );
  const col = data.values || [];
  for (let i = 1; i < col.length; i++) {
    if (col[i][0] === jobId) return i + 1; // +1 because sheets are 1-indexed
  }
  return null;
}

// ── Update status in an existing row ─────────────────────────────────────────
async function updateStatus(token, sheetId, jobId, newStatus) {
  const rowNum = await findRowNumber(token, sheetId, jobId);
  if (!rowNum) return false;

  await gFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Jobs!J${rowNum}?valueInputOption=RAW`,
    token,
    {
      method: 'PUT',
      body: JSON.stringify({ values: [[newStatus]] }),
    }
  );
  return true;
}

// ── Main handler ──────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'Not authenticated' });

  const token = session.access_token;

  try {
    // GET /api/sheets — load all jobs
    if (req.method === 'GET') {
      const sheetId = await findOrCreateSheet(token);
      const rows = await readRows(token, sheetId);
      return res.json({ sheetId, rows });
    }

    // POST /api/sheets — save or update a job
    if (req.method === 'POST') {
      const { job, action } = req.body;
      const sheetId = req.body.sheetId || await findOrCreateSheet(token);

      if (action === 'updateStatus') {
        const updated = await updateStatus(token, sheetId, job.id, job.status);
        if (!updated) {
          // Row doesn't exist yet — append it
          await appendRow(token, sheetId, job);
        }
        return res.json({ ok: true });
      }

      if (action === 'save') {
        const rowNum = await findRowNumber(token, sheetId, job.id);
        if (rowNum) {
          await updateStatus(token, sheetId, job.id, job.status);
        } else {
          await appendRow(token, sheetId, job);
        }
        return res.json({ ok: true, sheetId });
      }

      return res.status(400).json({ error: 'Unknown action' });
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error('Sheets error:', err);
    res.status(500).json({ error: err.message });
  }
}
