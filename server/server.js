'use strict';

require('dotenv').config();

const express    = require('express');
const cors       = require('cors');
const { google } = require('googleapis');
const path       = require('path');
const crypto     = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());

// ── Serve static files ────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, '..')));
app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, '..', 'EvalForm.html'));
});
app.get('/adboard', (_req, res) => {
  res.sendFile(path.join(__dirname, '..', 'dashboard.html'));
});

// ── Admin auth ────────────────────────────────────────────────────────────────
function adminToken() {
  const pw = process.env.ADMIN_PASSWORD || 'admin';
  return crypto.createHmac('sha256', 'eval-dash-v1').update(pw).digest('hex');
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || '';
  const token  = header.startsWith('Bearer ') ? header.slice(7) : req.query.token;
  if (!token || token !== adminToken()) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}

// POST /adboard/login
app.post('/adboard/login', (req, res) => {
  const { password } = req.body || {};
  if (!password || password !== (process.env.ADMIN_PASSWORD || 'admin')) {
    return res.status(401).json({ error: 'סיסמה שגויה' });
  }
  res.json({ token: adminToken() });
});

// GET /api/dashboard
app.get('/api/dashboard', requireAuth, async (req, res) => {
  try {
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const sheets  = getSheetsClient();
    const tabName = await resolveTabName(sheets, sheetId);

    const lastCol = columnLetter(SHEET_HEADERS.length);
    const result  = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `${tabName}!A:${lastCol}`,
    });

    const rows = result.data.values || [];
    if (rows.length <= 1) return res.json({ data: [], headers: SHEET_HEADERS });

    const data = rows.slice(1).map((row, idx) => ({
      id:            idx,
      timestamp:     row[0]  || '',
      questionnaire: row[1]  || '',
      empName:       row[2]  || '',
      evalName:      row[3]  || '',
      evalDate:      row[4]  || '',
      avgScore:      row[5]  || '',
      allFields:     row,
    }));

    res.json({ data, headers: SHEET_HEADERS });
  } catch (err) {
    console.error('Dashboard error:', err.message);
    res.status(500).json({ error: 'Failed to load data' });
  }
});

// ── Google Sheets auth ────────────────────────────────────────────────────────
function getSheetsClient() {
  const raw = process.env.GOOGLE_CREDENTIALS_JSON;
  if (!raw) throw new Error('GOOGLE_CREDENTIALS_JSON env var is not set');

  const credentials = JSON.parse(raw);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// ── Column headers (must match row order in buildRow) ─────────────────────────
const SHEET_HEADERS = [
  'חותמת זמן',
  'שאלון',
  'שם עובד/ת',
  'שם מעריך/ה',
  'תאריך שיחה',
  'ציון ממוצע',
  // Section: מקצוענות ומצוינות (q1–q6, all questionnaires)
  'מקצוענות | בקיאות ושליטה מקצועית',
  'מקצוענות | תפוקות עבודה',
  'מקצוענות | איכות ובטיחות',
  'מקצוענות | שיפור מתמיד',
  'מקצוענות | הבנת הסביבה העסקית',
  'מקצוענות | עמידה ביעדים',
  // Section: אחריות ומחויבות (q7–q9, all questionnaires)
  'אחריות | מחויבות ליחידה',
  'אחריות | ראש גדול',
  'אחריות | לויאליות',
  // Section: יחסי אנוש (q10–q12, all questionnaires)
  'יחסי אנוש | שיתוף פעולה',
  'יחסי אנוש | עבודת צוות',
  'יחסי אנוש | קבלת משוב',
  // Section: יוזמה (q13–q15, all questionnaires)
  'יוזמה | נגישות לשטח',
  'יוזמה | פרואקטיביות',
  'יוזמה | גורם לדברים לקרות',
  // Section: מנהיג שינוי (q16–q18, questionnaire 3 only — empty for Q1/Q2)
  'מנהיג שינוי | דוגמא אישית',
  'מנהיג שינוי | מניע ומעצים',
  'מנהיג שינוי | אימון וחניכה',
  // Section: מנהל אנשים (q19–q21, questionnaire 3 only — empty for Q1/Q2)
  'מנהל אנשים | יעד וכיוון',
  'מנהל אנשים | ניהול צוות מנהלים',
  'מנהל אנשים | חיבור לשטח',
  // Section: עבודה לצד לחימה (q16 for Q1/Q2, q22 for Q3)
  'לחימה | תפקוד',
  // Goals
  'יעד 1 - תיאור',
  'יעד 1 - יעד',
  'יעד 1 - תאריך',
  'יעד 2 - תיאור',
  'יעד 2 - יעד',
  'יעד 2 - תאריך',
  'יעד 3 - תיאור',
  'יעד 3 - יעד',
  'יעד 3 - תאריך',
  // Notes
  'יעד פיתוח אישי',
  'נקודות לחיזוק',
  'נקודות לשיפור',
];

// ── Row builder ───────────────────────────────────────────────────────────────
// Maps submitted JSON body → ordered array matching SHEET_HEADERS.
// Q1 and Q2 have 16 questions (q1–q15 shared + q16 = warfare).
// Q3 has 22 questions (q1–q15 shared + q16–q21 manager + q22 = warfare).
function buildRow(body) {
  const qType = String(body.questionnaire || '1');
  const isQ3  = qType === '3';

  const qName = {
    '1': 'שאלון 1',
    '2': 'שאלון 2',
    '3': 'שאלון 3',
  }[qType] || qType;

  // q1–q15 are shared across all questionnaires
  const shared = Array.from({ length: 15 }, (_, i) => body[`q${i + 1}`] || '');

  // Manager columns: q16–q21 (Q3 only, empty otherwise)
  const managerCols = isQ3
    ? Array.from({ length: 6 }, (_, i) => body[`q${i + 16}`] || '')
    : ['', '', '', '', '', ''];

  // Warfare: q16 for Q1/Q2, q22 for Q3
  const warfare = isQ3 ? (body.q22 || '') : (body.q16 || '');

  return [
    new Date().toLocaleString('he-IL', { timeZone: 'Asia/Jerusalem' }),
    qName,
    body.empName    || '',
    body.evalName   || '',
    body.evalDate   || '',
    body.avgScore   || '',
    ...shared,
    ...managerCols,
    warfare,
    body.goal1_desc   || '',
    body.goal1_target || '',
    body.goal1_date   || '',
    body.goal2_desc   || '',
    body.goal2_target || '',
    body.goal2_date   || '',
    body.goal3_desc   || '',
    body.goal3_target || '',
    body.goal3_date   || '',
    body.personalDev  || '',
    body.strengths    || '',
    body.improvements || '',
  ];
}

// ── Validation ────────────────────────────────────────────────────────────────
function validateBody(body) {
  const required = ['empName', 'evalName', 'evalDate', 'questionnaire',
                    'goal1_desc', 'goal1_target', 'goal1_date',
                    'personalDev', 'strengths', 'improvements'];
  const missing = required.filter(f => !body[f] || String(body[f]).trim() === '');
  return missing;
}

// ── Resolve actual tab name (handles Hebrew default "גיליון1" etc.) ───────────
async function resolveTabName(sheets, spreadsheetId) {
  const configured = process.env.SHEET_TAB_NAME || 'Sheet1';
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const titles = meta.data.sheets.map(s => s.properties.title);
  if (titles.includes(configured)) return configured;
  // Fall back to first sheet and log a helpful message
  console.log(`Tab "${configured}" not found. Available tabs: [${titles.join(', ')}]. Using "${titles[0]}".`);
  console.log(`Tip: set SHEET_TAB_NAME=${titles[0]} in your .env to silence this.`);
  return titles[0];
}

// ── Ensure header row exists ───────────────────────────────────────────────────
async function ensureHeaders(sheets, spreadsheetId, tabName) {
  const range = `${tabName}!A1:${columnLetter(SHEET_HEADERS.length)}1`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  const existing = (res.data.values || [])[0];
  if (!existing || existing.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      requestBody: { values: [SHEET_HEADERS] },
    });
  }
}

function columnLetter(n) {
  // Converts 1-based column index to A, B, ..., Z, AA, AB, ...
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

// ── POST /submit ──────────────────────────────────────────────────────────────
app.post('/submit', async (req, res) => {
  try {
    const body    = req.body;
    const missing = validateBody(body);
    if (missing.length > 0) {
      return res.status(400).json({ error: `Missing required fields: ${missing.join(', ')}` });
    }

    const sheetId  = process.env.GOOGLE_SHEET_ID;
    const sheets   = getSheetsClient();
    const tabName  = await resolveTabName(sheets, sheetId);

    if (!sheetId) {
      return res.status(500).json({ error: 'GOOGLE_SHEET_ID env var is not set' });
    }

    await ensureHeaders(sheets, sheetId, tabName);

    const row = buildRow(body);
    await sheets.spreadsheets.values.append({
      spreadsheetId:   sheetId,
      range:           `${tabName}!A1`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] },
    });

    return res.json({ success: true });

  } catch (err) {
    const message = err?.response?.data?.error?.message || err?.message || String(err);
    console.error('Submit error:', message);
    console.error('Full error:', JSON.stringify(err?.response?.data || err, null, 2));
    return res.status(500).json({ error: message });
  }
});

// ── Health check ──────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ status: 'ok' }));

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
