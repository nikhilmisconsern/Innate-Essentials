// =============================================================
//  KRA / KPI DASHBOARD — Google Apps Script (Code.gs)
//  Deploy as Web App: Execute as "Me", Access "Anyone"
//
//  USERS SHEET (tab named "Users"):
//  Row 1  : email | password | role | name | sheetTab  ← headers
//  Row 2+ : one user per row
//
//  EMPLOYEE SHEET FORMAT:
//  Row 1  : "Q1" | "Week Range" | KPI titles (merged across 3 cols each)
//  Row 2  : "Week No" | "From" | "To" | "Planned" "Achieved" "Score" × N | "Total Score"
//  Row 3–15: Weekly data
//  Row 16  : blank separator
//  Row 17  : "WTD" header
//  Row 18  : WTD sub-headers
//  Row 19  : WTD data
// =============================================================

const USERS_SHEET_NAME = 'Users';

// =============================================================
//  doGet — Router  (supports both JSON and JSONP via ?callback=)
// =============================================================
function doGet(e) {
  const callback = e.parameter.callback; // JSONP support for file:// origins

  const out = obj => {
    const json = JSON.stringify(obj);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── List all visible sheet tabs (admin auto-discovery) ──
    if (e.parameter.list === '1') {
      const tabs = ss.getSheets()
        .filter(s => !s.isSheetHidden() && s.getName() !== USERS_SHEET_NAME)
        .map(s => s.getName());
      return out({ success: true, tabs });
    }

    // ── Login: validate credentials from Users sheet ────────
    if (e.parameter.action === 'login') {
      const email    = String(e.parameter.email || '').trim().toLowerCase();
      const password = String(e.parameter.password || '');
      return out(validateLogin(ss, email, password));
    }

    // ── Change password ─────────────────────────────────────
    if (e.parameter.action === 'changePassword') {
      const email       = String(e.parameter.email       || '').trim().toLowerCase();
      const oldPassword = String(e.parameter.oldPassword || '');
      const newPassword = String(e.parameter.newPassword || '');
      return out(changePassword(ss, email, oldPassword, newPassword));
    }

    // ── Active quarter — read from the "Active Quarter" cell in Users sheet ──
    if (e.parameter.action === 'getQuarterSetting') {
      return out({ success: true, quarter: getQuarterOverride(ss) });
    }

    // ── Admin: list all users (for building the team tab list) ──
    if (e.parameter.action === 'getUsers') {
      const adminUser = String(e.parameter.adminUser || '').trim().toLowerCase();
      const adminPass = String(e.parameter.adminPass || '');
      return out(getUsersList(ss, adminUser, adminPass));
    }

    // ── Fetch employee dashboard data ───────────────────────
    const tabName = (e.parameter.employee || '').trim();
    if (!tabName) return out({ success: false, error: 'No employee specified' });

    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return out({ success: false, error: 'Sheet tab not found: ' + tabName });

    return out(parseSheet(sheet, tabName));

  } catch (err) {
    return out({ success: false, error: err.message + ' | ' + err.stack });
  }
}

// =============================================================
//  VALIDATE LOGIN — checks Users sheet
// =============================================================
function validateLogin(ss, email, password) {
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return { success: false, error: 'Users sheet not found. Create a tab named "Users".' };

  if (!isValidEmail(email)) return { success: false, error: 'Please enter a valid email address.' };

  const data = sheet.getDataRange().getValues();
  // Row 0 = headers; data starts at row 1
  for (let r = 1; r < data.length; r++) {
    const row       = data[r];
    const rowEmail  = String(row[0] || '').trim().toLowerCase();
    const pw        = String(row[1] || '');
    if (rowEmail === email && pw === password) {
      return {
        success: true,
        user: {
          id:       rowEmail,
          pw:       pw,
          role:     String(row[2] || 'employee').trim().toLowerCase(),
          name:     String(row[3] || rowEmail),
          sheetTab: String(row[4] || '')
        }
      };
    }
  }
  return { success: false, error: 'Incorrect email or password.' };
}

// =============================================================
//  CHANGE PASSWORD — writes new password back to Users sheet
// =============================================================
function changePassword(ss, email, oldPassword, newPassword) {
  if (!newPassword || newPassword.length < 6)
    return { success: false, error: 'New password must be at least 6 characters.' };

  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return { success: false, error: 'Users sheet not found.' };

  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][0] || '').trim().toLowerCase();
    const pw       = String(data[r][1] || '');
    if (rowEmail === email) {
      if (pw !== oldPassword)
        return { success: false, error: 'Current password is incorrect.' };
      sheet.getRange(r + 1, 2).setValue(newPassword); // column B = password
      SpreadsheetApp.flush();
      return { success: true, message: 'Password changed successfully.' };
    }
  }
  return { success: false, error: 'User not found.' };
}

// =============================================================
//  ACTIVE QUARTER — controlled directly from the "Users" sheet.
//  Add a header "Active Quarter" anywhere in row 1 (e.g. column G) and type
//  Q1 / Q2 / Q3 / Q4 in the cell below it (row 2) to set the quarter that
//  new visits should land on. Leave it blank to follow today's date.
// =============================================================
const QUARTER_HEADER_TEXT = 'active quarter';
const VALID_QUARTERS      = ['Q1', 'Q2', 'Q3', 'Q4'];

function getQuarterOverride(ss) {
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return null;
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let col = -1;
  for (let c = 0; c < headerRow.length; c++) {
    if (String(headerRow[c]||'').trim().toLowerCase() === QUARTER_HEADER_TEXT) { col = c; break; }
  }
  if (col === -1) return null;
  const raw = String(sheet.getRange(2, col+1).getValue() || '').trim().toUpperCase();
  return VALID_QUARTERS.includes(raw) ? raw : null;
}

function getCurrentFiscalQuarter() {
  const m = new Date().getMonth(); // 0=Jan ... 11=Dec
  if (m >= 3 && m <= 5)  return 'Q1'; // Apr,May,Jun
  if (m >= 6 && m <= 8)  return 'Q2'; // Jul,Aug,Sep
  if (m >= 9 && m <= 11) return 'Q3'; // Oct,Nov,Dec
  return 'Q4';                        // Jan,Feb,Mar
}

// =============================================================
//  LIST USERS — admin only (credentials checked against Users sheet)
// =============================================================
function getUsersList(ss, adminUser, adminPass) {
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return { success: false, error: 'Users sheet not found.' };

  const data = sheet.getDataRange().getValues();
  let isAdmin = false;
  const users = [];

  for (let r = 1; r < data.length; r++) {
    const email    = String(data[r][0] || '').trim().toLowerCase();
    const pw       = String(data[r][1] || '');
    const role     = String(data[r][2] || 'employee').trim().toLowerCase();
    const name     = String(data[r][3] || email);
    const sheetTab = String(data[r][4] || '');
    if (!email) continue;
    if (email === adminUser && pw === adminPass && role === 'admin') isAdmin = true;
    users.push({ email, role, name, sheetTab });
  }

  if (!isAdmin) return { success: false, error: 'Not authorized.' };
  return { success: true, users };
}

// =============================================================
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// =============================================================
//  SHEET PARSER
//  Each employee tab holds up to 4 quarter blocks stacked vertically,
//  each starting with a row whose first cell is "Q1" / "Q2" / "Q3" / "Q4".
// =============================================================
function parseSheet(sheet, tabName) {
  const range    = sheet.getDataRange();
  const vals     = range.getValues();
  const dispVals = range.getDisplayValues();
  const numRows  = vals.length;
  const numCols  = vals[0].length;

  // ── Locate every quarter block start (row whose col A is Q1/Q2/Q3/Q4) ──
  const starts = [];
  for (let r = 0; r < numRows; r++) {
    const cell0 = String(vals[r][0] || '').trim().toUpperCase();
    if (VALID_QUARTERS.includes(cell0)) starts.push({ label: cell0, row: r });
  }
  // Fallback: sheet has no Q1–Q4 markers — treat the whole sheet as one block
  if (!starts.length) starts.push({ label: getCurrentFiscalQuarter(), row: 0 });

  const quarters   = {};
  let   kpiDefsRef = null; // KPI column layout — assumed consistent across quarters

  starts.forEach((s, i) => {
    const endRow = (i + 1 < starts.length) ? starts[i+1].row : numRows;
    const block  = parseQuarterBlock(vals, dispVals, s.row, endRow, numCols);
    quarters[s.label] = { weeks: block.weeks, summary: block.summary };
    if (!kpiDefsRef && block.kpiDefs.length) kpiDefsRef = block.kpiDefs;
  });

  const kpiList = (kpiDefsRef || []).map(k => ({ k: k.key, n: k.name, u: guessUnit(k.name) }));

  return {
    success: true,
    employeeName: tabName,
    department: guessDept(kpiDefsRef || []),
    year: '2026-2027',
    quarters,
    kpiList,
    lastUpdated: new Date().toISOString()
  };
}

// Parses a single Q1/Q2/Q3/Q4 block: [startRow .. endRow)
function parseQuarterBlock(vals, dispVals, startRow, endRow, numCols) {
  // ── Find the "Planned / Achieved" sub-header row within this block ──
  let subRowIdx = -1, titleRowIdx = -1;
  for (let r = startRow; r < Math.min(startRow + 5, endRow); r++) {
    const rowStr = vals[r].map(c => String(c).trim().toLowerCase()).join('|');
    if (rowStr.includes('planned') && rowStr.includes('achieved')) {
      subRowIdx   = r;
      titleRowIdx = r - 1 >= startRow ? r - 1 : startRow;
      break;
    }
  }
  if (subRowIdx === -1) { titleRowIdx = startRow; subRowIdx = startRow + 1; }

  const titleRow = vals[titleRowIdx] || [];
  const subRow   = vals[subRowIdx]   || [];

  // ── Detect KPI column groups ──────────────────────────────
  const kpiDefs = [];
  let lastTitle = '';
  for (let c = 0; c < numCols; c++) {
    const t = String(titleRow[c] || '').trim();
    if (t && !['q1','q2','q3','q4','week range',''].includes(t.toLowerCase())) lastTitle = t;
    const s = String(subRow[c] || '').trim().toLowerCase();
    if (s === 'planned' || s === 'planned(wtd)') {
      const rawName = lastTitle || ('KPI_' + (kpiDefs.length + 1));
      kpiDefs.push({ name: cleanName(rawName), key: makeKey(rawName, kpiDefs), colP: c, colA: c + 1, colS: c + 2 });
    }
  }

  // ── Locate Week No / From / To / Total columns ────────────
  let colWeek = -1, colFrom = -1, colTo = -1, colTotal = -1;
  for (let c = 0; c < numCols; c++) {
    const s = String(subRow[c] || '').trim().toLowerCase();
    if (s === 'week no' || s === 'week') colWeek  = c;
    else if (s === 'from')               colFrom  = c;
    else if (s === 'to')                 colTo    = c;
    else if (s === 'score' && c > (kpiDefs[kpiDefs.length - 1]?.colS || 0)) colTotal = c;
  }
  if (colTotal === -1) colTotal = numCols - 1;

  // ── Parse weekly rows ─────────────────────────────────────
  const weeks = [];
  let wtdStartRow = -1;
  const dataStart = subRowIdx + 1;

  for (let r = dataStart; r < endRow; r++) {
    const row   = vals[r];
    const cell0 = String(row[0] || '').trim().toUpperCase();
    if (cell0 === 'WTD') { wtdStartRow = r; break; }

    const weekLabel = colWeek >= 0 ? String(row[colWeek] || '').trim() : String(row[0] || '').trim();
    if (!weekLabel || !weekLabel.toLowerCase().startsWith('week')) continue;
    if (weekLabel.toLowerCase().includes('13') || weekLabel.toLowerCase().includes('buffer')) continue;

    const fromDate = colFrom >= 0 ? fmtDate(row[colFrom]) : '';
    const toDate   = colTo   >= 0 ? fmtDate(row[colTo])   : '';

    const kpis = kpiDefs.map(k => ({
      n: k.name, k: k.key,
      p: toNum(row[k.colP]), a: toNum(row[k.colA]), s: toNum(row[k.colS]),
      pd: fmtDisplay(dispVals[r][k.colP]),
      ad: fmtDisplay(dispVals[r][k.colA]),
      u: guessUnit(k.name)
    }));

    let ws = toNum(row[colTotal]);
    if (ws === 0 && kpis.length) ws = +kpis.reduce((s, k) => s + k.s, 0).toFixed(1);

    if (!kpis.some(k => k.a !== 0 || k.p !== 0)) continue;

    const wnum = weekLabel.replace(/[^0-9]/g, '');
    weeks.push({ id: 'W' + wnum.padStart(2, '0'), l: weekLabel, f: fromDate, t: toDate, kpis, ws });
  }

  // ── Parse WTD summary ─────────────────────────────────────
  const summary = {};

  if (wtdStartRow !== -1) {
    for (let r = wtdStartRow; r < Math.min(wtdStartRow + 6, endRow); r++) {
      const row = vals[r];
      const numericCount = kpiDefs.filter(k =>
        !isNaN(parseFloat(row[k.colA])) && String(row[k.colA]).trim() !== ''
      ).length;
      if (numericCount >= Math.max(1, Math.floor(kpiDefs.length / 2))) {
        kpiDefs.forEach(k => {
          summary['WTD_' + k.key + '_P']  = toNum(row[k.colP]);
          summary['WTD_' + k.key + '_A']  = toNum(row[k.colA]);
          summary['WTD_' + k.key + '_S']  = toNum(row[k.colS]);
          summary['WTD_' + k.key + '_PD'] = fmtDisplay(dispVals[r][k.colP]);
          summary['WTD_' + k.key + '_AD'] = fmtDisplay(dispVals[r][k.colA]);
        });
        summary.WTD_Total = toNum(row[colTotal]);
        break;
      }
    }
  }

  return { weeks, summary, kpiDefs };
}

// =============================================================
//  HELPERS
// =============================================================
function fmtDisplay(v) { const s = String(v || '').trim(); return s === '' ? '0' : s; }
function cleanName(raw) { return raw.replace(/\(.*?\)/g, '').replace(/\s+/g, ' ').trim(); }
function makeKey(raw, existing) {
  let key = raw.toLowerCase().replace(/\(.*?\)/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/, '').substring(0, 18);
  const used = new Set(existing.map(k => k.key));
  let final = key, i = 2;
  while (used.has(final)) final = key + '_' + (i++);
  return final;
}
function guessUnit(name) {
  const n = name.toLowerCase();
  if (n.includes('cvr') || n.includes('repeat') || n.includes('rate') || n.includes('%')) return '%';
  if (n.includes('troas') || n.includes('roi') || n.includes('roas')) return 'x';
  return '';
}
function guessDept(kpiDefs) {
  const all = kpiDefs.map(k => k.name.toLowerCase()).join(' ');
  if (all.includes('reach') || all.includes('cvr') || all.includes('troas')) return 'Marketing';
  if (all.includes('sales growth') || all.includes('paid orders'))            return 'Sales';
  if (all.includes('static') || all.includes('video') || all.includes('reel')) return 'Creative';
  return 'Operations';
}
function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'string' && v.endsWith('%')) return parseFloat(v) || 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const dd = String(v.getDate()).padStart(2, '0');
    const mm = String(v.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + v.getFullYear();
  }
  return String(v);
}
