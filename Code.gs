/**
 * CSE 65-C Portal API
 * Login: Student ID + Gmail only. No shared/fixed password.
 *
 * First successful login:
 *   - Student ID must already exist in the Students sheet.
 *   - If the student's Gmail cell is blank, the submitted Gmail is saved.
 * Later logins:
 *   - Student ID + Gmail must match the Students sheet database.
 *
 * Recommended: keep this Apps Script bound to the Google Sheet used by the portal.
 * If it is standalone, set SPREADSHEET_ID below.
 */

const SPREADSHEET_ID = ''; // Leave blank for a sheet-bound script.

const SHEET_ALIASES = {
  students: ['Students', 'Student'],
  routine: ['Class Routine', 'Routine', 'ClassRoutine'],
  courses: ['Courses', 'Course'],
  notices: ['Notices', 'Notice'],
  loginRequests: ['Login Requests', 'LoginRequests', 'Login Request']
};

const GMAIL_HEADER = 'GMAIL';
const GMAIL_ALIASES = ['gmail', 'email', 'emailaddress', 'studentemail', 'studentgmail'];
const GMAIL_REGISTERED_AT_HEADER = 'GMAIL_REGISTERED_AT';

function doGet(e) {
  try {
    const ss = getSpreadsheet_();
    const payload = {};

    addPublicSheet_(payload, ss, SHEET_ALIASES.students, true);
    addPublicSheet_(payload, ss, SHEET_ALIASES.routine, false);
    addPublicSheet_(payload, ss, SHEET_ALIASES.courses, false);
    addPublicSheet_(payload, ss, SHEET_ALIASES.notices, false);

    return json_({ ok: true, ...payload });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = String(body.action || '').trim();

    switch (action) {
      case 'authenticateEmail':
        return json_(authenticateEmail_(body));
      case 'loginRequest':
        return json_(saveLoginRequest_(body));
      default:
        return json_({ ok: false, message: 'Unknown API action. Please refresh the portal and try again.' });
    }
  } catch (err) {
    return json_({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}

function authenticateEmail_(body) {
  const studentId = normalizeStudentId_(body.studentId);
  const submittedEmail = normalizeEmail_(body.email);

  if (!studentId) return { ok: false, message: 'Student ID is required.' };
  if (!isValidGmail_(submittedEmail)) {
    return { ok: false, message: 'Please enter a valid Gmail address ending with @gmail.com.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ctx = getStudentContext_(studentId, true);
    if (!ctx) return { ok: false, message: 'Student ID not found in the Students sheet.' };

    const storedRaw = ctx.getAny(GMAIL_ALIASES);
    const storedEmail = normalizeEmail_(storedRaw);
    let firstRegistration = false;

    if (!storedEmail) {
      ctx.setAny(GMAIL_HEADER, GMAIL_ALIASES, submittedEmail);
      ctx.setAny(GMAIL_REGISTERED_AT_HEADER, ['gmailregisteredat'], new Date());
      SpreadsheetApp.flush();
      firstRegistration = true;
    } else if (storedEmail !== submittedEmail) {
      return {
        ok: false,
        message: 'This Gmail does not match the Gmail registered for this Student ID.'
      };
    }

    return {
      ok: true,
      firstRegistration,
      student: {
        id: studentId,
        name: ctx.name,
        email: storedEmail || submittedEmail
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function saveLoginRequest_(body) {
  const ss = getSpreadsheet_();
  let sheet = findSheetByAliases_(ss, SHEET_ALIASES.loginRequests);
  if (!sheet) sheet = ss.insertSheet('Login Requests');

  const wantedHeaders = ['Timestamp', 'Student ID', 'Name', 'Gmail', 'Photo', 'Status', 'Activate'];
  wantedHeaders.forEach(h => ensureHeader_(sheet, h));

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = headerMap_(headers);
  const row = new Array(sheet.getLastColumn()).fill('');

  const setCell = (aliases, value) => {
    const idx = findHeaderIndex_(map, aliases);
    if (idx >= 0) row[idx] = value;
  };

  const photo = String(body.photo || '');
  const safePhoto = photo.length > 45000
    ? '[Photo kept only on student device - too large for a Sheet cell]'
    : photo;

  setCell(['timestamp'], body.timestamp ? new Date(body.timestamp) : new Date());
  setCell(['studentid'], normalizeStudentId_(body.studentId));
  setCell(['name'], String(body.name || ''));
  setCell(['gmail', 'email'], normalizeEmail_(body.email));
  setCell(['photo'], safePhoto);
  setCell(['status'], String(body.status || 'Logged In'));
  setCell(['activate'], String(body.activate || 'Active'));

  sheet.appendRow(row);
  return { ok: true };
}

function addPublicSheet_(payload, ss, aliases, isStudents) {
  const sheet = findSheetByAliases_(ss, aliases);
  if (!sheet) return;
  payload[sheet.getName()] = isStudents ? readPublicStudents_(sheet) : readSheetObjects_(sheet);
}

/**
 * Public Students payload deliberately does NOT return Gmail values.
 * The browser only receives ID/name/badge/type. Gmail matching happens server-side.
 */
function readPublicStudents_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];

  const headers = values[0];
  const map = headerMap_(headers);
  const idCol = findHeaderIndex_(map, ['id', 'studentid']);
  const nameCol = findHeaderIndex_(map, ['name', 'studentname']);
  const badgeCol = findHeaderIndex_(map, ['badge', 'badage']);
  const typeCol = findHeaderIndex_(map, ['type', 'studenttype']);

  if (idCol < 0 || nameCol < 0) throw new Error('Students sheet must have ID and Name headers.');

  const rows = [];
  const seen = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = normalizeStudentId_(row[idCol]);
    const name = String(row[nameCol] || '').trim();
    if (!id || !name || isHeaderLikeStudent_(id, name) || seen[id]) continue;
    seen[id] = true;

    rows.push({
      id,
      name,
      badge: badgeCol >= 0 ? String(row[badgeCol] || '').trim() : '',
      type: typeCol >= 0 ? String(row[typeCol] || '').trim() : ''
    });
  }
  return rows;
}

function readSheetObjects_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];

  const headers = values[0].map((h, i) => String(h || '').trim() || ('Column' + (i + 1)));
  return values.slice(1)
    .filter(row => row.some(v => String(v || '').trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] !== undefined ? row[i] : '');
      return obj;
    });
}

function getStudentContext_(studentId, ensureGmailColumns) {
  const ss = getSpreadsheet_();
  const sheet = findSheetByAliases_(ss, SHEET_ALIASES.students);
  if (!sheet) throw new Error('Students sheet not found.');
  if (sheet.getLastRow() < 2) return null;

  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  let map = headerMap_(headers);

  if (ensureGmailColumns) {
    if (findHeaderIndex_(map, GMAIL_ALIASES) < 0) ensureHeader_(sheet, GMAIL_HEADER);
    if (findHeaderIndex_(map, ['gmailregisteredat']) < 0) ensureHeader_(sheet, GMAIL_REGISTERED_AT_HEADER);
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    map = headerMap_(headers);
  }

  const idCol = findHeaderIndex_(map, ['id', 'studentid']);
  const nameCol = findHeaderIndex_(map, ['name', 'studentname']);
  if (idCol < 0 || nameCol < 0) throw new Error('Students sheet must have ID and Name headers.');

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeStudentId_(values[i][idCol]) !== studentId) continue;

    const rowNumber = i + 2;
    return {
      sheet,
      rowNumber,
      headers,
      map,
      name: String(values[i][nameCol] || '').trim(),
      refreshMap: function() {
        this.headers = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getDisplayValues()[0];
        this.map = headerMap_(this.headers);
      },
      getAny: function(aliases) {
        const idx = findHeaderIndex_(this.map, aliases);
        return idx >= 0 ? this.sheet.getRange(this.rowNumber, idx + 1).getValue() : '';
      },
      setAny: function(preferredHeader, aliases, value) {
        let idx = findHeaderIndex_(this.map, aliases);
        if (idx < 0) {
          ensureHeader_(this.sheet, preferredHeader);
          this.refreshMap();
          idx = findHeaderIndex_(this.map, [preferredHeader]);
        }
        this.sheet.getRange(this.rowNumber, idx + 1).setValue(value);
      }
    };
  }
  return null;
}

function ensureHeader_(sheet, headerName) {
  if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1).setValue(headerName);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const wanted = normalizeKey_(headerName);
  if (headers.some(h => normalizeKey_(h) === wanted)) return;
  sheet.getRange(1, lastCol + 1).setValue(headerName);
}

function findSheetByAliases_(ss, aliases) {
  const wanted = aliases.map(normalizeKey_);
  return ss.getSheets().find(sh => wanted.indexOf(normalizeKey_(sh.getName())) !== -1) || null;
}

function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = normalizeKey_(h);
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

function findHeaderIndex_(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizeKey_(aliases[i]);
    if (map[key] !== undefined) return map[key];
  }
  return -1;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Bind this script to the portal spreadsheet or set SPREADSHEET_ID.');
  return ss;
}

function parseBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  try { return JSON.parse(raw); }
  catch (err) { throw new Error('Invalid JSON request.'); }
}

function normalizeStudentId_(value) {
  const digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 9) return digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6);
  return String(value || '').trim();
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidGmail_(value) {
  return /^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@gmail\.com$/i.test(normalizeEmail_(value));
}

function normalizeKey_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function isHeaderLikeStudent_(id, name) {
  const a = normalizeKey_(id);
  const b = normalizeKey_(name);
  return (a === 'id' && b === 'name') || a === 'studentid';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Admin helper. If a student typed the wrong Gmail during first registration,
 * run this manually in Apps Script, then the student can register a Gmail again.
 * Example: adminResetStudentGmail('261-115-111');
 */
function adminResetStudentGmail(studentId) {
  const id = normalizeStudentId_(studentId);
  const ctx = getStudentContext_(id, true);
  if (!ctx) throw new Error('Student ID not found: ' + id);
  ctx.setAny(GMAIL_HEADER, GMAIL_ALIASES, '');
  ctx.setAny(GMAIL_REGISTERED_AT_HEADER, ['gmailregisteredat'], '');
  SpreadsheetApp.flush();
  return 'Gmail registration cleared for ' + id + '.';
}
