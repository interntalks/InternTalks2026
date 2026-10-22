/**
 * EVENT CHECK-IN BACKEND  (Google Apps Script, bound to your attendee sheet)
 *
 * What it does
 *  - action=list     -> returns every attendee (token, name, email, checkedIn) so all
 *                       devices load the same list straight from the sheet
 *  - action=checkin  -> marks one attendee "Present" (with time + device) and tells the
 *                       app if that person was ALREADY checked in on another device
 *
 * Columns are found by header name in row 1 (any order):
 *   token : header containing "token", "qr", "ticket", "code" or exactly "id"
 *   name  : header(s) containing "name"      email : header containing "mail"
 * Three columns are added automatically if missing:
 *   "Checked In", "Check-in Time", "Checked In By"
 *
 * INSTALL
 *  1. Sheet > Extensions > Apps Script. Paste this whole file (keep your QR-generation code).
 *  2. If you already have a doGet(e), delete THIS doGet and add these 3 lines at the top of yours:
 *       var p = (e && e.parameter) || {};
 *       if (p.action || p.format === 'json') return json_(handleApi_(p));
 *  3. Deploy > Manage deployments > pencil > Version: New version > Deploy.
 *     (Execute as: Me, Who has access: Anyone.  The URL stays the same.)
 */

var SHEET_NAME = ''; // leave empty to use the first sheet, or put your tab name, e.g. 'Attendees'

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action || p.format === 'json') return json_(handleApi_(p));
  return ContentService.createTextOutput('Event check-in API is running');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApi_(p) {
  try {
    var action = p.action || 'lookup';
    if (action === 'list') return apiList_();
    if (action === 'checkin') return apiCheckin_(p);
    if (action === 'lookup') return apiLookup_(p);
    return { success: false, error: 'Unknown action' };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

function getCols_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  var low = header.map(function (h) { return h.toLowerCase(); });

  function findFirst(re) {
    for (var i = 0; i < low.length; i++) if (re.test(low[i])) return i;
    return -1;
  }

  var token = findFirst(/token/);
  if (token < 0) token = findFirst(/qr/);
  if (token < 0) token = findFirst(/ticket|^code$/);
  if (token < 0) token = findFirst(/^id$/);
  if (token < 0) token = 0;

  var nameCols = [];
  for (var i = 0; i < low.length; i++) {
    if (/name/.test(low[i]) && i !== token) nameCols.push(i);
  }

  var cols = {
    token: token,
    nameCols: nameCols,
    email: findFirst(/mail/),
    present: low.indexOf('checked in'),
    time: low.indexOf('check-in time'),
    by: low.indexOf('checked in by')
  };

  var next = header.length;
  [['present', 'Checked In'], ['time', 'Check-in Time'], ['by', 'Checked In By']].forEach(function (c) {
    if (cols[c[0]] < 0) {
      sheet.getRange(1, next + 1).setValue(c[1]);
      cols[c[0]] = next;
      next++;
    }
  });
  return cols;
}

function nameOf_(row, cols) {
  return cols.nameCols.map(function (i) { return String(row[i] || '').trim(); })
    .filter(Boolean).join(' ');
}

function emailOf_(row, cols) {
  return cols.email >= 0 ? String(row[cols.email] || '').trim() : '';
}

function toMs_(v) {
  if (!v) return 0;
  if (v instanceof Date) return v.getTime();
  var t = Date.parse(v);
  return isNaN(t) ? 0 : t;
}

function findRow_(data, cols, token) {
  var t = String(token || '').trim();
  if (!t) return -1;
  var tl = t.toLowerCase();
  var i;
  for (i = 1; i < data.length; i++) if (String(data[i][cols.token]).trim() === t) return i;
  for (i = 1; i < data.length; i++) if (String(data[i][cols.token]).trim().toLowerCase() === tl) return i;
  for (i = 1; i < data.length; i++) {
    for (var c = 0; c < data[i].length; c++) {
      if (String(data[i][c]).trim().toLowerCase() === tl) return i;
    }
  }
  return -1;
}

function apiList_() {
  var sheet = getSheet_();
  var cols = getCols_(sheet);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var tok = String(r[cols.token] || '').trim();
    if (!tok) continue;
    out.push({
      token: tok,
      name: nameOf_(r, cols),
      email: emailOf_(r, cols),
      checkedIn: String(r[cols.present] || '').trim() !== '',
      ts: toMs_(r[cols.time])
    });
  }
  return { success: true, attendees: out, total: out.length };
}

function apiLookup_(p) {
  var sheet = getSheet_();
  var cols = getCols_(sheet);
  var data = sheet.getDataRange().getValues();
  var token = String(p.token || p.t || p.id || p.qr || p.code || p.ticket || '').trim();
  var row = findRow_(data, cols, token);
  if (row < 0) return { success: false };
  return { success: true, name: nameOf_(data[row], cols), email: emailOf_(data[row], cols) };
}

function apiCheckin_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet_();
    var cols = getCols_(sheet);
    var data = sheet.getDataRange().getValues();
    var row = findRow_(data, cols, p.token);
    if (row < 0) return { success: true, status: 'notfound' };

    var r = data[row];
    var info = { name: nameOf_(r, cols), email: emailOf_(r, cols) };

    if (String(r[cols.present] || '').trim() !== '') {
      return { success: true, status: 'duplicate', name: info.name, email: info.email, ts: toMs_(r[cols.time]) };
    }

    var when = p.ts && !isNaN(Number(p.ts)) ? new Date(Number(p.ts)) : new Date();
    sheet.getRange(row + 1, cols.present + 1).setValue('Present');
    sheet.getRange(row + 1, cols.time + 1).setValue(when).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(row + 1, cols.by + 1).setValue(String(p.device || ''));
    SpreadsheetApp.flush();
    return { success: true, status: 'ok', name: info.name, email: info.email, ts: when.getTime() };
  } finally {
    lock.releaseLock();
  }
}
