/**
 * ADD THIS to your existing Google Apps Script (the one bound to your sheet).
 *
 * 1) Put the 3 lines at the top of your existing doGet(e), so lookups are
 *    answered with JSON and everything else behaves as before:
 *
 *      function doGet(e) {
 *        if (e && e.parameter && e.parameter.format === 'json') {
 *          return lookupAttendeeJson_(e);
 *        }
 *        // ... your existing doGet code stays here ...
 *      }
 *
 *    (If you have no doGet at all, use the one below as-is.)
 *
 * 2) Paste lookupAttendeeJson_ anywhere in the script.
 *
 * 3) Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *    The web app URL stays the same. "Who has access" must be "Anyone".
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.format === 'json') {
    return lookupAttendeeJson_(e);
  }
  return ContentService.createTextOutput('OK');
}

function lookupAttendeeJson_(e) {
  var p = e.parameter || {};
  var token = String(p.token || p.t || p.id || p.qr || p.code || p.ticket || '').trim();
  if (!token) {
    var keys = Object.keys(p).filter(function (k) { return k !== 'format'; });
    if (keys.length) token = String(p[keys[0]]).trim();
  }

  var out = { success: false };
  if (token) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rows = sheet.getDataRange().getValues();
    var header = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var nameCol = header.findIndex(function (h) { return /name/.test(h); });
    var emailCol = header.findIndex(function (h) { return /mail/.test(h); });

    for (var i = 1; i < rows.length; i++) {
      var hit = rows[i].some(function (cell) { return String(cell).trim() === token; });
      if (hit) {
        out = {
          success: true,
          name: nameCol >= 0 ? String(rows[i][nameCol]).trim() : '',
          email: emailCol >= 0 ? String(rows[i][emailCol]).trim() : ''
        };
        break;
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
