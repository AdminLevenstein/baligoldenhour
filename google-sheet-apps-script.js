// Golden Hour Club: saves each website request to this Google Sheet AND emails it to admin@levenstein.net.
// Paste into: Google Sheet > Extensions > Apps Script (replace everything), Save,
// then Deploy > Manage deployments > pencil > Version: New version > Deploy (the URL stays the same).

const SHEET_NAME = 'Requests';
const NOTIFY = 'admin@levenstein.net';
const SUBJECT = 'Bali Sunset Hour inquiry';
const HEADERS = ['Timestamp', 'Name', 'WhatsApp', 'Email', 'Company', 'Role', 'Industry', 'Where',
  'Best night', 'Interests', 'Experts', 'Ideas', 'Suggested guest', 'Consent', 'Status', 'Notes'];
// 'Best night' added 23 Oct 2026 (which weeknight members prefer). If the sheet already has rows,
// the script inserts the missing column in the right place on the next submission — nothing to do by hand.

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    if (p._honey) return done_();               // spam bot filled the hidden field
    const now = new Date();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F4E3C1');
    }
    // Make sure every header exists, inserting any new column where HEADERS puts it.
    var have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    HEADERS.forEach(function (h, i) {
      if (have.indexOf(h) !== -1) return;
      var at = i < have.length ? i + 1 : have.length + 1;
      sh.insertColumnBefore(at);
      sh.getRange(1, at).setValue(h).setFontWeight('bold').setBackground('#F4E3C1');
      have.splice(at - 1, 0, h);
    });
    const val = function (h) { return String(p[h] || '').slice(0, 2000); };
    sh.appendRow(have.map(function (h) {
      if (h === 'Timestamp') return now;
      if (h === 'Status') return 'New';
      if (h === 'Notes') return '';
      const v = val(h);
      return v ? "'" + v : '';                  // keep leading 0 in phone numbers, block formulas
    }));
    const when = Utilities.formatDate(now, 'Asia/Makassar', 'd MMM yyyy, HH:mm') + ' (Bali time)';
    const lines = ['Submitted: ' + when].concat(
      HEADERS.slice(1, 14).map(function (h) { return h + ': ' + (val(h) || '-'); }));
    MailApp.sendEmail({
      to: NOTIFY,
      subject: SUBJECT,
      body: lines.join('\n') + '\n\nAlso saved in the Google Sheet "Golden Hour Club — Requests".',
      replyTo: /@/.test(val('Email')) ? val('Email') : NOTIFY,
      name: 'Golden Hour Club website'
    });
    return done_();
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return done_(); }
function done_() { return ContentService.createTextOutput('ok'); }

// Run once from the editor (select testEmail > Run) to approve email permission and check the email arrives.
function testEmail() {
  MailApp.sendEmail(NOTIFY, SUBJECT + ' (test)', 'Test from the Golden Hour Club form script.');
}
