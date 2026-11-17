/**
 * NookBird sign-ups: the waitlist (now) and the Term 1 testers (later).
 *
 * A Google Apps Script web app in NookBird's own Google Workspace (the hello@ account). The website's email box posts
 * here; nothing reaches Google before the family presses Join. This script:
 *   - keeps the address and the time in a private Sheet (one tab per kind), only the owner can open;
 *   - sends the approved thank-you email once per address, from hello@nookbird.com;
 *   - ignores robots that fill the hidden field, and anything that isn't an email address.
 *
 * Setup (the owner, once): see README.md beside this file. Script properties:
 *   SHEET_ID  - the Sheet's id (from its address: /spreadsheets/d/<SHEET_ID>/edit)
 *   FROM      - hello@nookbird.com (must be a "Send mail as" address of the deploying account)
 *   TESTERS   - "on" to accept tester sign-ups; anything else refuses them (off until the testers page goes live)
 */

var KINDS = {
  waitlist: {
    tab: 'Waitlist',
    subject: "You're on the NookBird list",
    body: [
      'Hi,',
      '',
      'Thank you for joining the NookBird waitlist.',
      '',
      'What happens next. A small group of families starts using NookBird in Term 1, 2027. A few places are free, and the ' +
        'rest are at a reduced price. We\'ll email you as places open, in the order families joined, and friends of NookBird ' +
        'families are among the first we ask.',
      '',
      'What you\'ll need. A Windows 11 computer at home, ideally left on, and Telegram on each family member\'s phone. ' +
        'Because NookBird is a private app, it runs on your own computer, never on our servers. For now it supports ' +
        'Windows 11 computers only; we plan to add support for more. If that doesn\'t fit your home yet, reply and tell us; ' +
        'we\'d like to know.',
      '',
      'Who told you about NookBird? Just reply and tell us; we like to thank the family who did.',
      '',
      'Know another family who\'d like it? Forward them www.nookbird.com.',
      '',
      'If you have a question, just reply to this email.',
      '',
      'NookBird',
      'www.nookbird.com'
    ].join('\n')
  },
  tester: {
    tab: 'Testers',
    subject: 'Your NookBird tester application',
    // The approved words (marketing, 17 Nov 15:50) go here when the testers page is built; TESTERS stays off until then.
    body: null
  }
};

var EMAIL_RE = /^[^\s@<>()",;:]{1,64}@[A-Za-z0-9.-]{1,253}\.[A-Za-z]{2,}$/;

function doPost(e) {
  var p = (e && e.parameter) || {};
  var kind = String(p.kind || 'waitlist');
  var email = String(p.email || '').trim().toLowerCase();
  var robot = String(p.website || '') !== '';            // the hidden field: a person never fills it
  var spec = KINDS[kind];
  var props = PropertiesService.getScriptProperties();
  if (kind === 'tester' && props.getProperty('TESTERS') !== 'on') spec = null;
  if (!robot && spec && spec.body && email.length <= 254 && EMAIL_RE.test(email)) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sheet = sheetFor(props.getProperty('SHEET_ID'), spec.tab);
      if (!alreadyThere(sheet, email)) {
        sheet.appendRow([new Date(), email, kind, String(p.page || '').slice(0, 40)]);
        GmailApp.sendEmail(email, spec.subject, spec.body, { from: props.getProperty('FROM'), name: 'NookBird', replyTo: props.getProperty('FROM') });
      }
    } finally {
      lock.releaseLock();
    }
  }
  // The same answer whatever happened: a robot learns nothing, and a family that joined twice gets no second email.
  return HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8"><title>NookBird</title>' +
    '<body style="font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem;color:#1F2A27;background:#FBF7F1">' +
    '<h1 style="font-size:1.5rem">Thank you. Check your inbox.</h1>' +
    '<p><a href="https://www.nookbird.com/" style="color:#2F5D50">Back to NookBird</a></p></body>');
}

function doGet() {
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><p>NookBird sign-ups.</p>');
}

function sheetFor(id, tab) {
  var book = SpreadsheetApp.openById(id);
  var sheet = book.getSheetByName(tab);
  if (!sheet) {
    sheet = book.insertSheet(tab);
    sheet.appendRow(['When', 'Email', 'Kind', 'Page']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function alreadyThere(sheet, email) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var values = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) if (String(values[i][0]).toLowerCase() === email) return true;
  return false;
}

/** Run once by hand from the editor to check the setup: it sends the waitlist email to FROM itself and adds no row. */
function testEmail() {
  var props = PropertiesService.getScriptProperties();
  var to = props.getProperty('FROM');
  GmailApp.sendEmail(to, '[test] ' + KINDS.waitlist.subject, KINDS.waitlist.body, { from: to, name: 'NookBird', replyTo: to });
  SpreadsheetApp.openById(props.getProperty('SHEET_ID'));   // fails here if SHEET_ID is wrong
}
