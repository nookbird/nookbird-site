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
    // The approved words (the owner, to marketing, 17 Nov 15:50). TESTERS stays off until the page's box goes live.
    body: [
      'Hi,',
      '',
      'Thank you for applying to test NookBird in Term 1, 2027.',
      '',
      'Places are limited, so we have a few quick questions first. Each one tells us something we need to give your family ' +
        'a good start; where the reason isn\'t obvious, we\'ve said why. We never share your answers. Just reply to this ' +
        'email; it takes about five minutes.',
      '',
      'ABOUT YOUR FAMILY',
      '1. Your first name, and how many children you have at school or kinder.',
      '2. Which state you\'re in.',
      '',
      'ABOUT YOUR COMPUTER (NookBird runs on your own computer at home)',
      '3. Is it Windows 11? (NookBird runs on Windows 11 only for now. To check: Start > Settings > System > About, and ' +
        'look for "Edition" under Windows specifications.)',
      '4. Will it usually be on? (It works best left on, so reminders and briefs arrive on time, but it doesn\'t have to ' +
        'be on all the time.)',
      '5. Are you the administrator of the computer you\'d use, or do you have an administrator account on it? ' +
        '(Installing NookBird needs an administrator\'s approval.)',
      '6. In Windows Security, what does Smart App Control say: On, Evaluation or Off? (This setting decides whether our ' +
        'installer can run on your computer today. To check: Start > Windows Security > App & browser control > Smart ' +
        'App Control settings.)',
      '',
      'AND ONE LAST THING',
      '7. Who told you about NookBird? (Optional. NookBird grows by families telling families, and we like to thank the ' +
        'family who told you.)',
      '',
      'What testing means. Testers use NookBird through the term as part of normal family life and tell us honestly what ' +
        'works and what doesn\'t. In return you get NookBird early, help shape it, and active testers will be offered a ' +
        'discount on their first year.',
      '',
      'We\'ll reply within a few days to let you know whether we have a place for your family.',
      '',
      'NookBird',
      'www.nookbird.com'
    ].join('\n')
  }
};

// An ordinary address only: letters, digits and . _ % + - before the @ (never = + - @ first, so the Sheet can never read
// it as a formula), a plain domain after it.
var EMAIL_RE = /^[A-Za-z0-9_][A-Za-z0-9._%+-]{0,63}@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;
var PAGES = { home: true, testers: true, pilot: true };
var DAILY_EMAILS = 100;   // at most this many thank-you emails a day, whatever is sent here

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
        var page = PAGES[String(p.page || '')] ? String(p.page) : 'other';
        // Every cell is written as plain text, never a formula.
        var range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 5);
        range.setNumberFormat('@');
        var sent = underDailyCap(props) ? 'yes' : 'held (daily cap)';
        range.setValues([[new Date().toISOString(), email, kind, page, sent]]);
        if (sent === 'yes') GmailApp.sendEmail(email, spec.subject, spec.body, { from: props.getProperty('FROM'), name: 'NookBird', replyTo: props.getProperty('FROM') });
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
    sheet.appendRow(['When (UTC)', 'Email', 'Kind', 'Page', 'Email sent']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** A count of today's emails in the script's own properties: past DAILY_EMAILS, the row is kept and no email is sent. */
function underDailyCap(props) {
  var today = new Date().toISOString().slice(0, 10);
  var day = props.getProperty('CAP_DAY'), n = Number(props.getProperty('CAP_COUNT') || '0');
  if (day !== today) { day = today; n = 0; }
  if (n >= DAILY_EMAILS) return false;
  props.setProperties({ CAP_DAY: day, CAP_COUNT: String(n + 1) });
  return true;
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
