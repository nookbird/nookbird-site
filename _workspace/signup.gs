/**
 * NookBird sign-ups and applications: the waitlist, the testers' applications, and the owner's tracking of both.
 *
 * A Google Apps Script web app in NookBird's own Google Workspace (the hello@ account). The website's boxes post here;
 * nothing reaches Google before the family presses Join, Apply or Send. This script:
 *   - keeps one row per person in the private Sheet's `People` tab (only the owner can open it);
 *   - sends the approved thank-you email once per address, from hello@nookbird.com;
 *   - takes a tester's answers only with the one-time code made at Apply (it keeps a fingerprint, never the code);
 *   - emails the owner at each tester's answers, and a summary every morning at 08:00 Melbourne time;
 *   - keeps a separate `Counts` Sheet file with numbers only (no names, no emails), for operations;
 *   - ignores robots that fill the hidden field, and anything that isn't an email address.
 *
 * Setup (the owner, once): see README.md beside this file. Script properties:
 *   SHEET_ID   - the private Sheet's id (from its address: /spreadsheets/d/<SHEET_ID>/edit)
 *   COUNTS_ID  - the Counts Sheet file's id (a separate file, shared with operations)
 *   FROM       - hello@nookbird.com (must be a "Send mail as" address of the deploying account)
 *   OWNER_TO   - the owner's own address, for the two notices
 *   TESTERS    - "on" to accept tester applications; anything else refuses them
 */

var PAGE_URL = 'https://www.nookbird.com/pilot.html';
var ZONE = 'Australia/Melbourne';
var DAILY_EMAILS = 100;   // at most this many thank-you emails a day, whatever is sent here
var PLACES_ON_OUR_KEY = 20;

// An ordinary address only: letters, digits and . _ % + - before the @ (never = + - @ first, so the Sheet can never read
// it as a formula), a plain domain after it.
var EMAIL_RE = /^[A-Za-z0-9_][A-Za-z0-9._%+-]{0,63}@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;
var CODE_RE = /^[0-9a-f]{32}$/;
var PAGES = { home: true, testers: true, pilot: true };

// The People tab: one row per person. Columns 15 to 18 are the owner's; the script never overwrites them, except the
// stage moving from "waitlist" to "applied" when a waitlist family applies.
var COLS = ['When joined (UTC)', 'Kind', 'Email', 'Name', 'State', 'Children', 'Windows 11', 'Administrator',
  'Usually on', 'Smart App Control', 'How heard', 'Who told', 'Fits?', 'Answered (UTC)', 'Stage', 'Tester code',
  'Kind of place', 'Notes', 'Email sent', 'Page', 'Code fingerprint'];
var C = {}; COLS.forEach(function (name, i) { C[name] = i + 1; });
var STAGES = ['waitlist', 'applied', 'place offered', 'accepted', 'ready to install', 'active', 'paused', 'left'];
var PLACE_KINDS = ['free', 'paying', 'own key'];
var PLACE_STAGES = { 'place offered': true, 'accepted': true, 'ready to install': true, 'active': true, 'paused': true };

// The questions: each answer must be one of its choices (or blank). Name and "who told" are short free text.
var CHOICES = {
  state: ['ACT', 'NSW', 'NT', 'QLD', 'SA', 'TAS', 'VIC', 'WA'],
  windows: ['Yes', 'No', 'Not sure'],
  on: ['Yes, most of the time', 'Sometimes', 'Rarely'],
  admin: ['Yes', 'No', 'Not sure'],
  sac: ['On', 'Evaluation', 'Off', 'Not sure'],
  heard: ['A friend or family member', 'Our school or kinder community', 'Facebook or other social media',
    'An internet search', 'Other']
};

var WAITLIST_EMAIL = {
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
    'The NookBird team',
    'www.nookbird.com'
  ].join('\n')
};

/** The testers' thank-you (marketing's part B, approved by the owner 17 Nov 20:23), with the link to the questions. */
function testerEmail(link) {
  return {
    subject: 'Your NookBird tester application',
    body: [
      'Hi,',
      '',
      'Thank you for applying to test NookBird in Term 1, 2027.',
      '',
      'If you haven\'t answered our few quick questions yet, you can do it here: ' + link + ' It takes about two minutes.',
      '',
      'We\'ll email you within a few days to let you know whether we have a place for your family.',
      '',
      'What testing means. Testers use NookBird through the term as part of normal family life and tell us honestly what ' +
        'works and what doesn\'t. In return you get NookBird early, help shape it, and active testers will be offered a ' +
        'discount on their first year.',
      '',
      'If you have a question, just reply to this email.',
      '',
      'The NookBird team',
      'www.nookbird.com'
    ].join('\n')
  };
}

// ---------------------------------------------------------------------------------------------------------------------
// The web app

function doPost(e) {
  var p = (e && e.parameter) || {};
  var kind = String(p.kind || 'waitlist');
  var email = String(p.email || '').trim().toLowerCase();
  var robot = String(p.website || '') !== '';            // the hidden field: a person never fills it
  var props = PropertiesService.getScriptProperties();
  var ok = !robot && email.length <= 254 && EMAIL_RE.test(email);
  var outcome = 'ignored', code = null;
  if (ok && (kind === 'waitlist' || ((kind === 'tester' || kind === 'answers') && props.getProperty('TESTERS') === 'on'))) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sheet = peopleSheet(props);
      if (kind === 'waitlist') outcome = join(sheet, props, email, p);
      else if (kind === 'tester') {
        code = CODE_RE.test(String(p.code || '')) ? String(p.code) : newCode();
        outcome = apply(sheet, props, email, code, p);
      } else outcome = answer(sheet, props, email, String(p.code || ''), p);
      if (outcome !== 'ignored' && outcome !== 'already') rebuildCounts(sheet, props);
    } finally {
      lock.releaseLock();
    }
  }
  // The same answer for a robot as for anyone else. Without JavaScript, this page is what the family sees.
  if (kind === 'tester') return page(outcome === 'applied' ? thanksAndQuestions(email, code) : appliedBefore());
  if (kind === 'answers') return page(outcome === 'answered' ? '<h1>Thank you.</h1><p>We\'ll email you within a few days to let you know whether we have a place for your family.</p>'
    : outcome === 'already' ? '<h1>Your answers are already in. Thank you.</h1>'
    : '<h1>Thank you.</h1><p>If you applied and this was your first time answering, we have your answers. If not, just reply to our email.</p>');
  return page('<h1>Thank you. Check your inbox.</h1>');
}

/** ?status=1&e=<email>&c=<code>: whether this application's answers are in (for the page opened from the email's link). */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.status) {
    var answered = false;
    var email = String(p.e || '').trim().toLowerCase(), code = String(p.c || '');
    if (EMAIL_RE.test(email) && CODE_RE.test(code)) {
      var sheet = peopleSheet(PropertiesService.getScriptProperties());
      var row = rowOf(sheet, email);
      if (row && cell(sheet, row, 'Code fingerprint') === fingerprint(code) && cell(sheet, row, 'Answered (UTC)') !== '') answered = true;
    }
    return ContentService.createTextOutput(JSON.stringify({ answered: answered })).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><p>NookBird sign-ups.</p>');
}

function join(sheet, props, email, p) {
  if (rowOf(sheet, email)) return 'ignored';
  var sent = sendOnce(props, email, WAITLIST_EMAIL);
  appendRow(sheet, { 'Kind': 'waitlist', 'Email': email, 'Stage': 'waitlist', 'Email sent': sent, 'Page': pageOf(p) });
  return 'joined';
}

function apply(sheet, props, email, code, p) {
  var row = rowOf(sheet, email);
  if (row && cell(sheet, row, 'Kind') === 'tester') return 'ignored';   // one application per address, one code
  var mail = testerEmail(linkFor(email, code));
  var sent = sendOnce(props, email, mail);
  if (row) {
    // A waitlist family applies: the same row becomes an application.
    setCells(sheet, row, { 'Kind': 'tester', 'Email sent': sent, 'Page': pageOf(p), 'Code fingerprint': fingerprint(code) });
    var stage = cell(sheet, row, 'Stage');
    if (stage === '' || stage === 'waitlist') setCells(sheet, row, { 'Stage': 'applied' });
  } else {
    appendRow(sheet, { 'Kind': 'tester', 'Email': email, 'Stage': 'applied', 'Email sent': sent, 'Page': pageOf(p),
      'Code fingerprint': fingerprint(code) });
  }
  return 'applied';
}

function answer(sheet, props, email, code, p) {
  var row = rowOf(sheet, email);
  if (!row || !CODE_RE.test(code) || cell(sheet, row, 'Kind') !== 'tester') return 'ignored';
  if (cell(sheet, row, 'Code fingerprint') !== fingerprint(code)) return 'ignored';
  if (cell(sheet, row, 'Answered (UTC)') !== '') return 'already';
  var a = {
    'Name': text(p.name, 80),
    'Children': children(p.children),
    'State': choice(p.state, CHOICES.state),
    'Windows 11': choice(p.windows, CHOICES.windows),
    'Usually on': choice(p.on, CHOICES.on),
    'Administrator': choice(p.admin, CHOICES.admin),
    'Smart App Control': choice(p.sac, CHOICES.sac),
    'How heard': choice(p.heard, CHOICES.heard)
  };
  a['Who told'] = a['How heard'] === CHOICES.heard[0] ? text(p.who, 80) : '';
  a['Fits?'] = fits(a);
  a['Answered (UTC)'] = new Date().toISOString();
  setCells(sheet, row, a);
  noticeOwner(props, sheet, row, a);
  return 'answered';
}

/** yes: Windows 11 Yes, administrator Yes, Smart App Control Off. not yet: Smart App Control On, or not Windows 11. */
function fits(a) {
  if (a['Smart App Control'] === 'On' || a['Windows 11'] === 'No') return 'not yet';
  if (a['Windows 11'] === 'Yes' && a['Administrator'] === 'Yes' && a['Smart App Control'] === 'Off') return 'yes';
  return 'check';
}

// ---------------------------------------------------------------------------------------------------------------------
// The two notices

function noticeOwner(props, sheet, row, a) {
  var to = props.getProperty('OWNER_TO');
  if (!to) return;
  var subject = 'New tester application: ' + (a['State'] || 'state not given') + ', ' +
    (a['Windows 11'] === 'Yes' ? 'Windows 11' : 'Windows 11: ' + (a['Windows 11'] || 'not given')) + ', ' +
    'Smart App Control ' + (a['Smart App Control'] || 'not given') + ', ' + a['Fits?'];
  var body = [
    'Fits? ' + a['Fits?'],
    'State: ' + (a['State'] || '-'),
    'Children: ' + (a['Children'] || '-'),
    'Windows 11: ' + (a['Windows 11'] || '-'),
    'Administrator: ' + (a['Administrator'] || '-'),
    'Usually on: ' + (a['Usually on'] || '-'),
    'Smart App Control: ' + (a['Smart App Control'] || '-'),
    'How heard: ' + (a['How heard'] || '-') + (a['Who told'] ? ' (' + a['Who told'] + ')' : ''),
    '',
    'The row: ' + rowLink(props, sheet, row)
  ].join('\n');
  GmailApp.sendEmail(to, subject, body, { from: props.getProperty('FROM'), name: 'NookBird sign-ups' });
}

/** Run by the daily trigger at 08:00 Melbourne time (installDaily sets it up). */
function dailySummary() {
  var props = PropertiesService.getScriptProperties();
  var to = props.getProperty('OWNER_TO');
  var sheet = peopleSheet(props);
  var rows = allRows(sheet);
  var now = Date.now(), day = 24 * 3600 * 1000;
  var newWaitlist = 0, newApps = 0, newAnswers = 0, newFit = 0, noAnswerDay = 0, waiting3 = 0;
  rows.forEach(function (r) {
    var joined = Date.parse(r['When joined (UTC)']) || 0, answered = Date.parse(r['Answered (UTC)']) || 0;
    if (r['Kind'] === 'waitlist' && now - joined < day) newWaitlist++;
    if (r['Kind'] === 'tester' && now - joined < day) newApps++;
    if (answered && now - answered < day) { newAnswers++; if (r['Fits?'] === 'yes') newFit++; }
    if (r['Stage'] === 'applied' && !answered && now - joined > day) noAnswerDay++;
    if (r['Stage'] === 'applied' && now - joined > 3 * day) waiting3++;
  });
  rebuildCounts(sheet, props);
  if (!to) return;
  var body = [
    'The last 24 hours:',
    '- new waitlist sign-ups: ' + newWaitlist,
    '- new tester applications: ' + newApps,
    '- answers received: ' + newAnswers + ' (' + newFit + ' fit)',
    '',
    'Waiting:',
    '- applied, no answers after a day: ' + noAnswerDay,
    '- applied more than 3 days ago and still waiting for our reply: ' + waiting3,
    '',
    'The Sheet: https://docs.google.com/spreadsheets/d/' + props.getProperty('SHEET_ID') + '/edit'
  ].join('\n');
  GmailApp.sendEmail(to, 'NookBird sign-ups: the morning summary', body, { from: props.getProperty('FROM'), name: 'NookBird sign-ups' });
}

/** Run once by hand from the editor: makes the 08:00 Melbourne trigger for dailySummary (and removes any older one). */
function installDaily() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailySummary').timeBased().atHour(8).everyDays(1).inTimezone(ZONE).create();
}

// ---------------------------------------------------------------------------------------------------------------------
// The Counts file: numbers only, never a name or an email

function rebuildCounts(sheet, props) {
  var id = props.getProperty('COUNTS_ID');
  if (!id) return;
  var rows = allRows(sheet);
  var weeks = {}, byState = {}, bySac = {}, byFits = {}, byStage = {}, byHeard = {}, places = 0;
  rows.forEach(function (r) {
    var w = weekOf(r['When joined (UTC)']);
    if (w) {
      weeks[w] = weeks[w] || [0, 0, 0, 0];
      if (r['Kind'] === 'waitlist') weeks[w][0]++;
      else { weeks[w][1]++; if (r['Answered (UTC)'] !== '') weeks[w][2]++; else weeks[w][3]++; }
    }
    if (r['Kind'] === 'tester') {
      bump(byState, r['State'] || 'not given');
      if (r['Answered (UTC)'] !== '') { bump(bySac, r['Smart App Control'] || 'not given'); bump(byFits, r['Fits?']); bump(byHeard, r['How heard'] || 'not given'); }
    }
    bump(byStage, r['Stage'] || 'not given');
    if (PLACE_STAGES[r['Stage']] && (r['Kind of place'] === 'free' || r['Kind of place'] === 'paying')) places++;
  });
  var out = [['NookBird sign-ups: counts only (no names, no emails)', ''],
    ['Updated (Melbourne)', Utilities.formatDate(new Date(), ZONE, 'yyyy-MM-dd HH:mm')], ['', ''],
    ['Places used on NookBird\'s AI key', places + ' of ' + PLACES_ON_OUR_KEY], ['', '']];
  var block = function (title, obj, order) {
    out.push([title, '']);
    (order || Object.keys(obj).sort()).forEach(function (k) { if (obj[k]) out.push([k, obj[k]]); });
    out.push(['', '']);
  };
  block('Applications by state', byState);
  block('Smart App Control (answered)', bySac);
  block('Fits? (answered)', byFits, ['yes', 'check', 'not yet']);
  block('How they heard (answered)', byHeard, CHOICES.heard.concat(['not given']));
  block('By stage', byStage, STAGES.concat(['not given']));
  var table = [['Week starting (Monday)', 'Waitlist sign-ups', 'Tester applications', 'Applied and answered', 'Applied, no answers yet']];
  Object.keys(weeks).sort().forEach(function (w) { table.push([w].concat(weeks[w])); });
  var counts = SpreadsheetApp.openById(id).getSheets()[0];
  counts.clear();
  counts.getRange(1, 1, out.length, 2).setValues(out);
  counts.getRange(1, 4, table.length, 5).setValues(table);
}

// ---------------------------------------------------------------------------------------------------------------------
// The Sheet

function peopleSheet(props) {
  var book = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  var sheet = book.getSheetByName('People');
  if (!sheet) {
    sheet = book.insertSheet('People', 0);
    sheet.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    var rest = sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLS.length);
    rest.setNumberFormat('@');
    sheet.getRange(2, C['Stage'], sheet.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STAGES, true).setAllowInvalid(false).build());
    sheet.getRange(2, C['Kind of place'], sheet.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(PLACE_KINDS, true).setAllowInvalid(false).build());
    sheet.hideColumns(C['Code fingerprint']);
  }
  return sheet;
}

function appendRow(sheet, values) {
  var row = sheet.getLastRow() + 1;
  var line = COLS.map(function (name) { return name === 'When joined (UTC)' ? new Date().toISOString() : (values[name] || ''); });
  var range = sheet.getRange(row, 1, 1, COLS.length);
  range.setNumberFormat('@');   // every cell plain text, never a formula
  range.setValues([line]);
}

function setCells(sheet, row, values) {
  Object.keys(values).forEach(function (name) {
    var r = sheet.getRange(row, C[name]);
    r.setNumberFormat('@');
    r.setValue(values[name]);
  });
}

function cell(sheet, row, name) { return String(sheet.getRange(row, C[name]).getValue()); }

function rowOf(sheet, email) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var values = sheet.getRange(2, C['Email'], last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) if (String(values[i][0]).toLowerCase() === email) return i + 2;
  return 0;
}

function allRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, COLS.length).getValues().map(function (v) {
    var r = {}; COLS.forEach(function (name, i) { r[name] = String(v[i]); }); return r;
  });
}

function rowLink(props, sheet, row) {
  return 'https://docs.google.com/spreadsheets/d/' + props.getProperty('SHEET_ID') + '/edit#gid=' + sheet.getSheetId() + '&range=A' + row;
}

/**
 * Run once by hand from the editor, after the first deployment of this version: copies the rows of the old `Waitlist`
 * and `Testers` tabs into `People` (skipping any address already there), then renames the old tabs "... (old)".
 * Nothing is deleted.
 */
function moveOldTabs() {
  var props = PropertiesService.getScriptProperties();
  var book = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  var sheet = peopleSheet(props);
  ['Waitlist', 'Testers'].forEach(function (tab) {
    var old = book.getSheetByName(tab);
    if (!old) return;
    var last = old.getLastRow();
    var values = last < 2 ? [] : old.getRange(2, 1, last - 1, Math.min(5, old.getLastColumn())).getValues();
    values.forEach(function (v) {
      var email = String(v[1] || '').trim().toLowerCase();
      if (!EMAIL_RE.test(email) || rowOf(sheet, email)) return;
      var when = v[0] instanceof Date ? v[0].toISOString() : String(v[0]);
      var kind = String(v[2]) === 'tester' ? 'tester' : 'waitlist';
      appendRow(sheet, { 'Kind': kind, 'Email': email, 'Stage': kind === 'tester' ? 'applied' : 'waitlist',
        'Email sent': String(v[4] || 'yes'), 'Page': String(v[3] || '') });
      sheet.getRange(sheet.getLastRow(), C['When joined (UTC)']).setValue(when);
    });
    old.setName(tab + ' (old)');
  });
  rebuildCounts(sheet, props);
}

// ---------------------------------------------------------------------------------------------------------------------
// Small pieces

function sendOnce(props, email, mail) {
  if (!underDailyCap(props)) return 'held (daily cap)';
  GmailApp.sendEmail(email, mail.subject, mail.body, { from: props.getProperty('FROM'), name: 'NookBird', replyTo: props.getProperty('FROM') });
  return 'yes';
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

function newCode() { return Utilities.getUuid().replace(/-/g, '').toLowerCase(); }

function fingerprint(code) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, code, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); }).join('');
}

function linkFor(email, code) { return PAGE_URL + '?e=' + encodeURIComponent(email) + '&c=' + code + '#questions'; }

function pageOf(p) { return PAGES[String(p.page || '')] ? String(p.page) : 'other'; }

function choice(v, list) { v = String(v || ''); return list.indexOf(v) >= 0 ? v : ''; }

/** Short free text, on one line, never starting with a character a Sheet could read as a formula. */
function text(v, max) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, max).replace(/^[=+\-@]+/, '').trim();
}

function children(v) {
  var s = String(v || '').trim();
  if (!/^\d{1,2}$/.test(s)) return '';
  var n = Number(s);
  return n <= 20 ? String(n) : '';
}

function weekOf(iso) {
  var t = Date.parse(iso);
  if (!t) return '';
  var d = new Date(t);
  var dow = Number(Utilities.formatDate(d, ZONE, 'u'));   // 1 = Monday
  return Utilities.formatDate(new Date(t - (dow - 1) * 24 * 3600 * 1000), ZONE, 'yyyy-MM-dd');
}

function bump(obj, key) { obj[key] = (obj[key] || 0) + 1; }

function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

function page(inner) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>NookBird</title>' +
    '<body style="font-family:system-ui,sans-serif;max-width:36rem;margin:3rem auto;padding:0 1rem;color:#1F2A27;background:#FBF7F1;line-height:1.5">' +
    inner + '<p><a href="https://www.nookbird.com/" style="color:#2F5D50">Back to NookBird</a></p></body>');
}

function appliedBefore() {
  return '<h1>Thank you.</h1><p>If this address is new to us, check your inbox for our email. If you applied before, ' +
    'the link to our few quick questions is in the email we sent you then.</p>';
}

/** Without JavaScript: the thank-you, and the questions with this application's code, posted back here on Send. */
function thanksAndQuestions(email, code) {
  var radios = function (name, list) {
    return list.map(function (v) {
      return '<label style="display:block"><input type="radio" name="' + name + '" value="' + esc(v) + '"> ' + esc(v) + '</label>';
    }).join('');
  };
  var q = function (label, body, small) {
    return '<fieldset style="border:0;padding:0;margin:0 0 1.2rem"><legend style="font-weight:600">' + label + '</legend>' + body +
      (small ? '<small style="color:#5b6b66">' + small + '</small>' : '') + '</fieldset>';
  };
  return '<h1>Thank you, you\'re on our testers list.</h1>' +
    '<p>Now let\'s get to know your family a little, so we can check NookBird is ready for your home.</p>' +
    '<form method="post" action="' + ScriptApp.getService().getUrl() + '">' +
    '<input type="hidden" name="kind" value="answers"><input type="hidden" name="email" value="' + esc(email) + '">' +
    '<input type="hidden" name="code" value="' + esc(code) + '">' +
    q('Your name', '<input name="name" maxlength="80" autocomplete="name">') +
    q('How many children do you have at school or kinder? (optional)', '<input name="children" type="number" min="0" max="20">') +
    q('Which state are you in?', '<select name="state"><option value=""></option>' + CHOICES.state.map(function (s) { return '<option>' + s + '</option>'; }).join('') + '</select>') +
    q('Is your computer Windows 11?', radios('windows', CHOICES.windows), 'NookBird runs on Windows 11 only for now. To check: Start → Settings → System → About, and look for "Edition".') +
    q('Will it usually be on?', radios('on', CHOICES.on), 'It works best left on, so reminders and briefs arrive on time, but it doesn\'t have to be on all the time.') +
    q('Are you the administrator of this computer, or do you have an administrator account on it?', radios('admin', CHOICES.admin), 'Installing NookBird needs an administrator\'s approval.') +
    q('In Windows Security, what does Smart App Control say?', radios('sac', CHOICES.sac), 'This setting decides whether our installer can run on your computer today. To check: Start → Windows Security → App & browser control → Smart App Control settings.') +
    q('How did you hear about us? (optional)', radios('heard', CHOICES.heard)) +
    q('If a friend or family member: who told you? (optional)', '<input name="who" maxlength="80">', 'We like to thank the family who told you.') +
    '<button type="submit" style="padding:.7rem 1.6rem;border:0;border-radius:999px;background:#2F5D50;color:#fff;font:inherit">Send</button>' +
    '<p><small>Your answers go to us only when you press Send.</small></p></form>';
}

/** Run once by hand from the editor to check the setup: it sends the waitlist email to FROM itself and adds no row. */
function testEmail() {
  var props = PropertiesService.getScriptProperties();
  var to = props.getProperty('FROM');
  GmailApp.sendEmail(to, '[test] ' + WAITLIST_EMAIL.subject, WAITLIST_EMAIL.body, { from: to, name: 'NookBird', replyTo: to });
  SpreadsheetApp.openById(props.getProperty('SHEET_ID'));   // fails here if SHEET_ID is wrong
}
