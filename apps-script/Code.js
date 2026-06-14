/**
 * Togro waitlist + email automation
 * ---------------------------------------------------------------
 * - Website form POSTs here (name, email, user type) -> row in "Waitlist" tab.
 * - On signup, the subscriber gets the "welcome" email immediately.
 * - A daily trigger sends that day's countdown email to the whole list.
 * - ALL email content lives in the editable "Emails" tab:
 *      Active | When | Subject | Body | Button text | Button URL | Sent on
 *   * Active   = TRUE/FALSE  (untick to cancel/disable an email)
 *   * When     = the date to send, or the word "welcome" (sent on signup)
 *   * Body     = plain text; tokens {{first_name}} {{name}} {{days}} {{launch}}
 *   * Sent on  = filled automatically once a scheduled email has gone out
 *
 * Run setup() once from the editor to create the sheet + seed the emails +
 * install the daily trigger + grant permissions.
 */

var LAUNCH        = new Date(2026, 5, 22);          // 22 June 2026
var SHEET_TITLE   = 'Togro Waitlist';
var WAITLIST_TAB  = 'Waitlist';
var EMAILS_TAB    = 'Emails';
var WAITLIST_HEAD = ['Timestamp', 'Name', 'Email', 'User type', 'Source'];
var EMAILS_HEAD   = ['Active', 'When', 'Subject', 'Body', 'Button text', 'Button URL', 'Sent on'];
var SEND_HOUR     = 9;                               // daily send time (local)
var LOGO_URL      = 'https://cdn.jsdelivr.net/gh/teamtogro/togro-site-v3@main/images/logo.png';
var SITE_URL      = 'https://togro.co';

/* ───────────────────────── sheet plumbing ───────────────────────── */

function ss_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.create(SHEET_TITLE);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function waitlistSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(WAITLIST_TAB) || ss.getSheets()[0];
  sh.setName(WAITLIST_TAB);
  if (sh.getLastRow() === 0) {
    sh.appendRow(WAITLIST_HEAD);
    sh.getRange(1, 1, 1, WAITLIST_HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function emailsSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(EMAILS_TAB);
  if (!sh) sh = ss.insertSheet(EMAILS_TAB);
  if (sh.getLastRow() === 0) {
    sh.appendRow(EMAILS_HEAD);
    sh.getRange(1, 1, 1, EMAILS_HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() <= 1) {       // header only (or empty) -> seed the campaign
    seedEmails_(sh);
    sh.setColumnWidth(3, 320);      // Subject
    sh.setColumnWidth(4, 460);      // Body
  }
  return sh;
}

/* ───────────────────────── seed campaign ───────────────────────── */

function seedEmails_(sh) {
  var D = function (m, d) { return new Date(2026, m, d); };
  var rows = [
    ['TRUE', 'welcome', "You're on the Togro waitlist 🌿",
      "Hi {{first_name}},\n\nYou're on the list. Thanks for joining Togro before we launch.\n\nTogro is the intelligence layer for the countryside: live, verified, community-built reports of what's actually happening on the ground. Not a hiking app. Not a walking tracker.\n\nWe go live on {{launch}}. Over the next few days we'll show you exactly what that means, and on launch day your download link lands right here.\n\nSee you on the trail,\nThe Togro team", '', ''],

    ['TRUE', D(5, 15), "Togro isn't a hiking app ({{days}} days to go)",
      "Hi {{first_name}},\n\nMost outdoor apps tell you how far you walked. Togro tells you what's ahead.\n\nLivestock in the next field. A flooded path. A fallen tree. A closed gate. Reported by the people who just walked it, and verified before you set off.\n\n{{days}} days until launch.", '', ''],

    ['TRUE', D(5, 16), "Know before you go ({{days}} days to go)",
      "Hi {{first_name}},\n\nEvery report on Togro shows its confidence: type, distance, age, and how many walkers have confirmed it. A hazard you can't trust is just noise.\n\nThat's the difference between a map and an intelligence network.\n\n{{days}} days to go.", '', ''],

    ['TRUE', D(5, 17), "Earn as you walk ({{days}} days to go)",
      "Hi {{first_name}},\n\nSpot something, report it, earn credits, then spend them at rural pubs, cafes and farm shops along the way. The more your report is verified, the more it's worth.\n\nIt's a contribution economy, not a step counter.\n\n{{days}} days to go.", '', ''],

    ['TRUE', D(5, 18), "One map, four communities ({{days}} days to go)",
      "Hi {{first_name}},\n\nTogro connects walkers, farmers, rural venues and councils on one map. Walkers get safety and rewards; farmers reduce friction at the gate; venues get verified footfall; councils get path intelligence.\n\nOne trusted network.\n\n{{days}} days to go.", '', ''],

    ['TRUE', D(5, 19), "Mapped by the people who work the land ({{days}} days to go)",
      "Hi {{first_name}},\n\nFarmers post live land updates, like a bull in the west field or a repaired gate, so walkers arrive informed. Fewer gates left open, fewer surprises, safer routes.\n\nThe countryside, mapped by the people who know it.\n\n{{days}} days to go.", '', ''],

    ['TRUE', D(5, 20), "Climb the trust ladder ({{days}} days to go)",
      "Hi {{first_name}},\n\nThe more you contribute, the higher you climb: Scout, Pathfinder, Ranger, Guardian, Sentinel, Pioneer. From explorer to protector of the routes you love.\n\nAlmost there. {{days}} days to go.", '', ''],

    ['TRUE', D(5, 21), "Tomorrow.",
      "Hi {{first_name}},\n\nTomorrow, Togro goes live.\n\nYou were here before launch, so you'll be among the first to read the land like a local. Keep an eye on your inbox in the morning.\n\nSee you out there,\nThe Togro team", '', ''],

    ['TRUE', D(5, 22), "Togro is live: read the land like a local",
      "Hi {{first_name}},\n\nIt's here. Togro is live.\n\nLive countryside intelligence, verified routes, rewards as you walk, all in your pocket. Thank you for being one of the first.\n\nTap below to get the app.\n\nThe Togro team", 'Get the app', 'https://togro.co']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(2, 2, rows.length, 1).setNumberFormat('ddd d mmm');  // When column
}

/* ───────────────────────── setup + triggers ───────────────────────── */

function setup() {
  waitlistSheet_();
  emailsSheet_();
  installTrigger_();
  var url = ss_().getUrl();
  Logger.log('Togro Waitlist ready: ' + url);
  return url;
}

function installTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendScheduledEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendScheduledEmails').timeBased().everyDays(1).atHour(SEND_HOUR).create();
}

/* ───────────────────────── web app ───────────────────────── */

function doGet() {
  return ContentService.createTextOutput('Togro waitlist is live.');
}

/** Wipes the campaign rows and re-writes the latest seed content. Run from the editor. */
function reseedEmails() {
  var sh = emailsSheet_();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, EMAILS_HEAD.length).clearContent();
  seedEmails_(sh);
}

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    var name = (p.name || '').toString().slice(0, 200);
    var email = (p.email || '').toString().slice(0, 200);
    var type = (p.type || '').toString().slice(0, 80);
    waitlistSheet_().appendRow([new Date(), name, email, type, (p.source || 'website').toString().slice(0, 80)]);
    try { sendWelcome_(name, email); } catch (mailErr) { /* never block the signup */ }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── sending ───────────────────────── */

function sendWelcome_(name, email) {
  if (!email) return;
  var row = findEmailRow_(function (r) {
    return String(r.when).toLowerCase() === 'welcome' && isActive_(r.active);
  });
  if (!row) return;
  sendOne_(email, name, row);
}

/** Daily trigger: send today's active, unsent countdown email to everyone. */
function sendScheduledEmails() {
  var sh = emailsSheet_();
  var data = sh.getDataRange().getValues();        // includes header row 1
  var today = startOfDay_(new Date());
  var subs = getSubscribers_();
  for (var i = 1; i < data.length; i++) {
    var active = isActive_(data[i][0]);
    var when = data[i][1];
    var sentOn = data[i][6];
    if (!active) continue;
    if (sentOn) continue;
    if (!(when instanceof Date)) continue;          // skips "welcome"
    if (startOfDay_(when).getTime() !== today.getTime()) continue;  // exact-day match only
    var row = { active: active, when: when, subject: data[i][2], body: data[i][3], btnText: data[i][4], btnUrl: data[i][5] };
    subs.forEach(function (s) { try { sendOne_(s.email, s.name, row); } catch (err) {} });
    sh.getRange(i + 1, 7).setValue(new Date());     // mark Sent on
  }
}

function sendOne_(email, name, row) {
  var data = tokens_(name);
  var subject = render_(row.subject, data);
  var bodyHtml = htmlEmail_(render_(row.body, data), row.btnText, row.btnUrl);
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: bodyHtml, name: 'Togro', replyTo: 'team@togro.co' });
}

/* ───────────────────────── helpers ───────────────────────── */

function getSubscribers_() {
  var data = waitlistSheet_().getDataRange().getValues();
  var seen = {}, out = [];
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][2] || '').trim();
    if (!email || seen[email.toLowerCase()]) continue;
    seen[email.toLowerCase()] = true;
    out.push({ name: String(data[i][1] || ''), email: email });
  }
  return out;
}

function findEmailRow_(pred) {
  var data = emailsSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var r = { active: data[i][0], when: data[i][1], subject: data[i][2], body: data[i][3], btnText: data[i][4], btnUrl: data[i][5] };
    if (pred(r)) return r;
  }
  return null;
}

function isActive_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

function tokens_(name) {
  var first = String(name || '').trim().split(/\s+/)[0] || 'there';
  return {
    name: String(name || '').trim() || 'there',
    first_name: first,
    days: String(daysToLaunch_()),
    launch: Utilities.formatDate(LAUNCH, Session.getScriptTimeZone() || 'Europe/London', 'd MMMM')
  };
}

function render_(tpl, data) {
  return String(tpl == null ? '' : tpl).replace(/\{\{\s*(\w+)\s*\}\}/g, function (m, k) {
    return (data[k] != null) ? data[k] : m;
  });
}

function daysToLaunch_() {
  var d = Math.ceil((startOfDay_(LAUNCH).getTime() - startOfDay_(new Date()).getTime()) / 86400000);
  return d < 0 ? 0 : d;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function htmlEmail_(bodyText, btnText, btnUrl) {
  var FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
  var body = String(bodyText)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');

  var button = '';
  if (btnUrl && String(btnUrl).trim()) {
    button =
      '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:26px 0 4px"><tr>' +
        '<td style="border-radius:999px;background:#C8D91A">' +
          '<a href="' + String(btnUrl).trim() + '" style="display:inline-block;padding:14px 30px;font-family:' + FONT +
          ';font-size:15px;font-weight:700;color:#10140A;text-decoration:none;border-radius:999px">' +
          (String(btnText || 'Open').trim()) + ' &nbsp;&rarr;</a>' +
        '</td>' +
      '</tr></table>';
  }

  return '' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#EAEFF4;margin:0;padding:0">' +
 '<tr><td align="center" style="padding:28px 12px">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;max-width:600px;background:#ffffff;border:1px solid #DCE3EA;border-radius:18px;overflow:hidden">' +
    // header (navy) with logo + wordmark + tagline
    '<tr><td style="background:#111827;padding:22px 28px">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
        '<td style="vertical-align:middle">' +
          '<img src="' + LOGO_URL + '" width="30" height="30" alt="Togro" style="vertical-align:middle;border:0;display:inline-block">' +
          '<span style="font-family:' + FONT + ';color:#C8D91A;font-weight:800;font-size:21px;letter-spacing:-.5px;vertical-align:middle;padding-left:9px">Togro</span>' +
        '</td>' +
        '<td align="right" style="vertical-align:middle">' +
          '<span style="font-family:' + FONT + ';color:#7F8BA1;font-size:10px;font-weight:700;letter-spacing:2px">COUNTRYSIDE&nbsp;INTELLIGENCE</span>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>' +
    // lime accent bar
    '<tr><td style="height:4px;line-height:4px;font-size:0;background:#C8D91A">&nbsp;</td></tr>' +
    // body
    '<tr><td style="padding:32px 30px 28px;font-family:' + FONT + ';color:#0F172A;font-size:15px;line-height:1.65">' +
      body + button +
    '</td></tr>' +
    // footer
    '<tr><td style="padding:20px 30px;background:#F8FAFC;border-top:1px solid #EEF2F6;font-family:' + FONT + ';color:#64748B;font-size:12px;line-height:1.6">' +
      '<span style="color:#0F172A;font-weight:700">Togro</span> &middot; Real-time countryside intelligence<br>' +
      '<a href="' + SITE_URL + '" style="color:#0B7E74;text-decoration:none">togro.co</a> &middot; ' +
      '<a href="mailto:team@togro.co" style="color:#0B7E74;text-decoration:none">team@togro.co</a><br>' +
      '<span style="color:#94A3B8">You\'re receiving this because you joined the Togro waitlist. ' +
        '<a href="mailto:team@togro.co?subject=Unsubscribe" style="color:#94A3B8;text-decoration:underline">Unsubscribe</a>.</span>' +
    '</td></tr>' +
  '</table>' +
 '</td></tr>' +
'</table>';
}
