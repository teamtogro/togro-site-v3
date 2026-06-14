/**
 * Togro waitlist + tailored countdown email automation
 * ---------------------------------------------------------------
 * - Website form POSTs here (name, email, user type) -> row in "Waitlist" tab.
 * - On signup, the subscriber gets the "welcome" email immediately.
 * - A daily trigger sends that day's countdown email, TAILORED to each
 *   subscriber's user type (walker / venue / brand / community / council).
 * - ALL email content lives in the editable "Emails" tab:
 *      Active | When | Audience | Subject | Body | Button text | Button URL | Sent on
 *   * Active   = TRUE/FALSE   (untick to disable a row)
 *   * When     = the date to send, or the word "welcome" (sent on signup)
 *   * Audience = walker | venue | brand | community | council | all
 *   * Body     = plain text; tokens {{first_name}} {{name}} {{days}} {{launch}}.
 *               Lines starting with "• " render as bullets.
 *   * Sent on  = filled automatically once a scheduled row has gone out.
 * - Logo + countdown hero images are embedded (Assets.js) and inlined via CID,
 *   so they always render in Gmail/Outlook with no external hosting.
 *
 * Run setup() once from the editor to (re)create the sheet, seed the emails,
 * install the daily trigger and grant permissions. Run reseedEmails() any time
 * to overwrite the campaign rows with the latest copy below.
 */

var LAUNCH        = new Date(2026, 5, 22);          // 22 June 2026
var SHEET_TITLE   = 'Togro Waitlist';
var WAITLIST_TAB  = 'Waitlist';
var EMAILS_TAB    = 'Emails';
var WAITLIST_HEAD = ['Timestamp', 'Name', 'Email', 'User type', 'Source'];
var EMAILS_HEAD   = ['Active', 'When', 'Audience', 'Subject', 'Body', 'Button text', 'Button URL', 'Sent on'];
var SEND_HOUR     = 9;                               // daily send time (local)
var SITE_URL      = 'https://togro.co';
var AUDIENCES     = ['walker', 'venue', 'brand', 'community', 'council'];
var PARTNERS_TAB  = 'Partners';
var PARTNERS_HEAD = ['Timestamp', 'Type', 'Name', 'Organisation', 'Email', 'Phone', 'Notes', 'Source'];
var NOTIFY_EMAIL  = 'team@togro.co';   // where partner enquiries are sent

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

function partnersSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(PARTNERS_TAB);
  if (!sh) sh = ss.insertSheet(PARTNERS_TAB);
  if (sh.getLastRow() === 0) {
    sh.appendRow(PARTNERS_HEAD);
    sh.getRange(1, 1, 1, PARTNERS_HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 200); sh.setColumnWidth(7, 320);
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
    sh.setColumnWidth(4, 320);      // Subject
    sh.setColumnWidth(5, 520);      // Body
  }
  return sh;
}

/* ───────────────────────── tailored copy ─────────────────────────
   One benefit + three bullets per audience, reused across milestones. */

function audienceCopy_() {
  return {
    walker: {
      benefit: "Live hazards, livestock and diversions on the paths you actually walk — so you know before you set off.",
      bullets: ["Real-time alerts on your routes", "Report what you spot, earn rewards", "Verified by the people who walk it"],
      liveCta: "Open the map", liveFirst: "Open Togro, check your local loop, and log the first thing you spot."
    },
    venue: {
      benefit: "Put your pub or café on the walkers’ map and turn passing footfall into regulars with rewards they redeem in person.",
      bullets: ["Get discovered by nearby walkers", "Reward visits, build loyalty", "Live footfall around your area"],
      liveCta: "Claim your venue", liveFirst: "Claim your venue and set your first walker reward."
    },
    brand: {
      benefit: "Reach a high-intent outdoor community through real-world rewards and partnerships — not banner ads.",
      bullets: ["An engaged, outdoors-first audience", "Reward-led, not interruptive", "Measurable real-world redemption"],
      liveCta: "Start a partnership", liveFirst: "Tell us your goals and we’ll line up your first reward drop."
    },
    community: {
      benefit: "Protect local access. Verify paths, log issues and rally your group around one living map of the ground.",
      bullets: ["Log and verify path issues", "Keep access open, together", "Local knowledge, mapped"],
      liveCta: "Map your patch", liveFirst: "Add your group and verify the paths you know best."
    },
    council: {
      benefit: "Real-time hazard and path intelligence for your area, plus a live footfall dashboard — straight from the ground.",
      bullets: ["Live hazard & diversion data", "Footfall dashboard for your area", "Reported and verified by residents"],
      liveCta: "View the dashboard", liveFirst: "Open your area dashboard and see what’s live today."
    }
  };
}

function audienceLabel_(a) {
  return { walker: 'Walker', venue: 'Venue', brand: 'Brand', community: 'Community group', council: 'Council', all: 'Everyone' }[a] || a;
}

/* ───────────────────────── seed campaign ───────────────────────── */

function seedEmails_(sh) {
  var D = function (m, d) { return new Date(2026, m, d); };
  var AC = audienceCopy_();

  var welcomeBody =
    "Hi {{first_name}},\n\nYou’re on the list. Thanks for joining Togro before we launch.\n\n" +
    "Togro is the intelligence layer for the countryside: live, verified, community-built reports of what’s " +
    "actually happening on the ground. Not a hiking app. Not a walking tracker.\n\n" +
    "We go live on {{launch}}. Over the next week we’ll show you exactly what that means for you — and on " +
    "launch day your link lands right here.\n\nSee you on the trail,\nThe Togro team";

  var rows = [
    ['TRUE', 'welcome', 'all', "You’re on the Togro waitlist 🌿", welcomeBody, '', '', '']
  ];

  // milestone date, hero key, subject, intro line
  var miles = [
    { date: D(5, 15), key: '7',    subj: "One week to go, {{first_name}}",          intro: "Togro goes live on {{launch}}. Here’s what to expect." },
    { date: D(5, 19), key: '3',    subj: "3 days to go, {{first_name}}",            intro: "Almost there — Togro goes live on {{launch}}." },
    { date: D(5, 21), key: '1',    subj: "Tomorrow: Togro goes live",              intro: "This is it — Togro opens {{launch}}." },
    { date: D(5, 22), key: 'live', subj: "Togro is live: read the land like a local", intro: "It’s here. Real-time countryside intelligence, mapped by the people who walk it." }
  ];

  for (var mi = 0; mi < miles.length; mi++) {
    var m = miles[mi];
    for (var ai = 0; ai < AUDIENCES.length; ai++) {
      var a = AUDIENCES[ai], c = AC[a], body, btnText, btnUrl;
      if (m.key === 'live') {
        body = m.intro + "\n\n" + c.benefit + "\n\nYour first move: " + c.liveFirst;
        btnText = c.liveCta; btnUrl = SITE_URL;
      } else {
        body = m.intro + "\n\n" + c.benefit + "\n\n• " + c.bullets.join("\n• ");
        btnText = "See what’s coming"; btnUrl = SITE_URL;
      }
      rows.push(['TRUE', m.date, a, m.subj, body, btnText, btnUrl, '']);
    }
  }

  sh.getRange(2, 1, rows.length, EMAILS_HEAD.length).setValues(rows);
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

/** Wipes the campaign rows and re-writes the latest seed content. Run from the editor. */
function reseedEmails() {
  var sh = emailsSheet_();
  sh.clear();                                                   // drop stale 7-col layout
  sh.getRange(1, 1, 1, EMAILS_HEAD.length).setValues([EMAILS_HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  seedEmails_(sh);
  sh.setColumnWidth(4, 320);
  sh.setColumnWidth(5, 520);
}

/* ───────────────────────── web app ───────────────────────── */

function doGet() {
  return ContentService.createTextOutput('Togro waitlist is live.');
}

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.form === 'partner') return handlePartner_(p);
    var name = (p.name || '').toString().slice(0, 200);
    var email = (p.email || '').toString().slice(0, 200);
    var type = (p.type || '').toString().slice(0, 80);
    waitlistSheet_().appendRow([new Date(), name, email, type, (p.source || 'website').toString().slice(0, 80)]);
    try { sendWelcome_(name, email, type); } catch (mailErr) { /* never block the signup */ }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Partner enquiry (farmer / venue / brand / council): log it + email the team to follow up. */
function handlePartner_(p) {
  var s = function (k, n) { return (p[k] || '').toString().slice(0, n || 200); };
  var type = s('type', 80), name = s('name'), org = s('org'), email = s('email'),
      phone = s('phone', 40), notes = s('notes', 500), source = s('source', 80) || 'website';
  partnersSheet_().appendRow([new Date(), type, name, org, email, phone, notes, source]);
  try {
    var rows = [['Type', type], ['Name', name], ['Organisation', org], ['Email', email], ['Phone', phone], ['Notes', notes]]
      .filter(function (r) { return r[1]; })
      .map(function (r) { return '<tr><td style="padding:4px 14px 4px 0;color:#64748B">' + r[0] + '</td><td style="padding:4px 0;color:#0F172A;font-weight:600">' + r[1] + '</td></tr>'; })
      .join('');
    MailApp.sendEmail({
      to: NOTIFY_EMAIL, name: 'Togro', replyTo: email || NOTIFY_EMAIL,
      subject: 'New partner enquiry: ' + (type || 'partner') + (name ? ' — ' + name : ''),
      htmlBody: '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0F172A">' +
        '<p>New partner enquiry from the website:</p><table>' + rows + '</table>' +
        '<p style="color:#64748B;font-size:12px">Logged in the Partners tab of the Togro Waitlist sheet.</p></div>'
    });
  } catch (mailErr) { /* never block the submission */ }
  return json_({ ok: true });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── sending ───────────────────────── */

function sendWelcome_(name, email, type) {
  if (!email) return;
  var row = findEmailRow_(function (r) {
    return String(r.when).toLowerCase() === 'welcome' && isActive_(r.active) &&
           (audKey_(r.audience) === 'all' || audKey_(r.audience) === audKey_(type));
  });
  if (!row) return;
  sendOne_(email, name, row, null);   // welcome has no hero
}

/** Daily trigger: send today's active, unsent rows, matched to each subscriber's type. */
function sendScheduledEmails() {
  var sh = emailsSheet_();
  var data = sh.getDataRange().getValues();        // includes header row 1
  var today = startOfDay_(new Date());

  // collect today's active, unsent rows, keyed by audience
  var byAud = {}, rowIndexes = [];
  for (var i = 1; i < data.length; i++) {
    if (!isActive_(data[i][0])) continue;
    if (data[i][7]) continue;                       // already sent
    var when = data[i][1];
    if (!(when instanceof Date)) continue;          // skips "welcome"
    if (startOfDay_(when).getTime() !== today.getTime()) continue;
    var row = rowObj_(data[i]);
    byAud[audKey_(row.audience)] = row;
    rowIndexes.push(i + 1);
  }
  if (rowIndexes.length === 0) return;

  var heroKey = heroKeyForDate_(today);
  var subs = getSubscribers_();
  subs.forEach(function (s) {
    var row = byAud[audKey_(s.type)] || byAud['all'] || byAud['walker'];
    if (!row) return;
    try { sendOne_(s.email, s.name, row, heroKey); } catch (err) {}
  });

  rowIndexes.forEach(function (r) { sh.getRange(r, 8).setValue(new Date()); });  // mark Sent on
}

function sendOne_(email, name, row, heroKey) {
  var data = tokens_(name);
  var subject = render_(row.subject, data);
  var hasHero = heroKey && typeof HERO_B64 !== 'undefined' && HERO_B64[heroKey];
  var bodyHtml = htmlEmail_(render_(row.body, data), row.btnText, row.btnUrl, hasHero);

  var inlineImages = {};
  if (typeof LOGO_B64 !== 'undefined' && LOGO_B64) {
    inlineImages.togrologo = Utilities.newBlob(Utilities.base64Decode(LOGO_B64), 'image/png', 'logo.png');
  }
  if (hasHero) {
    inlineImages.hero = Utilities.newBlob(Utilities.base64Decode(HERO_B64[heroKey]), 'image/jpeg', 'hero.jpg');
  }
  MailApp.sendEmail({
    to: email, subject: subject, htmlBody: bodyHtml,
    name: 'Togro', replyTo: 'team@togro.co', inlineImages: inlineImages
  });
}

/* ───────────────────────── helpers ───────────────────────── */

function getSubscribers_() {
  var data = waitlistSheet_().getDataRange().getValues();
  var seen = {}, out = [];
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][2] || '').trim();
    if (!email || seen[email.toLowerCase()]) continue;
    seen[email.toLowerCase()] = true;
    out.push({ name: String(data[i][1] || ''), email: email, type: String(data[i][3] || '') });
  }
  return out;
}

function rowObj_(r) {
  return { active: r[0], when: r[1], audience: r[2], subject: r[3], body: r[4], btnText: r[5], btnUrl: r[6], sentOn: r[7] };
}

function findEmailRow_(pred) {
  var data = emailsSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var r = rowObj_(data[i]);
    if (pred(r)) return r;
  }
  return null;
}

/** Normalise any free-text user type into one of our audience keys. */
function audKey_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (!s) return 'walker';
  if (s === 'all' || s === 'everyone') return 'all';
  if (/walk|hik|ramb/.test(s)) return 'walker';
  if (/venue|pub|caf|shop|hospitalit/.test(s)) return 'venue';
  if (/brand|sponsor|retail|partner/.test(s)) return 'brand';
  if (/communit|group|volunt|club/.test(s)) return 'community';
  if (/council|tourism|gov|authorit|park/.test(s)) return 'council';
  return 'walker';
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

function heroKeyForDate_(d) {
  var days = Math.round((startOfDay_(LAUNCH).getTime() - startOfDay_(d).getTime()) / 86400000);
  if (days <= 0) return 'live';
  return String(days);   // matches HERO_B64 keys "7","3","1"
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

/* ───────────────────────── HTML email ───────────────────────── */

function htmlEmail_(bodyText, btnText, btnUrl, hasHero) {
  var FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";

  // body text -> paragraphs + bullet lines
  var lines = String(bodyText).split('\n');
  var htmlBody = '', inList = false;
  for (var i = 0; i < lines.length; i++) {
    var ln = lines[i]
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
    var isBullet = /^\s*•\s+/.test(lines[i]);
    if (isBullet) {
      if (!inList) { htmlBody += '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:6px 0 2px">'; inList = true; }
      htmlBody += '<tr><td valign="top" style="padding:3px 10px 3px 0;color:#9CBC32;font-size:16px;line-height:1.5">&bull;</td>' +
                  '<td style="padding:3px 0;color:#475569;font-size:15px;line-height:1.5">' + ln.replace(/^\s*•\s+/, '') + '</td></tr>';
    } else {
      if (inList) { htmlBody += '</table>'; inList = false; }
      if (ln.trim() === '') { htmlBody += '<div style="height:14px;line-height:14px;font-size:0">&nbsp;</div>'; }
      else { htmlBody += '<div>' + ln + '</div>'; }
    }
  }
  if (inList) htmlBody += '</table>';

  var hero = hasHero
    ? '<tr><td style="padding:0"><img src="cid:hero" width="600" alt="Togro" style="display:block;width:100%;height:auto;border:0"></td></tr>'
    : '';

  var button = '';
  if (btnUrl && String(btnUrl).trim()) {
    button =
      '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:24px 0 4px"><tr>' +
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
    // header (navy) with inline logo + wordmark
    '<tr><td style="background:#111827;padding:20px 28px">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
        '<td style="vertical-align:middle">' +
          '<img src="cid:togrologo" width="30" alt="Togro" style="vertical-align:middle;border:0;display:inline-block">' +
          '<span style="font-family:' + FONT + ';color:#ffffff;font-weight:800;font-size:21px;letter-spacing:-.5px;vertical-align:middle;padding-left:9px">togro</span>' +
        '</td>' +
        '<td align="right" style="vertical-align:middle">' +
          '<span style="font-family:' + FONT + ';color:#7F8BA1;font-size:10px;font-weight:700;letter-spacing:2px">COUNTRYSIDE&nbsp;INTELLIGENCE</span>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>' +
    // hero countdown image (inline)
    hero +
    // body
    '<tr><td style="padding:30px 30px 28px;font-family:' + FONT + ';color:#0F172A;font-size:15px;line-height:1.65">' +
      htmlBody + button +
    '</td></tr>' +
    // footer
    '<tr><td style="padding:20px 30px;background:#F8FAFC;border-top:1px solid #EEF2F6;font-family:' + FONT + ';color:#64748B;font-size:12px;line-height:1.6">' +
      '<span style="color:#0F172A;font-weight:700">Togro</span> &middot; Real-time countryside intelligence<br>' +
      'Questions? Call or WhatsApp <a href="tel:+447348955621" style="color:#0B7E74;text-decoration:none">+44&nbsp;7348&nbsp;955621</a> &middot; ' +
      '<a href="' + SITE_URL + '" style="color:#0B7E74;text-decoration:none">togro.co</a><br>' +
      '<span style="color:#94A3B8">You joined the Togro waitlist. ' +
        '<a href="mailto:team@togro.co?subject=Unsubscribe" style="color:#94A3B8;text-decoration:underline">Unsubscribe</a>.</span>' +
    '</td></tr>' +
  '</table>' +
 '</td></tr>' +
'</table>';
}
