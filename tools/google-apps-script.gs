/**
 * Luxury Standards Dinner — seat requests -> Google Sheet + confirmation email.
 *
 * Trimmed copy of the event-RSVP path in
 * eugene92-tt/tt-learning-platform/tools/google-apps-script.gs (Asah), deployed
 * on its OWN sheet so UHNW leads stay apart from Asah/HR leads.
 *
 * Deploy:
 *   1. Create a new Google Sheet (e.g. "Luxury Standards Dinner — Seats").
 *   2. Extensions -> Apps Script, delete the placeholder, paste this file.
 *   3. Set TOKEN below to any random string.
 *   4. Run previewRsvp20 once from the editor to authorise and read the email.
 *   5. Deploy -> New deployment -> Web app.  Execute as: Me.  Who has access: Anyone.
 *   6. Paste the /exec URL into ENDPOINT and the TOKEN into TOKEN at the top of
 *      the <script> in luxury-standards-dinner/index.html.
 *
 * Re-deploy (Manage deployments -> edit -> Version: New version) after any
 * change here, or the live URL keeps serving the old code.
 *
 * Each session has its own tab. The first SEATS distinct emails on a tab are
 * Confirmed and get the "seat confirmed" email with a calendar invite; anyone
 * after that is Waitlisted. The script cannot judge competing brands — check
 * the Brand / Industry columns and un-confirm by hand if two brands clash.
 */

const TOKEN = 'CHANGE-ME';
const SEATS = 4;

const SEND_CONFIRMATION = true;
const HTML_EMAIL = true;
const FROM_NAME = 'Thriving Talents';
const REPLY_TO  = 'sales@thrivingtalents.com';
const PAGE_URL  = 'https://uhnw.thrivingtalents.com/luxury-standards-dinner';

const LINEUP = 'Hosted by Michael Teoh Su Lim (Founder & CEO, Thriving Talents), joined by Reinald Goh, a luxury industry insider with over 15 years across Prada, Dior, Celine, Gucci and Tiffany & Co., and Ethan Lin, an expert in sales psychology and high-net-worth customer engagement.';
const WHERE  = 'Rakuzen (Private Dining Room), NU Sentral, KL Sentral';

const EVENTS = {
  'LSD-2026-11-19': {
    sheet: 'Thu 19 Nov 2026',
    title: 'Luxury Standards Dinner',
    // 7.00pm–8.30pm MYT (UTC+8)
    startUtc: '20261119T110000Z',
    endUtc: '20261119T123000Z',
    date: 'Thursday, 19 November 2026',
    other: 'Friday, 20 November'
  },
  'LSD-2026-11-20': {
    sheet: 'Fri 20 Nov 2026',
    title: 'Luxury Standards Dinner',
    startUtc: '20261120T110000Z',
    endUtc: '20261120T123000Z',
    date: 'Friday, 20 November 2026',
    other: 'Thursday, 19 November'
  }
};
const TIME = '7.00pm to 8.30pm (MYT)';
const GROUND = '#0E0B08', ACCENT = '#C9A35C';

const HEADERS = [
  'Email', 'Name', 'Brand', 'Industry', 'Job title', 'Mobile',
  'First seen', 'Last seen', 'Source', 'Submissions', 'Status', 'Confirmation'
];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const FONT_STACK = "Georgia, 'Times New Roman', serif";

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (TOKEN && TOKEN !== 'CHANGE-ME' && body.token !== TOKEN) return json({ ok: false, error: 'bad token' });

    const email = String(body.email || '').trim().toLowerCase();
    if (!EMAIL_RE.test(email)) return json({ ok: false, error: 'bad email' });

    const ev = Object.prototype.hasOwnProperty.call(EVENTS, body.source) ? EVENTS[body.source] : null;
    if (!ev) return json({ ok: false, error: 'unknown session' });

    return json(handleRsvp(body, email, ev));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Opening the /exec URL in a browser shows the deployment is live and how
// many seats each session has left.
function doGet() {
  const seats = {};
  Object.keys(EVENTS).forEach(function (k) { seats[k] = SEATS - confirmedCount(getSheet(EVENTS[k])); });
  return json({
    ok: true,
    service: 'tt-luxury-standards-dinner',
    seatsLeft: seats,
    mailQuotaRemaining: SEND_CONFIRMATION ? MailApp.getRemainingDailyQuota() : null
  });
}

function clip_(v) { return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').slice(0, 200); }

function handleRsvp(body, email, ev) {
  const f = body.fields || {};
  const name = clip_(body.name);
  const vals = { 'Brand': clip_(f.brand), 'Industry': clip_(f.industry), 'Job title': clip_(f.title), 'Mobile': clip_(f.mobile) };
  let isNew = false, row = 0, status = '';

  // Two people can submit for the last seat at once; the lock decides who got it.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet(ev);
    const c = columnsFor(sheet, HEADERS);
    const now = new Date();
    const found = findEmailRow(sheet, email, c.col['Email']);
    if (found === -1) {
      // ponytail: counts the Status column on every submit; fine for a 4-seat dinner.
      status = confirmedCount(sheet) < SEATS ? 'Confirmed' : 'Waitlist';
      sheet.appendRow(rowFor(sheet, c, Object.assign({
        'Email': email, 'Name': name, 'First seen': now, 'Last seen': now,
        'Source': clip_(body.qrSource), 'Submissions': 1, 'Status': status
      }, vals)));
      row = sheet.getLastRow();
      isNew = true;
    } else {
      // Same person again: one row, newest details, status unchanged.
      row = found;
      status = String(sheet.getRange(row, c.col['Status']).getValue());
      const set = function (h, v) { if (v) sheet.getRange(row, c.col[h]).setValue(v); };
      set('Name', name);
      Object.keys(vals).forEach(function (h) { set(h, vals[h]); });
      sheet.getRange(row, c.col['Last seen']).setValue(now);
      bump(sheet, row, c.col['Submissions']);
    }
  } finally {
    lock.releaseLock();
  }

  // One email per person per session, however many times the form is sent.
  if (isNew) sendRsvpConfirmation(ev, email, name, row, status);
  return { ok: true, status: status };
}

function confirmedCount(sheet) {
  const c = columnsFor(sheet, HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  return sheet.getRange(2, c.col['Status'], last - 1, 1).getValues()
    .filter(function (r) { return String(r[0]).trim() === 'Confirmed'; }).length;
}

function getSheet(ev) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ev.sheet);
  if (!sheet) {
    sheet = ss.insertSheet(ev.sheet);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/* ---------- Email ---------- */

function subjectFor(ev, status) {
  return status === 'Waitlist'
    ? 'You are on the waitlist: Luxury Standards Dinner, ' + ev.date.replace(/^\w+, /, '')
    : 'Your seat is confirmed: Luxury Standards Dinner, ' + ev.date.replace(/^\w+, /, '');
}

function rsvpMail(ev, to, name, status) {
  const opts = {
    to: to,
    subject: subjectFor(ev, status),
    body: rsvpText(ev, name, status),
    name: FROM_NAME,
    replyTo: REPLY_TO
  };
  if (status !== 'Waitlist') {
    opts.attachments = [Utilities.newBlob(eventIcs(ev), 'text/calendar', icsName(ev))];
  }
  if (HTML_EMAIL) opts.htmlBody = rsvpHtml(ev, name, status);
  return opts;
}

function sendRsvpConfirmation(ev, email, name, row, status) {
  const sheet = getSheet(ev);
  const confirmCol = columnsFor(sheet, HEADERS).col['Confirmation'];
  const stamp = function (v) { sheet.getRange(row, confirmCol).setValue(v); };
  if (!SEND_CONFIRMATION) return stamp('off');
  try {
    if (MailApp.getRemainingDailyQuota() < 1) return stamp('quota exhausted');
    MailApp.sendEmail(rsvpMail(ev, email, name, status));
    stamp(new Date());
  } catch (err) {
    stamp('failed: ' + String(err).slice(0, 200));
  }
}

/** Run by hand after a quota day: retries every row whose Confirmation is not a date. Safe to run twice. */
function resendMissingRsvpConfirmations() {
  let sent = 0;
  Object.keys(EVENTS).forEach(function (key) {
    const ev = EVENTS[key];
    const sheet = getSheet(ev);
    const c = columnsFor(sheet, HEADERS);
    const last = sheet.getLastRow();
    if (last < 2) return;
    const rows = sheet.getRange(2, 1, last - 1, c.width).getValues();
    const at = function (r, h) { return r[c.col[h] - 1]; };
    for (let i = 0; i < rows.length; i++) {
      if (at(rows[i], 'Confirmation') instanceof Date) continue;
      const to = String(at(rows[i], 'Email')).trim();
      if (!EMAIL_RE.test(to)) continue;
      if (MailApp.getRemainingDailyQuota() < 1) break;
      sendRsvpConfirmation(ev, to, String(at(rows[i], 'Name')).trim(), i + 2, String(at(rows[i], 'Status')));
      sent++;
    }
  });
  Logger.log('resendMissingRsvpConfirmations: sent ' + sent);
}

/** Pick one from the Run menu to read the email in your own inbox. Does not touch the sheet. */
function previewRsvp20() { previewRsvp('LSD-2026-11-20', 'Confirmed'); }
function previewWaitlist20() { previewRsvp('LSD-2026-11-20', 'Waitlist'); }
function previewRsvp19() { previewRsvp('LSD-2026-11-19', 'Confirmed'); }
function previewRsvp(key, status) {
  const opts = rsvpMail(EVENTS[key], Session.getEffectiveUser().getEmail(), 'Eugene', status);
  opts.subject = '[preview] ' + opts.subject;
  MailApp.sendEmail(opts);
}

function icsName(ev) { return 'luxury-standards-dinner-' + ev.startUtc.slice(0, 8) + '.ics'; }

function googleCalendarLink(ev) {
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent(ev.title) +
    '&dates=' + ev.startUtc + '/' + ev.endUtc +
    '&location=' + encodeURIComponent(WHERE) +
    '&details=' + encodeURIComponent(PAGE_URL);
}

// RFC 5545: CRLF line ends, text values escaped, stable UID per session.
function eventIcs(ev) {
  const esc = function (t) { return String(t).replace(/\\/g, '\\\\').replace(/([,;])/g, '\\$1').replace(/\n/g, '\\n'); };
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Thriving Talents//Event RSVP//EN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:' + icsName(ev).replace('.ics', '') + '@thrivingtalents.com',
    'DTSTAMP:' + stamp,
    'DTSTART:' + ev.startUtc,
    'DTEND:' + ev.endUtc,
    'SUMMARY:' + esc(ev.title),
    'LOCATION:' + esc(WHERE),
    'DESCRIPTION:' + esc(LINEUP + '\n' + PAGE_URL),
    'URL:' + PAGE_URL,
    'END:VEVENT',
    'END:VCALENDAR'
  ].map(fold).join('\r\n') + '\r\n';
}

function fold(line) {
  const out = [];
  while (line.length > 75) { out.push(line.slice(0, 75)); line = ' ' + line.slice(75); }
  out.push(line);
  return out.join('\r\n');
}

// One paragraph per line, no hard wraps: mail clients reflow plain text.
function rsvpText(ev, name, status) {
  const hi = name ? 'Hi ' + name + ',' : 'Hi,';
  if (status === 'Waitlist') {
    return [
      hi,
      'Thank you for your interest in the Luxury Standards Dinner. All ' + SEATS + ' seats for ' + ev.date + ' have now been taken, so we have placed you on the waitlist.',
      'If a seat opens up we will write to you straight away. You are also welcome to request a seat on ' + ev.other + ' instead: ' + PAGE_URL,
      'Thriving Talents\n' + REPLY_TO
    ].join('\n\n');
  }
  return [
    hi,
    'Thank you. Your seat at the Luxury Standards Dinner is confirmed.',
    'When: ' + ev.date + ', ' + TIME + '\nWhere: ' + WHERE,
    LINEUP,
    'A calendar invite is attached. To add it to Google Calendar instead: ' + googleCalendarLink(ev),
    'If your plans change, just reply to this email and let us know, so we can offer the seat to the next guest.',
    'See you at the table,\nThriving Talents\n' + REPLY_TO
  ].join('\n\n');
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Tables and inline styles only: Outlook renders with Word, Gmail strips <style>.
function rsvpHtml(ev, name, status) {
  const ink = '#1A1410', muted = '#6B6158';
  const wait = status === 'Waitlist';
  const p = function (t) { return '<p style="margin:0 0 16px;">' + t + '</p>'; };
  const fact = function (k, v) {
    return '<tr><td style="padding:6px 16px 6px 0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:.16em;' +
      'text-transform:uppercase;color:' + muted + ';white-space:nowrap;vertical-align:top;">' + k + '</td>' +
      '<td style="padding:6px 0;font-size:16px;color:' + ink + ';">' + v + '</td></tr>';
  };
  const day = escapeHtml(ev.date.replace(/^\w+, /, ''));
  const body = wait ? [
    p('Thank you for your interest in the Luxury Standards Dinner. All ' + SEATS + ' seats for ' + escapeHtml(ev.date) + ' have now been taken, so we have placed you on the <strong>waitlist</strong>.'),
    p('If a seat opens up we will write to you straight away. You are also welcome to <a href="' + PAGE_URL + '" style="color:' + ink + ';text-decoration:underline;">request a seat on ' + escapeHtml(ev.other) + '</a> instead.')
  ] : [
    p('Thank you. Your seat at the Luxury Standards Dinner is <strong>confirmed</strong>.'),
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:4px 0 20px;border-collapse:collapse;font-family:' + FONT_STACK + ';">',
    fact('When', escapeHtml(ev.date) + '<br>' + escapeHtml(TIME)),
    fact('Where', escapeHtml(WHERE)),
    '</table>',
    p(escapeHtml(LINEUP)),
    p('A calendar invite is attached. Using Google Calendar? <a href="' + googleCalendarLink(ev) + '" style="color:' + ink + ';text-decoration:underline;">Add it there instead</a>.'),
    p('If your plans change, just reply to this email and let us know, so we can offer the seat to the next guest.')
  ];
  return [
    '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light only">',
    '<title>' + escapeHtml(subjectFor(ev, status)) + '</title></head>',
    '<body style="margin:0;padding:0;background:#F2EDE4;">',
    '<div style="display:none;max-height:0;max-width:0;opacity:0;overflow:hidden;font-size:1px;line-height:1px;color:#F2EDE4;">',
    wait ? 'You are on the waitlist for ' + day + '.' : 'Your seat is confirmed. ' + day + ', 7pm at Rakuzen, NU Sentral.',
    '</div>',
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#F2EDE4" style="background:#F2EDE4;border-collapse:collapse;">',
    '<tr><td align="center" style="padding:24px 12px;">',
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;background:#FFFFFF;border-collapse:collapse;">',
    '<tr><td bgcolor="' + GROUND + '" style="background:' + GROUND + ';padding:30px 32px;font-family:' + FONT_STACK + ';">',
    '<div style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:.24em;text-transform:uppercase;color:' + ACCENT + ';">Special Dinner Invitation &middot; ' + day + '</div>',
    '<div style="margin-top:10px;font-size:28px;line-height:1.2;color:#FFFFFF;">Luxury Standards Dinner</div>',
    '<div style="margin-top:6px;font-size:15px;font-style:italic;color:#D9CBB0;">An Intimate Evening of Conversations, Connections &amp; Luxury Insights</div>',
    '</td></tr>',
    '<tr><td bgcolor="' + ACCENT + '" style="background:' + ACCENT + ';font-size:0;line-height:0;height:2px;">&#8203;</td></tr>',
    '<tr><td style="padding:32px;font-family:' + FONT_STACK + ';font-size:16px;line-height:1.6;color:' + ink + ';">',
    p(name ? 'Hi ' + escapeHtml(name) + ',' : 'Hi,'),
    body.join(''),
    '<p style="margin:28px 0 0;">' + (wait ? '' : 'See you at the table,<br>') + '<strong>Thriving Talents</strong><br>',
    '<a href="mailto:' + REPLY_TO + '" style="color:' + ink + ';text-decoration:underline;">' + REPLY_TO + '</a></p>',
    '</td></tr>',
    '<tr><td style="padding:18px 32px 26px;border-top:1px solid #F2EDE4;font-family:Arial,sans-serif;font-size:12px;line-height:1.55;color:' + muted + ';">',
    'You are receiving this because you requested a seat at ' + PAGE_URL.replace('https://', '') + '. We use your details only to run this dinner.',
    '</td></tr>',
    '</table></td></tr></table></body></html>'
  ].join('');
}

/* ---------- Sheet helpers (same as Asah) ---------- */

function findEmailRow(sheet, email, col) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const values = sheet.getRange(2, col || 1, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return -1;
}

/** Finds columns by header text, so the team can add or reorder columns by hand. */
function columnsFor(sheet, required) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const row1 = sheet.getRange(1, 1, 1, width).getValues()[0];
  const col = {};
  row1.forEach(function (h, i) {
    const key = String(h).trim();
    if (key && !(key in col)) col[key] = i + 1;
  });
  let end = width;
  required.forEach(function (h) {
    if (!col[h]) {
      end += 1;
      sheet.getRange(1, end).setValue(h).setFontWeight('bold');
      col[h] = end;
    }
  });
  return { col: col, width: end };
}

function rowFor(sheet, c, values) {
  const row = [];
  for (let i = 0; i < Math.max(c.width, sheet.getLastColumn()); i++) row.push('');
  Object.keys(values).forEach(function (h) { row[c.col[h] - 1] = values[h]; });
  return row;
}

function bump(sheet, row, column) {
  const cell = sheet.getRange(row, column);
  cell.setValue((Number(cell.getValue()) || 1) + 1);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
