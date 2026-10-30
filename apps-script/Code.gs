/**
 * Serve sign-ups: Google Sheet backend for the lobby TV and the sign-up page.
 *
 * Setup (once):
 *   1. Make a new Google Sheet → Extensions → Apps Script. Paste this whole file over Code.gs and save.
 *   2. Pick `setup` in the function menu and click Run. Approve the permissions it asks for.
 *      It creates a Teams tab (filled with the starting teams and positions) and a Signups tab.
 *   3. Optional: Project Settings → Script properties → add NOTIFY_EMAIL = the address(es) that should hear
 *      about every sign-up. The first one is also the reply-to on volunteers' confirmation emails.
 *   4. Deploy → New deployment → type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 *   5. Copy the Web app URL into CONFIG.apiUrl in index.html.
 *
 * After editing this file later: Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
 * The URL stays the same.
 *
 * Teams tab: one row per position. team, position, needed (how many blank tags to hang at first; there's no cap,
 *   and every position always keeps one blank tag), lead_email (optional: that person also gets an email for each
 *   sign-up). Rows show in sheet order. Set needed to 0 to hide a position.
 * Signups tab: every sign-up. Set status to "removed" (or delete the row) to free the spot.
 *
 * Privacy: the page only ever receives team, position, needed and each volunteer's first name + last initial.
 * Emails, phone numbers and notes never leave the sheet.
 */

const TEAMS = 'Teams';
const SIGNUPS = 'Signups';
const TEAM_COLS = ['team', 'position', 'needed', 'lead_email'];
const SIGNUP_COLS = ['submitted_at', 'team', 'position', 'name', 'email', 'phone', 'notes', 'status'];
const GONE = ['removed', 'cancelled', 'canceled', 'hidden', 'no'];

const CHURCH = 'Oceanside Church';

// Starting teams. Each spot shows as its own card, so keep positions to one short word.
// Edit them in the sheet afterwards; this only fills an empty Teams tab.
const START = [
  ['Logistics', 'Parking', 3], ['Logistics', 'Coffee', 2], ['Logistics', 'Setup', 2],
  ['Worship', 'Vocals', 2], ['Worship', 'Acoustic', 1], ['Worship', 'Electric', 1],
  ['Worship', 'Drums', 1], ['Worship', 'Bass', 1], ['Worship', 'Keys', 1],
  ['Production', 'Camera', 2], ['Production', 'Sound', 2], ['Production', 'Visuals', 2], ['Production', 'Setup', 1],
  ['Chairs', 'Setup', 4], ['Chairs', 'Teardown', 3],
  ['ONE80 Kids', 'Nursery', 2], ['ONE80 Kids', 'Preschool', 2], ['ONE80 Kids', 'Elementary', 2], ['ONE80 Kids', 'Check-in', 1],
  ['Welcome', 'Greeter', 3], ['Welcome', 'Info', 2], ['Welcome', 'Usher', 2],
];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const teams = ss.getSheetByName(TEAMS) || ss.getSheets()[0].setName(TEAMS);
  ensureHeaders_(teams, TEAM_COLS);
  if (teams.getLastRow() < 2) {
    const head = table_(teams).head;
    const rows = START.map(([team, position, needed]) => head.map(h => ({ team, position, needed }[h] ?? '')));
    teams.getRange(2, 1, rows.length, head.length).setValues(rows);
  }
  const signups = ss.getSheetByName(SIGNUPS) || ss.insertSheet(SIGNUPS);
  ensureHeaders_(signups, SIGNUP_COLS);
  const sc = table_(signups).head.indexOf('status') + 1;
  signups.getRange(2, sc, signups.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['confirmed', 'removed'], true).setAllowInvalid(true).build());
  CacheService.getScriptCache().remove('teams');
}

/**
 * One-off, for testing: fills about half the spots with sample names so every state shows on the TV.
 * Run `clearSignups` before launch to empty the Signups tab again (headers and dropdowns stay).
 */
function seedDemo() {
  const names = ['Sarah Klassen', 'Mike Dhillon', 'Jen Okafor', 'Tom Reyes', 'Priya Nair', 'Dave Thiessen', 'Amy Chen',
    'Luke Barrett', 'Grace Mwangi', 'Ben Fraser', 'Hannah Lee', 'Josh Martin', 'Kate Sullivan', 'Noah Singh',
    'Olivia Brooks', 'Sam Wiebe', 'Rachel Kim', 'Chris Dueck', 'Emma Wright', 'Jacob Pham'];
  clearSignups();
  const sh = sheet_(SIGNUPS), head = table_(sh).head, now = new Date();
  const rows = [];
  let k = 0;
  positions_().forEach((p, i) => {
    const take = i % 4 === 0 ? p.needed : Math.floor(p.needed / 2);   // a few full positions, most half full
    for (let j = 0; j < take && k < names.length; j++) {
      const rec = { submitted_at: now, team: p.team, position: p.position, name: names[k++], status: 'confirmed', notes: 'Sample' };
      rows.push(head.map(h => rec[h] ?? ''));
    }
  });
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  CacheService.getScriptCache().remove('teams');
}

function clearSignups() {
  const sh = sheet_(SIGNUPS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  CacheService.getScriptCache().remove('teams');
}

/* ── Read: teams, positions and who has signed up ──
   Cached for 30 s so a busy page doesn't re-read the sheet for every visitor. Any edit to the sheet
   (by hand or by a sign-up) clears the cache, so changes show on the next refresh. */
function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('teams');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const out = JSON.stringify(publicTeams_());
  cache.put('teams', out, 30);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function publicTeams_() {
  const filled = filledBy_();
  const teams = [];
  positions_().forEach(p => {
    let t = teams.find(x => x.team === p.team);
    if (!t) teams.push(t = { team: p.team, positions: [] });
    t.positions.push({ position: p.position, needed: p.needed, filled: (filled[key_(p.team, p.position)] || []).map(s => short_(s.name)) });
  });
  return { teams };
}

// Simple trigger: runs on every hand edit in the sheet (no setup needed) and drops the cached data.
function onEdit() { CacheService.getScriptCache().remove('teams'); }

/* ── Write: a sign-up ── */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.website) return json_({ ok: true }); // honeypot field: bots fill it, people never see it
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const res = d.action === 'signup' ? signup_(d) : { ok: false, error: 'unknown_action' };
    if (res.ok) CacheService.getScriptCache().remove('teams');   // the new name shows on the TV straight away
    return json_(res);
  } finally {
    lock.releaseLock();
  }
}

function signup_(d) {
  const team = clean_(d.team, 60), position = clean_(d.position, 60);
  const name = clean_(d.name, 80), email = clean_(d.email, 120), phone = clean_(d.phone, 40), notes = clean_(d.notes, 400);
  if (!team || !position || !name || !(isEmail_(email) || phone)) return { ok: false, error: 'missing' };
  const p = positions_().find(x => x.team === team && x.position === position);
  if (!p) return { ok: false, error: 'not_found' };
  const taken = filledBy_()[key_(team, position)] || [];
  if (email && taken.some(s => s.email.toLowerCase() === email.toLowerCase())) return { ok: false, error: 'already' };

  const sh = sheet_(SIGNUPS);
  const rec = { submitted_at: new Date(), team, position, name, email, phone, notes, status: 'confirmed' };
  sh.appendRow(table_(sh).head.map(h => safe_(rec[h])));

  const spot = team + ' · ' + position;
  if (isEmail_(email)) confirm_(email, name, team, position);
  notify_([p.lead], 'New volunteer · ' + spot + ' · ' + name, name + ' joined ' + spot, [
    ['The spot', [['Team', team], ['Position', position], ['Signed up', String(taken.length + 1)]]],
    ['Volunteer', [['Name', name], ['Email', email], ['Phone', phone], ['Notes', notes]]],
    ['To do', null, [
      'Get in touch with ' + name.split(/\s+/)[0] + ' this week and welcome them to the team.',
      'Not a fit? On the Signups tab, set their status to "removed" to free the spot.',
    ]],
  ]);
  return { ok: true, receipt: isEmail_(email) };
}

/* ── Helpers ── */
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function table_(sh) {
  const v = sh.getDataRange().getValues();
  return { head: v.shift().map(h => String(h).trim().toLowerCase()), rows: v };
}

// Every position on the Teams tab that needs at least one person, in sheet order.
function positions_() {
  const { head, rows } = table_(sheet_(TEAMS));
  const get = (r, k) => (head.indexOf(k) < 0 ? '' : String(r[head.indexOf(k)]).trim());
  return rows.map(r => ({ team: get(r, 'team'), position: get(r, 'position'), needed: parseInt(get(r, 'needed'), 10) || 0, lead: get(r, 'lead_email') }))
    .filter(p => p.team && p.position && p.needed > 0);
}

// Current sign-ups grouped by team + position, oldest first.
function filledBy_() {
  const { head, rows } = table_(sheet_(SIGNUPS));
  const get = (r, k) => (head.indexOf(k) < 0 ? '' : String(r[head.indexOf(k)]).trim());
  const out = {};
  rows.forEach(r => {
    const name = get(r, 'name');
    if (!name || GONE.includes(get(r, 'status').toLowerCase())) return;
    const k = key_(get(r, 'team'), get(r, 'position'));
    if (!out[k]) out[k] = [];
    out[k].push({ name, email: get(r, 'email') });
  });
  return out;
}

function key_(team, position) { return team.toLowerCase() + '\u0000' + position.toLowerCase(); }

// "Sarah Klassen" → "Sarah K."
function short_(name) {
  const w = String(name).trim().split(/\s+/);
  return w.length > 1 ? w[0] + ' ' + w[w.length - 1][0].toUpperCase() + '.' : w[0];
}

function ensureHeaders_(sh, cols) {
  const width = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim().toLowerCase());
  cols.forEach(c => {
    if (head.includes(c)) return;
    const at = head[0] === '' && head.length === 1 ? 1 : head.length + 1;
    sh.getRange(1, at).setValue(c);
    if (at === 1) head[0] = c; else head.push(c);
  });
  sh.setFrozenRows(1);
}

function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max); }

// Stop text that starts with = + - @ from being treated as a spreadsheet formula.
function safe_(v) {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function isEmail_(a) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a || ''); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// NOTIFY_EMAIL can hold several addresses, separated by commas, semicolons or spaces.
function admins_() {
  const raw = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || '';
  return raw.split(/[\s,;]+/).map(a => a.trim()).filter(isEmail_);
}

// Email trouble never blocks a sign-up: the row is already saved by the time this runs.
function send_(opts) {
  try { MailApp.sendEmail(Object.assign({ name: CHURCH }, opts)); }
  catch (err) { console.error('Email to ' + opts.to + ' failed: ' + err); }
}

/* Email layout: cream card, Ink Blue headings, one table per section. Plain text goes along for mail apps that want it. */
const INK = '#201F1F', INK_BLUE = '#1A376B', BLUE = '#3C7BEA', CREAM = '#F3EFE6';
const FONT = 'Arial,Helvetica,sans-serif';
function frame_(inner) {
  return '<div style="background:' + CREAM + ';padding:28px 12px;font-family:' + FONT + ';color:' + INK + '">' +
    '<div style="max-width:540px;margin:0 auto;background:#FFFFFF;border-radius:16px;padding:26px 26px 22px">' +
    '<div style="font:700 11px/1 ' + FONT + ';letter-spacing:.16em;text-transform:uppercase;color:' + BLUE + ';margin-bottom:14px">Volunteer teams</div>' +
    inner + '</div></div>';
}
function heading_(t) {
  return '<div style="margin:22px 0 8px;padding-bottom:6px;border-bottom:2px solid ' + CREAM + ';font:700 12px/1 ' + FONT + ';letter-spacing:.12em;text-transform:uppercase;color:' + INK_BLUE + '">' + esc_(t) + '</div>';
}
function rowsHtml_(rows) {
  return '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:15px;line-height:1.4">' +
    rows.filter(r => r[1]).map(r => '<tr><td style="padding:5px 12px 5px 0;width:96px;vertical-align:top;color:' + INK_BLUE + ';font:700 12px/1.6 ' + FONT + ';text-transform:uppercase;letter-spacing:.06em">' +
      esc_(r[0]) + '</td><td style="padding:5px 0;vertical-align:top">' + esc_(r[1]) + '</td></tr>').join('') + '</table>';
}
function stepsHtml_(steps) {
  return '<ol style="margin:4px 0 0;padding-left:20px;font-size:15px;line-height:1.5">' + steps.map(s => '<li style="margin:0 0 6px">' + esc_(s) + '</li>').join('') + '</ol>';
}
function sectionsText_(sections) {
  return sections.map(([title, rows, steps]) => title.toUpperCase() + '\n' +
    (rows ? rows.filter(r => r[1]).map(r => r[0] + ': ' + r[1]).join('\n') : steps.map((s, k) => (k + 1) + '. ' + s).join('\n'))).join('\n\n');
}

// Staff email to NOTIFY_EMAIL plus any extra addresses (the team lead): headline, sections, button to the sheet.
function notify_(extra, subject, headline, sections) {
  const to = [...new Set(admins_().concat(extra.filter(isEmail_)))];
  if (!to.length) return;
  const url = SpreadsheetApp.getActive().getUrl();
  const html = frame_('<h1 style="margin:0;font:800 24px/1.15 ' + FONT + ';color:' + INK_BLUE + '">' + esc_(headline) + '</h1>' + sections.map(([title, rows, steps]) => heading_(title) + (rows ? rowsHtml_(rows) : stepsHtml_(steps))).join('') +
    '<p style="margin:22px 0 0"><a href="' + url + '" style="display:inline-block;padding:12px 20px;border-radius:999px;background:' + INK_BLUE + ';color:#FFFFFF;text-decoration:none;font:700 13px/1 ' + FONT + ';letter-spacing:.06em;text-transform:uppercase">Open the sheet</a></p>');
  send_({ to: to.join(','), subject, htmlBody: html, body: headline + '\n\n' + sectionsText_(sections) + '\n\nThe sheet: ' + url });
}

// The volunteer's confirmation: thanks, where they signed up, and what happens next. Replies go to the staff.
function confirm_(to, name, team, position) {
  const first = name.split(/\s+/)[0];
  const rows = [['Team', team], ['Position', position]];
  const steps = [
    'Someone from the ' + team + ' team will get in touch this week.',
    'They’ll help you get ready for your first Sunday.',
  ];
  const html = frame_(
    '<h1 style="margin:0 0 10px;font:800 26px/1.15 ' + FONT + ';color:' + INK_BLUE + '">You’re on the team. Thank you.</h1>' +
    '<p style="margin:0;font-size:16px;line-height:1.5">Hi ' + esc_(first) + ', thanks for signing up to serve.</p>' +
    heading_('Your spot') + rowsHtml_(rows) + heading_('What happens next') + stepsHtml_(steps) +
    '<p style="margin:20px 0 0;font-size:15px;line-height:1.5">Have a question, or need to change your spot? Just reply to this email.</p>' +
    '<p style="margin:14px 0 0;font-size:15px">' + esc_(CHURCH) + '</p>');
  const text = 'You’re on the team. Thank you.\n\nHi ' + first + ', thanks for signing up to serve.\n\n' +
    sectionsText_([['Your spot', rows], ['What happens next', null, steps]]) +
    '\n\nHave a question, or need to change your spot? Just reply to this email.\n\n' + CHURCH;
  const replyTo = admins_()[0];
  send_(Object.assign({ to, subject: 'You’re on the ' + team + ' team', htmlBody: html, body: text }, replyTo ? { replyTo } : {}));
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
