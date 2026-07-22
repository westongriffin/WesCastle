/**
 * Wes's Castle — royal ledger (backend)
 * Google Apps Script web app storing rooms & reservations in a Google Sheet.
 *
 * DEPLOY:
 *  1. script.google.com → New project → paste this file.
 *  2. Deploy → New deployment → Web app.
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  3. Copy the /exec URL and paste it into index.html (API_URL const),
 *     or into Keeper's Entrance → Royal ledger field.
 *
 * The spreadsheet ("Wes's Castle Ledger") is created automatically in your
 * Drive on the first request. Default admin PIN: 1234 — change it in the app
 * (Keeper's Entrance → Royal decrees) right away.
 */

var DEFAULT_ADMIN_PIN = '1234';

var SEED_ROOMS = [
  ['jack',    'Jack',              '🪣', 'Went up the hill and never left. Twin bed, zero candlesticks.', '2 compliments / night', '', 1],
  ['jill',    'Jill',              '⛰️', 'Came tumbling after. Objectively the better view — don’t tell Jack.', '2 compliments / night', '', 1],
  ['media',   'The Media Room',    '🎬', 'Fall asleep mid-movie in surround-sound glory. The recliner absolutely counts as a bed.', '1 movie pick / night', '', 4],
  ['east',    'The East Wing',     '🌅', 'Where the sun rises first and the coffee arrives last. Early risers only.', '3 compliments / night', '', 1],
  ['west',    'The West Wing',     '🦅', 'Dramatic hallway walk-and-talks included at no extra charge.', '3 compliments / night', '', 1],
  ['dungeon', "Dawson’s Dungeon",  '⛓️', 'No windows. No rules. No refunds. Our most exclusive suite.', '1 unspecified favor / night', 'dark', 1],
];

var ROOM_HEAD    = ['id','name','emoji','tag','rate','flags','closed','ownerName','ownerPin','ownerNotify','capacity'];
var BOOKING_HEAD = ['id','roomId','type','guestName','start','end','code','createdBy','createdAt','guestEmail','lockCodeId','doorCode'];

/**
 * RUN ME ONCE after pasting a new version: pick "authorizeCastle" in the
 * editor toolbar dropdown and press ▶ Run. Google will show a permission
 * dialog — approve it. This grants the script every permission it needs
 * (including sending email, which notifications require).
 */
function authorizeCastle() {
  var quota = MailApp.getRemainingDailyQuota(); // touches the mail permission
  var ss = ss_();                               // touches Drive/Sheets
  var net = UrlFetchApp.fetch('https://connect.getseam.com/health', {muteHttpExceptions: true}).getResponseCode(); // touches external requests
  Logger.log('✅ Authorized. Ledger: "' + ss.getName() + '". Email quota left today: ' + quota + '. External reach (Seam): HTTP ' + net + '.');
}

/**
 * ONE-TIME REPAIR: if the Rooms sheet was seeded with garbled emoji/text
 * (bad clipboard encoding), pick this function in the editor toolbar
 * dropdown and click ▶ Run once. Rewrites name/emoji/tag/rate from
 * SEED_ROOMS; leaves closed/owner columns untouched.
 */
function repairSeeds() {
  var ss = ss_();
  var sh = ss.getSheetByName('Rooms');
  var rows = readAll_(sh);
  var capCol = ROOM_HEAD.indexOf('capacity') + 1;
  SEED_ROOMS.forEach(function (seed) {
    var r = rows.filter(function (x) { return x.id === seed[0]; })[0];
    if (!r) return;
    sh.getRange(r._row, 2, 1, 4).setValues([[seed[1], seed[2], seed[3], seed[4]]]);
    if (!Number(r.capacity)) sh.getRange(r._row, capCol).setValue(seed[6] || 1);
  });
  Logger.log('Rooms sheet repaired.');
}

/* ---------------- entry points ---------------- */

function doGet() {
  return json_(handle_({action: 'state'}));
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ok: false, error: 'Bad request.'}); }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return json_(handle_(req)); }
  catch (err) { return json_({ok: false, error: 'Ledger error: ' + err.message}); }
  finally { lock.releaseLock(); }
}

/* ---------------- storage ---------------- */

function ss_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; }
  }
  if (!id) {
    ss = SpreadsheetApp.create("Wes's Castle Ledger");
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheet_(ss, 'Rooms', ROOM_HEAD, SEED_ROOMS.map(function (r) {
    return [r[0], r[1], r[2], r[3], r[4], r[5], 'FALSE', '', '', '', r[6] || 1];
  }));
  ensureSheet_(ss, 'Bookings', BOOKING_HEAD, []);
  ensureSheet_(ss, 'Settings', ['key', 'value'], [['adminPin', DEFAULT_ADMIN_PIN]]);
  ensureSheet_(ss, 'Log', ['when', 'to', 'subject', 'result'], []);
  ensureSheet_(ss, 'Sessions', ['token', 'email', 'code', 'expires', 'created'], []);
  ensureSheet_(ss, 'Guests', ['email', 'passHash', 'salt', 'name', 'created'], []);
  ensureCols_(ss.getSheetByName('Rooms'), ROOM_HEAD);      // upgrades older ledgers in place
  ensureCols_(ss.getSheetByName('Bookings'), BOOKING_HEAD);
  return ss;
}

function ensureCols_(sh, head) {
  var row1 = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  head.forEach(function (h) {
    if (row1.indexOf(h) < 0) { sh.getRange(1, row1.length + 1).setValue(h); row1.push(h); }
  });
}

function ensureSheet_(ss, name, head, seedRows) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    seedRows.forEach(function (r) { sh.appendRow(r); });
  }
  return sh;
}

function readAll_(sh) {
  var vals = sh.getDataRange().getValues();
  var head = vals.shift();
  return vals.map(function (row, i) {
    var o = {_row: i + 2};
    head.forEach(function (h, j) { o[h] = row[j]; });
    return o;
  });
}

function dateStr_(v) {
  // Sheets may hand dates back as Date objects; normalize to YYYY-MM-DD
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function db_() {
  var ss = ss_();
  var rooms = readAll_(ss.getSheetByName('Rooms')).map(function (r) {
    r.closed = String(r.closed).toUpperCase() === 'TRUE';
    r.ownerPin = String(r.ownerPin || '');
    r.ownerName = String(r.ownerName || '');
    return r;
  });
  var bookings = readAll_(ss.getSheetByName('Bookings')).map(function (b) {
    b.start = dateStr_(b.start); b.end = dateStr_(b.end);
    b.id = String(b.id); b.code = String(b.code || '');
    return b;
  });
  var settings = {};
  readAll_(ss.getSheetByName('Settings')).forEach(function (s) { settings[s.key] = String(s.value); });
  return {ss: ss, rooms: rooms, bookings: bookings, settings: settings};
}

/* ---------------- helpers ---------------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function fmtD_(s) { // '2026-07-17' → 'Jul 17' (no Date parsing — avoids TZ shifts)
  var M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var p = String(s).split('-');
  return M[(+p[1]) - 1] + ' ' + (+p[2]);
}

function nights_(a, b) { return Math.round((new Date(b) - new Date(a)) / 86400000); }

function splitAddrs_(s) {
  return String(s || '').split(/[,;\s]+/).filter(function (x) { return x.indexOf('@') > 0; });
}

function hashPass_(salt, pass) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + String(pass), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { var v = (b + 256) % 256; return (v < 16 ? '0' : '') + v.toString(16); }).join('');
}

function makeToken_(d, email) { // signed-in device stays valid 180 days
  var tok = Utilities.getUuid();
  d.ss.getSheetByName('Sessions').appendRow([tok, email, '', new Date().getTime() + 180 * 24 * 3600 * 1000, new Date().getTime()]);
  return tok;
}

function setSetting_(d, key, val) {
  var sh = d.ss.getSheetByName('Settings');
  var row = readAll_(sh).filter(function (s) { return s.key === key; })[0];
  if (row) sh.getRange(row._row, 2).setValue(val);
  else sh.appendRow([key, val]);
}

/* ---------------- Yale door codes via Seam ---------------- */

var SEAM_API = 'https://connect.getseam.com';
// Door codes cover the WHOLE stay: midnight at the start of check-in day
// through midnight at the END of checkout day (hour 24 rolls to next 00:00).
var CHECKIN_HOUR = 0, CHECKOUT_HOUR = 24;

function seamCfg_(d) {
  var key = String(d.settings.seamKey || ''), dev = String(d.settings.seamDeviceId || '');
  return key && dev ? {key: key, dev: dev} : null;
}

function seamFetch_(d, path, payload, keyOverride) {
  var key = keyOverride || String(d.settings.seamKey || '');
  var res = UrlFetchApp.fetch(SEAM_API + path, {
    method: 'post', contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + key},
    payload: JSON.stringify(payload || {}), muteHttpExceptions: true
  });
  var body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) {}
  body._status = res.getResponseCode();
  return body;
}

function isoAt_(dateStr, hour) {
  var p = dateStr.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], hour, 0, 0).toISOString();
}

/** Programs a stay-window door code on the lock. Never throws — a lock hiccup
 *  must not break a booking. Returns the pin, or null. */
function createLockCode_(d, b) {
  var cfg = seamCfg_(d); if (!cfg) return null;
  try {
    var pin = String(Math.floor(100000 + Math.random() * 900000));
    var res = seamFetch_(d, '/access_codes/create', {
      device_id: cfg.dev,
      name: ('WC ' + b.code + ' ' + b.guestName).slice(0, 60),
      code: pin,
      starts_at: isoAt_(b.start, CHECKIN_HOUR),
      ends_at: isoAt_(b.end, CHECKOUT_HOUR)
    });
    if (res.access_code) {
      var sh = d.ss.getSheetByName('Bookings');
      sh.getRange(b._row, BOOKING_HEAD.indexOf('lockCodeId') + 1).setValue(res.access_code.access_code_id);
      sh.getRange(b._row, BOOKING_HEAD.indexOf('doorCode') + 1).setValue(pin);
      logNotify_(d, 'lock', 'door code for ' + b.code, 'created — active all day ' + b.start + ' through ' + b.end);
      return pin;
    }
    logNotify_(d, 'lock', 'door code for ' + b.code, 'ERROR: ' + ((res.error && res.error.message) || ('HTTP ' + res._status)));
  } catch (e) { logNotify_(d, 'lock', 'door code for ' + b.code, 'ERROR: ' + e.message); }
  return null;
}

function deleteLockCode_(d, b) {
  if (!seamCfg_(d) || !String(b.lockCodeId || '')) return;
  try {
    seamFetch_(d, '/access_codes/delete', {access_code_id: String(b.lockCodeId)});
    logNotify_(d, 'lock', 'door code for ' + b.code, 'revoked');
  } catch (e) { logNotify_(d, 'lock', 'door code for ' + b.code, 'ERROR revoking: ' + e.message); }
}

function updateLockCode_(d, b, start, end) {
  if (!seamCfg_(d) || !String(b.lockCodeId || '')) return;
  try {
    seamFetch_(d, '/access_codes/update', {access_code_id: String(b.lockCodeId),
      starts_at: isoAt_(start, CHECKIN_HOUR), ends_at: isoAt_(end, CHECKOUT_HOUR)});
    logNotify_(d, 'lock', 'door code for ' + b.code, 'window moved to ' + start + ' → ' + end);
  } catch (e) { logNotify_(d, 'lock', 'door code for ' + b.code, 'ERROR moving: ' + e.message); }
}

/** Pretty room name straight from SEED_ROOMS (immune to a garbled sheet). */
function roomName_(room) {
  if (!room) return 'a chamber';
  var s = SEED_ROOMS.filter(function (x) { return x[0] === room.id; })[0];
  return (s && s[1]) || room.name || room.id;
}

function logNotify_(d, to, subject, result) {
  try { d.ss.getSheetByName('Log').appendRow([new Date(), to, subject, result]); } catch (e) {}
}

/* ---------------- pretty guest emails (castle stationery) ----------------
   Guests get branded HTML mail; notify_ / testNotify stay PLAIN text on
   purpose — keeper addresses may be SMS gateways that choke on HTML. */
function eWrap_(inner) {
  return '<div style="background:#f6f1e6;padding:26px 12px;margin:0">' +
    '<div style="max-width:520px;margin:0 auto;background:#fffdf7;border:1px solid #e6dcc6;border-radius:16px;overflow:hidden;font-family:Georgia,\'Times New Roman\',serif;color:#2b2338">' +
    '<div style="background:#38206b;padding:20px;text-align:center">' +
    '<div style="font-size:34px;line-height:1.2">&#127984;</div>' +
    '<div style="color:#f3e8c9;font-size:21px;font-weight:bold;letter-spacing:.5px">Wes&#8217;s Castle</div></div>' +
    '<div style="padding:22px 26px 10px">' + inner + '</div>' +
    '<div style="background:#f6f1e6;padding:13px;text-align:center;font-size:11.5px;color:#6b6078;font-style:italic">No gold changes hands &mdash; payment is accepted in compliments, favors, and doing your own dishes.</div>' +
    '</div></div>';
}
function eH_(t) { return '<div style="font-size:20px;font-weight:bold;color:#38206b;margin:0 0 8px">' + t + '</div>'; }
function eP_(t) { return '<p style="font-size:14px;line-height:1.6;margin:10px 0;color:#2b2338">' + t + '</p>'; }
function eRows_(pairs) {
  var tr = pairs.map(function (p) {
    return '<tr><td style="padding:7px 12px;color:#6b6078;font-size:12.5px;white-space:nowrap;vertical-align:top;font-family:Arial,sans-serif">' + p[0] + '</td>' +
           '<td style="padding:7px 12px;font-size:14px;font-weight:bold">' + p[1] + '</td></tr>';
  }).join('');
  return '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#faf6ec;border-radius:10px;width:100%;margin:12px 0">' + tr + '</table>';
}
function eCode_(label, code) {
  return '<div style="border:2px dashed #c9a227;background:#f3e8c9;border-radius:12px;padding:12px;text-align:center;margin:14px 0">' +
    '<div style="font-size:10.5px;letter-spacing:1.5px;color:#6d5410;text-transform:uppercase;font-family:Arial,sans-serif">' + label + '</div>' +
    '<div style="font-family:Menlo,Consolas,monospace;font-size:26px;font-weight:bold;letter-spacing:4px;color:#38206b;margin-top:2px">' + code + '</div></div>';
}
function eBtn_(url, label) {
  return '<div style="text-align:center;margin:16px 0"><a href="' + url + '" ' +
    'style="background:#4c2a85;color:#ffffff;text-decoration:none;font-family:Arial,sans-serif;font-size:14px;font-weight:bold;padding:11px 22px;border-radius:10px;display:inline-block">' + label + '</a></div>';
}
function site_(d) { return String(d.settings.siteUrl || '').replace(/\/+$/, ''); }
function eManageBtn_(d) { var s = site_(d); return s ? eBtn_(s, '&#127984; Manage thy stay') : ''; }
function eLockHelp_(d) {
  var s = site_(d);
  return eP_('<b>Using the door lock:</b> enter your door code, then press <b>&#10003;</b>. Leaving? Just tap the circle Yale button at the top &mdash; no code needed.') +
    (s ? eBtn_(s + '/#lockhelp', '&#127916; Watch: how the door lock works') : '');
}
function hasAccount_(d, email) {
  var em = String(email || '').trim().toLowerCase();
  if (!em) return false;
  return readAll_(d.ss.getSheetByName('Guests')).some(function (g) {
    return String(g.email).toLowerCase() === em;
  });
}
function gcalUrl_(roomNameStr, guestName, code, start, end) {
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent('🏰 Wes’s Castle — ' + roomNameStr) +
    '&dates=' + start.replace(/-/g, '') + '/' + end.replace(/-/g, '') +
    '&details=' + encodeURIComponent('Guest: ' + guestName + '\nConfirmation code: ' + code);
}
function eCalRow_(roomNameStr, guestName, code, start, end) {
  return '<div style="text-align:center;margin:14px 0;font-family:Arial,sans-serif;font-size:13.5px">' +
    '<a href="' + gcalUrl_(roomNameStr, guestName, code, start, end) + '" style="color:#4c2a85;font-weight:bold">&#128198; Add to Google Calendar</a>' +
    '<div style="color:#6b6078;font-size:12px;margin-top:3px">Apple or Outlook? Open the attached calendar invite.</div></div>';
}
function eKeyBtn_(d, email) {
  var s = site_(d);
  if (!s || hasAccount_(d, email)) return '';
  return '<div style="border-top:1px solid #e6dcc6;margin-top:18px;padding-top:12px">' +
    eP_('<b>Want all thy stays in one place?</b> Forge a castle key &mdash; a simple account that gathers every visit, past and future, on any device.') +
    '<div style="text-align:center;margin:10px 0"><a href="' + s + '/#createkey" ' +
    'style="background:#c9a227;color:#3d2e05;text-decoration:none;font-family:Arial,sans-serif;font-size:14px;font-weight:bold;padding:11px 22px;border-radius:10px;display:inline-block">&#128273; Create a castle key</a></div></div>';
}
function icsEsc_(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/[,;]/g, function (m) { return '\\' + m; }).replace(/\n/g, '\\n');
}
function stayIcs_(roomNameStr, guestName, code, start, end) {
  var ymd = function (s) { return s.replace(/-/g, ''); };
  var stamp = Utilities.formatDate(new Date(), 'GMT', "yyyyMMdd'T'HHmmss'Z'");
  var ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Wes’s Castle//EN', 'BEGIN:VEVENT',
    'UID:' + code + '@wes-castle', 'DTSTAMP:' + stamp,
    'DTSTART;VALUE=DATE:' + ymd(start), 'DTEND;VALUE=DATE:' + ymd(end),
    'SUMMARY:' + icsEsc_('🏰 Wes’s Castle — ' + roomNameStr),
    'DESCRIPTION:' + icsEsc_('Guest: ' + guestName + '\nConfirmation code: ' + code),
    'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
  return Utilities.newBlob(ics, 'text/calendar', 'WesCastle-' + code + '.ics');
}

/** Alert the Crown + the room's keeper. Addresses may be email or carrier
 *  SMS-gateway addresses (e.g. 5551234567@vtext.com) — kept short for texts.
 *  Every attempt is receipted in the Log sheet. */
function notify_(d, room, subject, body, room2) {
  var admin = splitAddrs_(d.settings.adminNotify);
  var owner = (room ? splitAddrs_(room.ownerNotify) : [])
    .concat(room2 ? splitAddrs_(room2.ownerNotify) : []);
  var addrs = admin.concat(owner);
  if (!addrs.length) { logNotify_(d, '(no one)', subject, 'skipped — no admin/keeper notify addresses set'); return; }
  if (room && !owner.length) logNotify_(d, '(no keeper)', subject, 'room has no keeper notify address — admin only');
  var seen = {};
  addrs.forEach(function (a) {
    if (seen[a]) return; seen[a] = 1;
    try { MailApp.sendEmail(a, subject, body); logNotify_(d, a, subject, 'sent'); }
    catch (e) { logNotify_(d, a, subject, 'ERROR: ' + e.message); }
  });
}

function mailGuest_(d, addr, subject, body, html, attachments) {
  var a = splitAddrs_(addr)[0];
  if (!a) return;
  try {
    var opts = {};
    if (html) opts.htmlBody = eWrap_(html);
    if (attachments && attachments.length) opts.attachments = attachments;
    MailApp.sendEmail(a, subject, body, (opts.htmlBody || opts.attachments) ? opts : undefined);
    logNotify_(d, a, subject, 'sent (guest)');
  }
  catch (e) { logNotify_(d, a, subject, 'ERROR (guest): ' + e.message); }
}

function auth_(d, pin) {
  pin = String(pin || '');
  if (!pin) return null;
  if (pin === d.settings.adminPin) {
    return {role: 'admin', name: 'The Crown', rooms: d.rooms.map(function (r) { return r.id; })};
  }
  var mine = d.rooms.filter(function (r) { return r.ownerPin && r.ownerPin === pin; });
  if (mine.length) {
    return {role: 'owner', name: mine[0].ownerName || 'Keeper', rooms: mine.map(function (r) { return r.id; })};
  }
  return null;
}

function overlap_(d, roomId, start, end, excludeId) {
  return d.bookings.some(function (b) {
    return b.roomId === roomId && b.id !== excludeId && b.type !== 'request' &&
           start < b.end && end > b.start;
  });
}

/** Remove [start,end) from every block overlapping it — the block splits into
 *  up to two remainder blocks. Bottom-up deletes keep row numbers honest. */
function carveBlocks_(d, roomId, start, end) {
  var sh = d.ss.getSheetByName('Bookings');
  var blocks = d.bookings.filter(function (b) {
    return b.roomId === roomId && b.type === 'block' && start < b.end && end > b.start;
  }).sort(function (a, b) { return b._row - a._row; });
  blocks.forEach(function (b) {
    sh.deleteRow(b._row);
    if (b.start < start) sh.appendRow(['b' + new Date().getTime() + 'a' + Math.floor(Math.random() * 1000),
      roomId, 'block', b.guestName, b.start, start, '', b.createdBy, today_(), '', '', '']);
    if (end < b.end) sh.appendRow(['b' + new Date().getTime() + 'b' + Math.floor(Math.random() * 1000),
      roomId, 'block', b.guestName, end, b.end, '', b.createdBy, today_(), '', '', '']);
  });
  return blocks.length;
}

function addDays_(s, n) {
  var p = s.split('-').map(Number);
  var dt = new Date(p[0], p[1] - 1, p[2] + n);
  return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
}

/** Capacity-aware conflict for GUEST bookings: a room with capacity N hosts up
 *  to N overlapping stays per night; blocks always claim the whole room. */
function bookingConflict_(d, room, start, end, excludeId, ignoreBlocks) {
  var cap = Number(room.capacity) || 1;
  var overl = d.bookings.filter(function (b) {
    return b.roomId === room.id && b.id !== excludeId && b.type !== 'request' &&
           !(ignoreBlocks && b.type === 'block') && start < b.end && end > b.start;
  });
  if (overl.some(function (b) { return b.type === 'block'; })) return true;
  if (overl.length < cap) return false;
  var day = start;
  while (day < end) {
    var n = 0;
    for (var i = 0; i < overl.length; i++) if (overl[i].start <= day && day < overl[i].end) n++;
    if (n >= cap) return true;
    day = addDays_(day, 1);
  }
  return false;
}

function publicState_(d) {
  return {
    ok: true,
    adminNotifySet: splitAddrs_(d.settings.adminNotify).length > 0,
    seam: {set: !!seamCfg_(d), lockName: String(d.settings.seamLockName || '')},
    siteUrl: String(d.settings.siteUrl || ''),
    rooms: d.rooms.map(function (r) {
      return {id: r.id, name: r.name, emoji: r.emoji, tag: r.tag, rate: r.rate,
              dark: String(r.flags).indexOf('dark') >= 0, closed: r.closed, ownerName: r.ownerName,
              ownerNotifySet: splitAddrs_(r.ownerNotify).length > 0, capacity: Number(r.capacity) || 1};
    }),
    bookings: d.bookings.map(function (b) {
      return {id: b.id, roomId: b.roomId, type: b.type, guestName: b.guestName, start: b.start, end: b.end};
    })
  };
}

function setRoom_(d, roomId, field, value) {
  var r = d.rooms.filter(function (x) { return x.id === roomId; })[0];
  if (!r) return false;
  var col = ROOM_HEAD.indexOf(field) + 1;
  d.ss.getSheetByName('Rooms').getRange(r._row, col).setValue(value);
  return true;
}

/* ---------------- actions ---------------- */

function handle_(p) {
  var d = db_();
  var action = p.action;

  if (action === 'state') return publicState_(d);

  if (action === 'login') {
    var a0 = auth_(d, p.pin);
    return a0 ? {ok: true, role: a0.role, name: a0.name, rooms: a0.rooms}
              : {ok: false, error: 'The gate does not recognize that PIN.'};
  }

  if (action === 'book') {
    var room = d.rooms.filter(function (r) { return r.id === p.roomId; })[0];
    if (!room) return {ok: false, error: 'No such chamber.'};
    if (room.closed) return {ok: false, error: 'That chamber is closed by order of the Crown.'};
    if (!p.guestName || !String(p.guestName).trim()) return {ok: false, error: 'The royal ledger requires a name.'};
    if (!p.start || !p.end || p.start >= p.end) return {ok: false, error: 'Pick a valid check-in and check-out.'};
    if (p.start < today_()) return {ok: false, error: 'The castle does not accept bookings in the past. Yet.'};
    if (bookingConflict_(d, room, p.start, p.end)) return {ok: false, error:
      (Number(room.capacity) || 1) > 1 ? 'The chamber is full for part of that range.' : 'Alas — someone claimed those nights first.'};
    var code = 'WC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
    var gEmail = String(p.guestEmail || '').trim();
    var shBk = d.ss.getSheetByName('Bookings');
    shBk.appendRow(
      ['b' + new Date().getTime(), p.roomId, 'guest', String(p.guestName).trim(),
       p.start, p.end, code, 'guest', today_(), gEmail, '', '']);
    var doorPin = createLockCode_(d, {_row: shBk.getLastRow(), code: code,
      guestName: String(p.guestName).trim(), start: p.start, end: p.end});
    var when = fmtD_(p.start) + ' → ' + fmtD_(p.end) + ' (' + nights_(p.start, p.end) + 'n)';
    notify_(d, room, '🏰 New booking: ' + roomName_(room),
      String(p.guestName).trim() + ' booked ' + roomName_(room) + ', ' + when + '. Code ' + code + '.' +
      (doorPin ? ' Door code ' + doorPin + '.' : ''));
    mailGuest_(d, gEmail, "🏰 You're booked at Wes's Castle",
      'Your chamber is secured!\n\n' + roomName_(room) + '\n' + fmtD_(p.start) + ' → ' + fmtD_(p.end) +
      ' (' + nights_(p.start, p.end) + ' nights)\n\nConfirmation code: ' + code +
      (doorPin ? '\n🔐 Door code: ' + doorPin + ' — works on the keypad all day, from check-in day through checkout day.' : '') +
      '\n\nUse your confirmation code on the castle site to view, add to calendar, or cancel your stay.' +
      '\nCheck-out is before the King starts vacuuming pointedly.',
      eH_('Thy chamber is secured! &#127881;') +
      eP_('Rejoice, ' + String(p.guestName).trim() + ' &mdash; the castle awaits.') +
      eRows_([['Chamber', roomName_(room)],
              ['Check-in', fmtD_(p.start) + ' — any time'],
              ['Check-out', fmtD_(p.end) + ' — any time'],
              ['Nights', String(nights_(p.start, p.end))]]) +
      eCode_('Confirmation code', code) +
      (doorPin ? eCode_('&#128272; Front door code', doorPin) + eLockHelp_(d) : '') +
      eCalRow_(roomName_(room), String(p.guestName).trim(), code, p.start, p.end) +
      eManageBtn_(d) +
      eKeyBtn_(d, gEmail) +
      eP_('<span style="color:#6b6078;font-size:12.5px">Your confirmation code views, calendars, or cancels this stay on the castle site. Check-out is before the King starts vacuuming pointedly.</span>'),
      [stayIcs_(roomName_(room), String(p.guestName).trim(), code, p.start, p.end)]);
    return {ok: true, code: code, doorCode: doorPin || undefined};
  }

  if (action === 'cancelByCode') {
    var code2 = String(p.code || '').trim().toUpperCase();
    var hit = d.bookings.filter(function (b) { return b.code && b.code.toUpperCase() === code2; })[0];
    if (!hit) return {ok: false, error: 'No reservation bears that code.'};
    deleteLockCode_(d, hit);
    d.ss.getSheetByName('Bookings').deleteRow(hit._row);
    var cRoom = d.rooms.filter(function (r) { return r.id === hit.roomId; })[0];
    var cWhen = fmtD_(hit.start) + ' → ' + fmtD_(hit.end);
    notify_(d, cRoom, '🏰 Cancelled: ' + roomName_(cRoom),
      hit.guestName + "'s stay " + cWhen + ' was cancelled. Dates are free again.');
    mailGuest_(d, hit.guestEmail, "🏰 Your Wes's Castle stay is cancelled",
      'Your reservation in ' + roomName_(cRoom) + ' (' + cWhen + ') has been cancelled.' +
      '\nThe castle mourns, briefly, then re-lists the chamber.',
      eH_('Thy stay is cancelled') +
      eP_('Your reservation has been struck from the royal ledger:') +
      eRows_([['Chamber', roomName_(cRoom)], ['Dates', cWhen]]) +
      (String(hit.doorCode || '') ? eP_('Your door code has been revoked from the lock.') : '') +
      eP_('The castle mourns, briefly, then re-lists the chamber. Book again any time.') +
      eManageBtn_(d));
    return {ok: true};
  }

  if (action === 'findByCode') {
    var code3 = String(p.code || '').trim().toUpperCase();
    var hit2 = d.bookings.filter(function (b) { return b.code && b.code.toUpperCase() === code3; })[0];
    if (!hit2) return {ok: false, error: 'No reservation bears that code.'};
    return {ok: true, booking: {roomId: hit2.roomId, guestName: hit2.guestName,
      start: hit2.start, end: hit2.end, code: hit2.code, pending: hit2.type === 'request',
      doorCode: String(hit2.doorCode || '') || undefined}};
  }

  if (action === 'requestBook') {
    var rqRoom = d.rooms.filter(function (r) { return r.id === p.roomId; })[0];
    if (!rqRoom) return {ok: false, error: 'No such chamber.'};
    if (rqRoom.closed) return {ok: false, error: 'That chamber is closed by order of the Crown.'};
    if (!p.guestName || !String(p.guestName).trim()) return {ok: false, error: 'The royal ledger requires a name.'};
    if (!p.start || !p.end || p.start >= p.end) return {ok: false, error: 'Pick a valid check-in and check-out.'};
    if (p.start < today_()) return {ok: false, error: 'The castle does not accept requests for the past. Yet.'};
    if (bookingConflict_(d, rqRoom, p.start, p.end, null, true))
      return {ok: false, error: 'Other guests already hold those nights — a request cannot help there.'};
    var rqCode = 'WC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
    var rqEmail = String(p.guestEmail || '').trim();
    d.ss.getSheetByName('Bookings').appendRow(
      ['b' + new Date().getTime(), p.roomId, 'request', String(p.guestName).trim(),
       p.start, p.end, rqCode, 'guest', today_(), rqEmail, '', '']);
    var rqWhen = fmtD_(p.start) + ' → ' + fmtD_(p.end) + ' (' + nights_(p.start, p.end) + 'n)';
    notify_(d, rqRoom, '🙏 Request: ' + roomName_(rqRoom),
      String(p.guestName).trim() + ' asks for ' + rqWhen + ' (currently blocked). ' +
      "Approve or deny in the Keeper's Entrance. Code " + rqCode + '.');
    mailGuest_(d, rqEmail, "🏰 Thy request is with the Crown",
      'Your request is in:\n\n' + roomName_(rqRoom) + '\n' + rqWhen +
      '\n\nCode: ' + rqCode + '\n\nThe castle will send word once the keeper decides.',
      eH_('&#128591; Thy request is with the Crown') +
      eP_('Those dates are held by the castle, but your plea has been heard:') +
      eRows_([['Chamber', roomName_(rqRoom)], ['Dates requested', rqWhen]]) +
      eCode_('Request code', rqCode) +
      eP_('The Crown and the chamber&#8217;s keeper will decide, and a raven will bring their verdict. If approved, this code becomes your booking code.') +
      eManageBtn_(d));
    return {ok: true, code: rqCode, pending: true};
  }

  /* ----- guest accounts (email + password; emailed code for resets) ----- */

  if (action === 'signup') {
    var suEmail = String(p.email || '').trim().toLowerCase();
    var suPass = String(p.password || '');
    if (suEmail.indexOf('@') < 1) return {ok: false, error: 'Enter a real email address.'};
    if (suPass.length < 6) return {ok: false, error: 'Passwords need at least 6 characters.'};
    var shG = d.ss.getSheetByName('Guests');
    var exists = readAll_(shG).filter(function (g) { return String(g.email).toLowerCase() === suEmail; })[0];
    if (exists) return {ok: false, error: 'That email already holds a castle key — sign in, or reset thy password.'};
    var salt = Utilities.getUuid();
    shG.appendRow([suEmail, hashPass_(salt, suPass), salt, String(p.name || '').trim(), new Date().getTime()]);
    logNotify_(d, suEmail, 'castle key', 'account created');
    return {ok: true, token: makeToken_(d, suEmail), email: suEmail, name: String(p.name || '').trim()};
  }

  if (action === 'signin') {
    var siEmail = String(p.email || '').trim().toLowerCase();
    var gRow = readAll_(d.ss.getSheetByName('Guests')).filter(function (g) {
      return String(g.email).toLowerCase() === siEmail;
    })[0];
    if (!gRow || hashPass_(String(gRow.salt), String(p.password || '')) !== String(gRow.passHash))
      return {ok: false, error: 'The gate does not recognize that email & password.'};
    return {ok: true, token: makeToken_(d, siEmail), email: siEmail, name: String(gRow.name || '')};
  }

  if (action === 'resetStart') {
    var rsEmail = String(p.email || '').trim().toLowerCase();
    var rsGuest = readAll_(d.ss.getSheetByName('Guests')).filter(function (g) {
      return String(g.email).toLowerCase() === rsEmail;
    })[0];
    if (!rsGuest) return {ok: false, error: 'No castle key is held by that email. Create one instead.'};
    var shS = d.ss.getSheetByName('Sessions');
    var code6 = String(Math.floor(100000 + Math.random() * 900000));
    var expS = new Date().getTime() + 10 * 60 * 1000;
    var pendingRow = readAll_(shS).filter(function (s) {
      return String(s.email).toLowerCase() === rsEmail && !String(s.token);
    })[0];
    if (pendingRow) {
      shS.getRange(pendingRow._row, 3).setValue(code6);
      shS.getRange(pendingRow._row, 4).setValue(expS);
    } else {
      shS.appendRow(['', rsEmail, code6, expS, new Date().getTime()]);
    }
    try {
      MailApp.sendEmail(rsEmail, '🏰 Password reset code: ' + code6,
        'Speak this at the gate to set a new password: ' + code6 +
        '\n\nIt expires in 10 minutes. If thou didst not request it, ignore this raven.',
        {htmlBody: eWrap_(
          eH_('&#128273; Reset thy password') +
          eP_('Speak this code at the gate to set a new password:') +
          eCode_('Reset code &middot; expires in 10 minutes', code6) +
          eP_('<span style="color:#6b6078;font-size:12.5px">If thou didst not request it, ignore this raven and thy password stands.</span>'))});
      logNotify_(d, rsEmail, 'reset code', 'sent');
    } catch (e) {
      logNotify_(d, rsEmail, 'reset code', 'ERROR: ' + e.message);
      return {ok: false, error: 'Could not send the code: ' + e.message};
    }
    return {ok: true};
  }

  if (action === 'resetFinish') {
    var rfEmail = String(p.email || '').trim().toLowerCase();
    var rfPass = String(p.password || '');
    if (rfPass.length < 6) return {ok: false, error: 'Passwords need at least 6 characters.'};
    var shS2 = d.ss.getSheetByName('Sessions');
    var nowR = new Date().getTime();
    var codeRow = readAll_(shS2).filter(function (s) {
      return String(s.email).toLowerCase() === rfEmail && String(s.code) &&
             String(s.code) === String(p.code || '').trim() && Number(s.expires) > nowR;
    })[0];
    if (!codeRow) return {ok: false, error: 'That code is wrong or expired. Request a fresh one.'};
    var shG2 = d.ss.getSheetByName('Guests');
    var gRow2 = readAll_(shG2).filter(function (g) { return String(g.email).toLowerCase() === rfEmail; })[0];
    if (!gRow2) return {ok: false, error: 'No castle key is held by that email.'};
    var salt2 = Utilities.getUuid();
    shG2.getRange(gRow2._row, 2).setValue(hashPass_(salt2, rfPass));
    shG2.getRange(gRow2._row, 3).setValue(salt2);
    shS2.getRange(codeRow._row, 3).setValue('');
    logNotify_(d, rfEmail, 'castle key', 'password reset');
    return {ok: true, token: makeToken_(d, rfEmail), email: rfEmail, name: String(gRow2.name || '')};
  }

  if (action === 'myStays') {
    var sessRow = readAll_(d.ss.getSheetByName('Sessions')).filter(function (s) {
      return String(s.token) && String(s.token) === String(p.token || '');
    })[0];
    if (!sessRow || Number(sessRow.expires) < new Date().getTime())
      return {ok: false, error: 'Thy castle key expired — sign in again.', expired: true};
    var em = String(sessRow.email).toLowerCase();
    var mine = d.bookings.filter(function (b) {
      return (b.type === 'guest' || b.type === 'request') && String(b.guestEmail || '').toLowerCase() === em;
    }).map(function (b) {
      return {roomId: b.roomId, guestName: b.guestName, start: b.start, end: b.end, code: b.code,
              pending: b.type === 'request', doorCode: String(b.doorCode || '') || undefined};
    });
    return {ok: true, email: em, stays: mine};
  }

  /* ----- authorized ops ----- */
  var a = auth_(d, p.pin);
  if (!a) return {ok: false, error: 'The gate does not recognize that PIN.'};
  function canRoom(id) { return a.role === 'admin' || a.rooms.indexOf(id) >= 0; }

  if (action === 'block') {
    if (!canRoom(p.roomId)) return {ok: false, error: 'That chamber is not yours to command.'};
    if (!p.start || !p.end || p.start >= p.end) return {ok: false, error: 'Pick a valid date range.'};
    if (overlap_(d, p.roomId, p.start, p.end)) return {ok: false, error: 'Those dates already have a booking or block. Release it first.'};
    d.ss.getSheetByName('Bookings').appendRow(
      ['b' + new Date().getTime(), p.roomId, 'block', String(p.note || '').trim() || 'Blocked',
       p.start, p.end, '', a.role, today_(), '', '', '']);
    var bRoom = d.rooms.filter(function (r) { return r.id === p.roomId; })[0];
    notify_(d, bRoom, '🏰 Dates blocked: ' + roomName_(bRoom),
      fmtD_(p.start) + ' → ' + fmtD_(p.end) + ' blocked by ' + a.name +
      (p.note ? ' — ' + String(p.note).trim() : '') + '.');
    return {ok: true};
  }

  if (action === 'unbook') {
    var b2 = d.bookings.filter(function (b) { return b.id === String(p.bookingId); })[0];
    if (!b2) return {ok: false, error: 'Reservation not found.'};
    if (!canRoom(b2.roomId)) return {ok: false, error: 'That chamber is not yours to command.'};
    deleteLockCode_(d, b2);
    d.ss.getSheetByName('Bookings').deleteRow(b2._row);
    var uRoom = d.rooms.filter(function (r) { return r.id === b2.roomId; })[0];
    var uWhen = fmtD_(b2.start) + ' → ' + fmtD_(b2.end);
    notify_(d, uRoom, '🏰 Released: ' + roomName_(uRoom),
      (b2.type === 'block' ? 'Block "' + b2.guestName + '"' : b2.guestName + "'s stay") +
      ' (' + uWhen + ') was released by ' + a.name + '.');
    mailGuest_(d, b2.guestEmail, "🏰 Your Wes's Castle stay is cancelled",
      'Your reservation in ' + roomName_(uRoom) + ' (' + uWhen + ') was cancelled by the castle.' +
      '\nQuestions? Reply to this email and the Crown shall answer.',
      eH_('Thy stay was cancelled by the castle') +
      eRows_([['Chamber', roomName_(uRoom)], ['Dates', uWhen]]) +
      (String(b2.doorCode || '') ? eP_('Your door code has been revoked from the lock.') : '') +
      eP_('Questions? Reply to this email and the Crown shall answer.'));
    return {ok: true};
  }

  if (action === 'unblockRange') {
    var ub = d.bookings.filter(function (b) { return b.id === String(p.bookingId); })[0];
    if (!ub || ub.type !== 'block') return {ok: false, error: 'Block not found.'};
    if (!canRoom(ub.roomId)) return {ok: false, error: 'That chamber is not yours to command.'};
    var uS = p.start > ub.start ? p.start : ub.start;
    var uE = p.end < ub.end ? p.end : ub.end;
    if (!p.start || !p.end || uS >= uE) return {ok: false, error: 'Pick dates inside the block.'};
    var ubRoom = d.rooms.filter(function (r) { return r.id === ub.roomId; })[0];
    carveBlocks_(d, ub.roomId, uS, uE);
    notify_(d, ubRoom, '🏰 Unblocked: ' + roomName_(ubRoom),
      fmtD_(uS) + ' → ' + fmtD_(uE) + ' released from block "' + ub.guestName + '" by ' + a.name +
      '. The rest of the block stands.');
    return {ok: true};
  }

  if (action === 'approveRequest') {
    var apr = d.bookings.filter(function (b) { return b.id === String(p.bookingId); })[0];
    if (!apr || apr.type !== 'request') return {ok: false, error: 'Request not found — it may already be decided.'};
    if (!canRoom(apr.roomId)) return {ok: false, error: 'That chamber is not yours to command.'};
    var aprRoom = d.rooms.filter(function (r) { return r.id === apr.roomId; })[0];
    if (bookingConflict_(d, aprRoom, apr.start, apr.end, apr.id, true))
      return {ok: false, error: 'Other guests now hold those nights — the request cannot be approved as-is.'};
    // convert + program the door BEFORE carving (deletes above would shift this row)
    d.ss.getSheetByName('Bookings').getRange(apr._row, BOOKING_HEAD.indexOf('type') + 1).setValue('guest');
    var aprPin = createLockCode_(d, {_row: apr._row, code: apr.code,
      guestName: apr.guestName, start: apr.start, end: apr.end});
    carveBlocks_(d, apr.roomId, apr.start, apr.end);
    var aprWhen = fmtD_(apr.start) + ' → ' + fmtD_(apr.end) + ' (' + nights_(apr.start, apr.end) + 'n)';
    notify_(d, aprRoom, '✅ Approved: ' + roomName_(aprRoom),
      apr.guestName + ' is now booked ' + aprWhen + ' (request approved by ' + a.name + ').' +
      (aprPin ? ' Door code ' + aprPin + '.' : ''));
    mailGuest_(d, apr.guestEmail, "🏰 Thy request is GRANTED",
      'Rejoice! The castle approved your stay:\n\n' + roomName_(aprRoom) + '\n' + aprWhen +
      (aprPin ? '\n🔐 Door code: ' + aprPin + ' — works on the keypad all day, from check-in day through checkout day.' : '') +
      '\n\nYour confirmation code is unchanged: ' + apr.code,
      eH_('&#9989; Thy request is GRANTED') +
      eP_('Rejoice, ' + apr.guestName + '! The Crown has parted the block for thee:') +
      eRows_([['Chamber', roomName_(aprRoom)],
              ['Check-in', fmtD_(apr.start) + ' — any time'],
              ['Check-out', fmtD_(apr.end) + ' — any time']]) +
      eCode_('Confirmation code', apr.code) +
      (aprPin ? eCode_('&#128272; Front door code', aprPin) + eLockHelp_(d) : '') +
      eCalRow_(roomName_(aprRoom), apr.guestName, apr.code, apr.start, apr.end) +
      eManageBtn_(d) +
      eKeyBtn_(d, apr.guestEmail),
      [stayIcs_(roomName_(aprRoom), apr.guestName, apr.code, apr.start, apr.end)]);
    return {ok: true};
  }

  if (action === 'denyRequest') {
    var den = d.bookings.filter(function (b) { return b.id === String(p.bookingId); })[0];
    if (!den || den.type !== 'request') return {ok: false, error: 'Request not found — it may already be decided.'};
    if (!canRoom(den.roomId)) return {ok: false, error: 'That chamber is not yours to command.'};
    var denRoom = d.rooms.filter(function (r) { return r.id === den.roomId; })[0];
    d.ss.getSheetByName('Bookings').deleteRow(den._row);
    var denWhen = fmtD_(den.start) + ' → ' + fmtD_(den.end);
    notify_(d, denRoom, '❌ Denied: ' + roomName_(denRoom),
      den.guestName + "'s request for " + denWhen + ' was denied by ' + a.name + '.');
    mailGuest_(d, den.guestEmail, "🏰 About thy request…",
      'Alas — the castle cannot host you ' + denWhen + ' in ' + roomName_(denRoom) +
      '. Those dates remain spoken for.\n\nPerhaps other dates? The calendar awaits.',
      eH_('Alas &mdash; about thy request') +
      eP_('The castle cannot host you these dates; they remain spoken for:') +
      eRows_([['Chamber', roomName_(denRoom)], ['Dates', denWhen]]) +
      eP_('The Crown regrets the news. Perhaps other dates? The calendar awaits thee.') +
      eManageBtn_(d));
    return {ok: true};
  }

  if (action === 'editBooking') {
    var eb = d.bookings.filter(function (b) { return b.id === String(p.bookingId); })[0];
    if (!eb) return {ok: false, error: 'Reservation not found.'};
    var newRoom = d.rooms.filter(function (r) { return r.id === p.roomId; })[0];
    if (!newRoom) return {ok: false, error: 'No such chamber.'};
    if (!canRoom(eb.roomId) || !canRoom(newRoom.id)) return {ok: false, error: 'That chamber is not yours to command.'};
    if (!p.start || !p.end || p.start >= p.end) return {ok: false, error: 'Pick a valid date range.'};
    var newName = String(p.guestName || '').trim() || eb.guestName;
    var clash = eb.type === 'block'
      ? overlap_(d, newRoom.id, p.start, p.end, eb.id)
      : bookingConflict_(d, newRoom, p.start, p.end, eb.id);
    if (clash) return {ok: false, error: 'Those dates clash with another booking or block.'};
    var shE = d.ss.getSheetByName('Bookings');
    shE.getRange(eb._row, BOOKING_HEAD.indexOf('roomId') + 1).setValue(newRoom.id);
    shE.getRange(eb._row, BOOKING_HEAD.indexOf('guestName') + 1).setValue(newName);
    shE.getRange(eb._row, BOOKING_HEAD.indexOf('start') + 1).setValue(p.start);
    shE.getRange(eb._row, BOOKING_HEAD.indexOf('end') + 1).setValue(p.end);
    updateLockCode_(d, eb, p.start, p.end);
    var oldRoomE = d.rooms.filter(function (r) { return r.id === eb.roomId; })[0];
    var movedE = newRoom.id !== eb.roomId;
    var whenE = fmtD_(p.start) + ' → ' + fmtD_(p.end) + ' (' + nights_(p.start, p.end) + 'n)';
    notify_(d, newRoom, '🏰 Changed: ' + roomName_(newRoom),
      newName + (eb.type === 'block' ? ' (block)' : '') + ' is now ' + whenE +
      (movedE ? ', moved from ' + roomName_(oldRoomE) : '') + '. Changed by ' + a.name + '.',
      movedE ? oldRoomE : null);
    mailGuest_(d, eb.guestEmail, "🏰 Thy Wes's Castle stay was updated",
      'The castle updated your reservation:\n\n' + roomName_(newRoom) + '\n' + whenE +
      (String(eb.doorCode || '') ? '\n🔐 Your door code ' + eb.doorCode + ' now follows the new dates.' : '') +
      '\n\nYour confirmation code is unchanged: ' + eb.code,
      eH_('&#9999;&#65039; Thy stay was updated') +
      eP_('The castle adjusted your reservation. The new arrangement:') +
      eRows_([['Chamber', roomName_(newRoom)],
              ['Check-in', fmtD_(p.start) + ' — any time'],
              ['Check-out', fmtD_(p.end) + ' — any time']]) +
      eCode_('Confirmation code (unchanged)', eb.code) +
      (String(eb.doorCode || '') ? eCode_('&#128272; Door code (follows new dates)', eb.doorCode) : '') +
      eCalRow_(roomName_(newRoom), newName, eb.code, p.start, p.end) +
      eP_('<span style="color:#6b6078;font-size:12.5px">The attached calendar invite carries the new dates — adding it replaces the old entry.</span>') +
      eManageBtn_(d),
      [stayIcs_(roomName_(newRoom), newName, eb.code, p.start, p.end)]);
    return {ok: true};
  }

  /* ----- admin only ----- */
  if (a.role !== 'admin') return {ok: false, error: 'Only the Crown may do that.'};

  if (action === 'assignOwner') {
    var name = String(p.ownerName || '').trim();
    if (!setRoom_(d, p.roomId, 'ownerName', name)) return {ok: false, error: 'No such chamber.'};
    if (!name) { setRoom_(d, p.roomId, 'ownerPin', ''); setRoom_(d, p.roomId, 'ownerNotify', ''); }
    else {
      if (!p.keepPin) setRoom_(d, p.roomId, 'ownerPin', String(p.ownerPin || '').trim());
      if (p.ownerNotify !== undefined && !p.keepNotify)
        setRoom_(d, p.roomId, 'ownerNotify', String(p.ownerNotify || '').trim());
    }
    if (p.capacity !== undefined) {
      var cap2 = Math.max(1, Math.min(20, parseInt(p.capacity, 10) || 1));
      setRoom_(d, p.roomId, 'capacity', cap2);
    }
    return {ok: true};
  }

  if (action === 'testNotify') {
    var tAddrs = splitAddrs_(d.settings.adminNotify);
    if (!tAddrs.length) return {ok: false, error: "Set the Crown's notification address first, then test."};
    try {
      MailApp.sendEmail(tAddrs[0], "🏰 Test decree from Wes's Castle",
        'Hear ye! Notifications are working. Email quota left today: ' + MailApp.getRemainingDailyQuota() + '.');
      return {ok: true, sent: tAddrs[0].replace(/^(..)[^@]*(@.*)$/, '$1…$2')};
    } catch (e) {
      return {ok: false, error: 'Google blocked the send (' + e.message + '). Fix: open the Apps Script editor, ' +
        'pick "authorizeCastle" in the toolbar dropdown, press Run, and approve the permission dialog.'};
    }
  }

  if (action === 'seamConnect') {
    var sKey = String(p.key || '').trim();
    if (!sKey) { // disconnect
      setSetting_(d, 'seamKey', ''); setSetting_(d, 'seamDeviceId', ''); setSetting_(d, 'seamLockName', '');
      return {ok: true, disconnected: true};
    }
    var devRes = seamFetch_(d, '/devices/list', {}, sKey);
    var locks = (devRes.devices || []).filter(function (x) { return x.can_program_online_access_codes; });
    if (!locks.length) return {ok: false, error: 'Seam answered, but no code-capable locks were found (' +
      ((devRes.error && devRes.error.message) || 'check the API key') + ').'};
    var pick = p.deviceId ? locks.filter(function (x) { return x.device_id === p.deviceId; })[0] : null;
    if (!pick && locks.length > 1)
      return {ok: true, choose: locks.map(function (x) {
        return {id: x.device_id, name: (x.properties || {}).name || x.device_type,
                online: !!(x.properties || {}).online};
      })};
    pick = pick || locks[0];
    var pickName = (pick.properties || {}).name || 'Lock';
    setSetting_(d, 'seamKey', sKey);
    setSetting_(d, 'seamDeviceId', pick.device_id);
    setSetting_(d, 'seamLockName', pickName);
    logNotify_(d, 'lock', 'Seam connected', pickName + ' (' + pick.device_id.slice(0, 8) + '…)');
    return {ok: true, lockName: pickName};
  }

  if (action === 'setSiteUrl') {
    setSetting_(d, 'siteUrl', String(p.value || '').trim());
    return {ok: true};
  }

  if (action === 'setAdminNotify') {
    var sh2 = d.ss.getSheetByName('Settings');
    var rows2 = readAll_(sh2);
    var row2 = rows2.filter(function (s) { return s.key === 'adminNotify'; })[0];
    var val = String(p.value || '').trim();
    if (row2) sh2.getRange(row2._row, 2).setValue(val);
    else sh2.appendRow(['adminNotify', val]);
    return {ok: true, set: splitAddrs_(val).length > 0};
  }

  if (action === 'toggleClosed') {
    var r3 = d.rooms.filter(function (x) { return x.id === p.roomId; })[0];
    if (!r3) return {ok: false, error: 'No such chamber.'};
    setRoom_(d, p.roomId, 'closed', r3.closed ? 'FALSE' : 'TRUE');
    return {ok: true, closed: !r3.closed};
  }

  if (action === 'setAdminPin') {
    var np = String(p.newPin || '').trim();
    if (np.length < 4) return {ok: false, error: 'PIN must be at least 4 characters.'};
    var sh = d.ss.getSheetByName('Settings');
    var rows = readAll_(sh);
    var row = rows.filter(function (s) { return s.key === 'adminPin'; })[0];
    if (row) sh.getRange(row._row, 2).setValue(np);
    else sh.appendRow(['adminPin', np]);
    return {ok: true};
  }

  return {ok: false, error: 'Unknown decree: ' + action};
}
