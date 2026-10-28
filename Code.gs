/**
 * FIN 4113 class backend (Google Apps Script, bound to a Google Sheet).
 *
 * Students can: load the roster dropdown, claim a name (one change allowed),
 * send a discussion question, and submit answers.
 * The instructor (with the key) can: read everything, log cold calls,
 * mark questions used, release a claimed name, and push the roster.
 *
 * No secrets live in this file. The instructor key is stored in
 * Script Properties by running setup() once from the editor.
 */

var TABS = {
  Roster:    ['Name', 'Group'],
  Claims:    ['Timestamp', 'Session', 'Name', 'Device', 'Status'],
  Responses: ['Timestamp', 'Session', 'Name', 'Device'],
  Questions: ['Timestamp', 'Session', 'Name', 'Device', 'Question', 'Used'],
  CallLog:   ['Timestamp', 'Session', 'Mode', 'ItemId', 'ItemText', 'Name', 'Outcome', 'Note']
};
var MAX_LEN = 5000;

/* ---------- one-time setup: run this from the editor ---------- */
function setup() {
  ensureTabs_();
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('INSTRUCTOR_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('INSTRUCTOR_KEY', key);
  }
  Logger.log('Instructor key: ' + key);
  return key;
}

/* ---------- entry points ---------- */
function doGet(e) {
  return json_({ ok: true, service: 'FIN4113', time: new Date().toISOString() });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    switch (req.action) {
      case 'roster':    return json_(roster_());
      case 'whoami':    return json_(whoami_(req));
      case 'claim':     return json_(locked_(function () { return claim_(req); }));
      case 'question':  return json_(locked_(function () { return question_(req); }));
      case 'submit':    return json_(locked_(function () { return submit_(req); }));
      // instructor only
      case 'admin_state': admin_(req); return json_(adminState_(req));
      case 'log_call':    admin_(req); return json_(locked_(function () { return logCall_(req); }));
      case 'mark_used':   admin_(req); return json_(locked_(function () { return markUsed_(req); }));
      case 'release':     admin_(req); return json_(locked_(function () { return release_(req); }));
      case 'set_roster':  admin_(req); return json_(locked_(function () { return setRoster_(req); }));
      default: return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- student actions ---------- */
function roster_() {
  var rows = rows_('Roster');
  var names = rows.map(function (r) { return String(r[0]).trim(); }).filter(String);
  names.sort(function (a, b) { return a.localeCompare(b); });
  return { ok: true, names: names };
}

function whoami_(req) {
  need_(req, ['session', 'device']);
  var c = claimState_(req.session, req.device);
  return { ok: true, name: c.active ? c.active.name : null, changesLeft: c.changesLeft };
}

function claim_(req) {
  need_(req, ['session', 'device', 'name']);
  var name = String(req.name).trim();
  if (roster_().names.indexOf(name) < 0) return { ok: false, error: 'not_on_roster' };

  var sh = tab_('Claims');
  var data = sh.getDataRange().getValues();
  var mine = claimState_(req.session, req.device, data);

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[1] === req.session && r[2] === name && r[4] === 'active') {
      if (r[3] === req.device) return { ok: true, name: name, changesLeft: mine.changesLeft };
      return { ok: false, error: 'taken' };
    }
  }
  if (mine.active) {
    if (mine.changesLeft < 1) return { ok: false, error: 'no_changes_left', name: mine.active.name };
    sh.getRange(mine.active.row, 5).setValue('changed');
    mine.changesLeft -= 1;
  }
  sh.appendRow([new Date(), req.session, name, req.device, 'active']);
  return { ok: true, name: name, changesLeft: mine.changesLeft };
}

function question_(req) {
  need_(req, ['session', 'device', 'name', 'question']);
  verify_(req);
  var q = clean_(req.question);
  if (q.length < 5) return { ok: false, error: 'question_too_short' };
  tab_('Questions').appendRow([new Date(), req.session, req.name, req.device, q, '']);
  return { ok: true };
}

function submit_(req) {
  need_(req, ['session', 'device', 'name', 'answers']);
  verify_(req);
  var sh = tab_('Responses');
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var keys = Object.keys(req.answers);
  keys.forEach(function (k) {
    if (header.indexOf(k) < 0) {
      header.push(k);
      sh.getRange(1, header.length).setValue(k);
    }
  });
  var row = header.map(function () { return ''; });
  row[0] = new Date(); row[1] = req.session; row[2] = req.name; row[3] = req.device;
  keys.forEach(function (k) { row[header.indexOf(k)] = clean_(req.answers[k]); });
  sh.appendRow(row);
  return { ok: true, received: keys.length };
}

/* ---------- instructor actions ---------- */
function adminState_(req) {
  need_(req, ['session']);
  var s = req.session;
  var pick = function (name) {
    var v = tab_(name).getDataRange().getValues();
    var head = v.shift();
    var out = [];
    v.forEach(function (r, i) {
      if (r[1] !== s && name !== 'Roster') return;
      var o = { _row: i + 2 };
      head.forEach(function (h, j) { o[h] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
      out.push(o);
    });
    return out;
  };
  return {
    ok: true,
    roster: pick('Roster'),
    claims: pick('Claims'),
    responses: pick('Responses'),
    questions: pick('Questions'),
    calllog: pick('CallLog')
  };
}

function logCall_(req) {
  need_(req, ['session', 'mode', 'name', 'outcome']);
  tab_('CallLog').appendRow([new Date(), req.session, clean_(req.mode), clean_(req.itemId || ''),
    clean_(req.itemText || ''), clean_(req.name), clean_(req.outcome), clean_(req.note || '')]);
  return { ok: true };
}

function markUsed_(req) {
  need_(req, ['row']);
  tab_('Questions').getRange(Number(req.row), 6).setValue(req.used === false ? '' : 'yes');
  return { ok: true };
}

function release_(req) {
  need_(req, ['session', 'name']);
  var sh = tab_('Claims');
  var data = sh.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === req.session && data[i][2] === req.name && data[i][4] === 'active') {
      sh.getRange(i + 1, 5).setValue('released'); n++;
    }
  }
  return { ok: true, released: n };
}

function setRoster_(req) {
  need_(req, ['roster']);
  var sh = tab_('Roster');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  var rows = req.roster.map(function (r) { return [clean_(r.name), clean_(r.group || '')]; });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true, count: rows.length };
}

/* ---------- helpers ---------- */
function claimState_(session, device, data) {
  data = data || tab_('Claims').getDataRange().getValues();
  var active = null, changed = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[1] !== session || r[3] !== device) continue;
    if (r[4] === 'active') active = { name: r[2], row: i + 1 };
    if (r[4] === 'changed') changed++;
  }
  return { active: active, changesLeft: Math.max(0, 1 - changed) };
}

function verify_(req) {
  var c = claimState_(req.session, req.device);
  if (!c.active || c.active.name !== req.name) throw new Error('not_claimed');
}

function admin_(req) {
  var key = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_KEY');
  if (!key || req.key !== key) throw new Error('bad_key');
}

function locked_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function need_(req, fields) {
  fields.forEach(function (f) {
    if (req[f] === undefined || req[f] === null || req[f] === '') throw new Error('missing_' + f);
  });
}

function clean_(v) {
  var s = String(v == null ? '' : v).slice(0, MAX_LEN);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // block formula injection
}

function ensureTabs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
}

function tab_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) { ensureTabs_(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  return sh;
}

function rows_(name) {
  var sh = tab_(name);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
