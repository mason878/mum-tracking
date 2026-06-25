/**
 * Mum Graphical Tracking — Apps Script backend (v2: readings + events + photos)
 * Reads/writes the bound Google Sheet; stores photos privately in Drive and
 * serves them back through the web app (no public sharing).
 */

var CROPS_SHEET = 'Crops';
var READINGS_SHEET = 'Readings';
var EVENTS_SHEET = 'Events';
var PHOTOS_SHEET = 'Photos';
var PHOTO_FOLDER_ID = '1sl7HWRGcf5kmU6tTeqjRdqgU965j8KGk';

var READINGS_HEADER = ['Crop ID', 'Week', 'Actual Height (in)', 'Logged By', 'Timestamp'];
var EVENTS_HEADER = ['Crop ID', 'Week', 'Date', 'Type', 'Note', 'Logged By', 'Timestamp'];
var PHOTOS_HEADER = ['Crop ID', 'Week', 'Date', 'File ID', 'File Name', 'Note', 'Logged By', 'Timestamp'];

/** Run this once from the editor after pasting, to grant the Drive permission. */
function authorize() {
  DriveApp.getFolderById(PHOTO_FOLDER_ID);
  SpreadsheetApp.getActiveSpreadsheet().getName();
  return 'authorized';
}

/* ---------- sheet helpers ---------- */
function ensureSheet(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function ensureReadings() { return ensureSheet(READINGS_SHEET, READINGS_HEADER); }
function ensureEvents()   { return ensureSheet(EVENTS_SHEET, EVENTS_HEADER); }
function ensurePhotos()   { return ensureSheet(PHOTOS_SHEET, PHOTOS_HEADER); }

/* ---------- routing ---------- */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'all';
  try {
    if (action === 'all') return _json(getAll());
    if (action === 'photo') return _json(getPhoto(e.parameter.id));
    return _json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (b.action === 'save') return _json({ ok: true, saved: saveReading(b.crop, b.week, b.actual, b.loggedBy) });
    if (b.action === 'saveBatch' && Array.isArray(b.readings)) {
      b.readings.forEach(function (r) { saveReading(r.crop, r.week, r.actual, b.loggedBy); });
      return _json({ ok: true, count: b.readings.length });
    }
    if (b.action === 'saveEvent') return _json(saveEvent(b));
    if (b.action === 'savePhoto') return _json(savePhoto(b));
    return _json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

/* ---------- data ---------- */
function cropsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CROPS_SHEET);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var n = all[i].getName();
    if (n !== READINGS_SHEET && n !== EVENTS_SHEET && n !== PHOTOS_SHEET) return all[i];
  }
  return all[0];
}

function getAll() {
  var crops = _rows(cropsSheet()).map(function (r) {
    return {
      id: String(r['Crop ID']), name: String(r['Name']), pot: String(r['Pot']),
      grower: String(r['Grower']), location: String(r['Location'] || ''),
      startWeek: Number(r['Start Week']), weeksToFinish: Number(r['Weeks to Finish']),
      startHeight: Number(r['Start Height (in)']), finishHeight: Number(r['Finish Height (in)']),
      finishWeek: Number(r['Finish Week']), order: Number(r['Order'])
    };
  });
  var readings = _rows(ensureReadings()).map(function (r) {
    return { crop: String(r['Crop ID']), week: Number(r['Week']),
      actual: r['Actual Height (in)'] === '' ? null : Number(r['Actual Height (in)']),
      loggedBy: String(r['Logged By'] || ''), timestamp: String(r['Timestamp'] || '') };
  }).filter(function (r) { return r.actual !== null && !isNaN(r.actual); });
  var events = _rows(ensureEvents()).map(function (r) {
    return { crop: String(r['Crop ID']), week: Number(r['Week']), date: String(r['Date'] || ''),
      type: String(r['Type'] || ''), note: String(r['Note'] || ''), loggedBy: String(r['Logged By'] || '') };
  });
  var photos = _rows(ensurePhotos()).map(function (r) {
    return { crop: String(r['Crop ID']), week: Number(r['Week']), date: String(r['Date'] || ''),
      fileId: String(r['File ID'] || ''), note: String(r['Note'] || ''), loggedBy: String(r['Logged By'] || '') };
  });
  return { ok: true, crops: crops, readings: readings, events: events, photos: photos, serverTime: new Date().toISOString() };
}

/** Upsert one reading keyed on (crop, week). */
function saveReading(crop, week, actual, loggedBy) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ensureReadings(); var data = sh.getDataRange().getValues(); var h = data[0];
    var ci = h.indexOf('Crop ID'), wi = h.indexOf('Week'), ai = h.indexOf('Actual Height (in)'),
        li = h.indexOf('Logged By'), ti = h.indexOf('Timestamp'), ts = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ci]) === String(crop) && Number(data[i][wi]) === Number(week)) {
        sh.getRange(i + 1, ai + 1).setValue(actual);
        sh.getRange(i + 1, li + 1).setValue(loggedBy || 'web');
        sh.getRange(i + 1, ti + 1).setValue(ts);
        return { crop: crop, week: week, actual: actual, mode: 'update' };
      }
    }
    sh.appendRow([crop, week, actual, loggedBy || 'web', ts]);
    return { crop: crop, week: week, actual: actual, mode: 'insert' };
  } finally { lock.releaseLock(); }
}

/** Log an event to one crop or many (body.crops array, or body.crop). */
function saveEvent(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ensureEvents();
    var crops = Array.isArray(b.crops) ? b.crops : [b.crop];
    var ts = new Date().toISOString();
    var date = b.date || ts.slice(0, 10);
    crops.forEach(function (c) {
      sh.appendRow([c, b.week || '', date, b.type || '', b.note || '', b.loggedBy || 'web', ts]);
    });
    return { ok: true, count: crops.length };
  } finally { lock.releaseLock(); }
}

/** Save a photo (base64) to the Drive folder, record a row. File stays private. */
function savePhoto(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var bytes = Utilities.base64Decode(b.b64);
    var mime = b.mime || 'image/jpeg';
    var safe = String(b.crop || 'crop').replace(/[^A-Za-z0-9]+/g, '_');
    var name = safe + '_wk' + (b.week || '') + '_' + new Date().getTime() + '.jpg';
    var blob = Utilities.newBlob(bytes, mime, name);
    var file = DriveApp.getFolderById(PHOTO_FOLDER_ID).createFile(blob);
    var ts = new Date().toISOString();
    ensurePhotos().appendRow([b.crop, b.week || '', (b.date || ts.slice(0, 10)),
      file.getId(), name, b.note || '', b.loggedBy || 'web', ts]);
    return { ok: true, fileId: file.getId() };
  } finally { lock.releaseLock(); }
}

/** Return a photo's bytes as base64 so the page can show it (files stay private). */
function getPhoto(id) {
  var blob = DriveApp.getFileById(id).getBlob();
  return { ok: true, mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

/* ---------- helpers ---------- */
function _rows(sheet) {
  var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
  var h = data[0], out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var o = {}; for (var c = 0; c < h.length; c++) o[h[c]] = data[i][c];
    out.push(o);
  }
  return out;
}
function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
