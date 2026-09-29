/**
 * Kholl'app — serveur d'inscriptions (Google Apps Script)
 * À coller dans : Google Sheet > Extensions > Apps Script
 * Onglets créés automatiquement : "Inscriptions" et "Journal".
 */

const SHEET_BOOKINGS = 'Inscriptions';
const SHEET_LOG = 'Journal';
const HEADERS = ['Horodatage', 'Année', 'Semaine', 'Date', 'Horaire', 'Examinateur', 'Salle', 'Étudiant', 'ID créneau'];

function doGet(e) {
  try {
    const year = String((e.parameter && e.parameter.year) || '');
    const rows = readBookings_().filter(function (b) { return !year || b.year === year; });
    return json_({ ok: true, bookings: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'Serveur occupé, réessayez.' });
  try {
    const p = JSON.parse(e.postData.contents);
    const sh = sheet_(SHEET_BOOKINGS, HEADERS);
    const rows = readBookings_();

    if (p.action === 'book') {
      if (!p.slotId || !p.student || !p.year || !p.week) return json_({ ok: false, error: 'Requête incomplète.' });
      if (rows.some(function (b) { return b.slotId === p.slotId; }))
        return json_({ ok: false, error: "Ce créneau vient d'être réservé par quelqu'un d'autre." });
      if (rows.some(function (b) { return b.year === String(p.year) && b.week === p.week && b.student === p.student; }))
        return json_({ ok: false, error: 'Vous êtes déjà inscrit(e) sur un créneau cette semaine.' });
      sh.appendRow([new Date(), String(p.year), "'" + p.week, "'" + p.date, p.time, p.examiner, p.room, p.student, p.slotId]);
      log_('Inscription', p);
      return json_({ ok: true });
    }

    if (p.action === 'cancel') {
      const data = sh.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][8]) === p.slotId && String(data[i][7]) === p.student) {
          sh.deleteRow(i + 1);
          log_('Désinscription', p);
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: 'Inscription introuvable.' });
    }

    return json_({ ok: false, error: 'Action inconnue.' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function readBookings_() {
  const sh = sheet_(SHEET_BOOKINGS, HEADERS);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[8]) continue;
    out.push({ year: String(r[1]), week: fmtDate_(r[2]), student: String(r[7]), slotId: String(r[8]) });
  }
  return out;
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function log_(what, p) {
  const sh = sheet_(SHEET_LOG, ['Horodatage', 'Action', 'Année', 'Date', 'Horaire', 'Salle', 'Étudiant']);
  sh.appendRow([new Date(), what, String(p.year), "'" + (p.date || ''), p.time || '', p.room || '', p.student]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
