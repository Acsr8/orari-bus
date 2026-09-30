/* ============================================================
   API Statistiche Bus - Google Apps Script
   Colonne del primo foglio:
   A: ID | B: Linea | C: Fermata Partenza | D: Fermata Arrivo
   E: Ora Partenza | F: Ora Arrivo | G: Minuti
   Le date sono testo nel formato AAAA-MM-GG HH:MM.
   I minuti sono calcolati qui: Ora Arrivo - Ora Partenza.
   ============================================================ */

const SECRET = "BUS";
const COLS = 7;

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet(e) {
  if (e.parameter.key !== SECRET) return json({ result: "error", error: "Chiave non valida." });
  if (e.parameter.action === "list") return listRecords();
  return json({ result: "error", error: "Azione non riconosciuta." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.key !== SECRET) return json({ result: "error", error: "Chiave non valida." });
    if (data.action === "create") return createRecord(data);
    if (data.action === "update") return updateRecord(data);
    if (data.action === "delete") return deleteRecord(data);
    return json({ result: "error", error: "Azione non riconosciuta." });
  } catch (err) {
    return json({ result: "error", error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// "2026-09-30 18:45" -> minuti dall'epoca, senza fuso orario (evita errori di ora legale)
function toMinutes(text) {
  const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(String(text).trim());
  if (!m) throw new Error("Formato data/ora non valido: " + text);
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) / 60000;
}

function buildRow(id, d) {
  const minuti = toMinutes(d.arrivo_ora) - toMinutes(d.partenza_ora);
  if (minuti < 0) throw new Error("L'arrivo è prima della partenza.");
  if (!d.linea || !d.partenza || !d.arrivo) throw new Error("Campi obbligatori mancanti.");
  return [id, d.linea, d.partenza, d.arrivo, d.partenza_ora, d.arrivo_ora, minuti];
}

function writeRow(sheet, row, values) {
  sheet.getRange(row, 5, 1, 2).setNumberFormat("@"); // testo: Sheets non converte in data
  sheet.getRange(row, 1, 1, COLS).setValues([values]);
}

function findRowById(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function createRecord(d) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const values = buildRow(id, d);
  writeRow(sheet, sheet.getLastRow() + 1, values);
  return json({ result: "success", id: id });
}

function updateRecord(d) {
  const sheet = getSheet();
  const row = findRowById(sheet, d.id);
  if (row === -1) return json({ result: "error", error: "Rilevazione non trovata." });
  writeRow(sheet, row, buildRow(d.id, d));
  return json({ result: "success" });
}

function deleteRecord(d) {
  const sheet = getSheet();
  const row = findRowById(sheet, d.id);
  if (row === -1) return json({ result: "error", error: "Rilevazione non trovata." });
  sheet.deleteRow(row);
  return json({ result: "success" });
}

function listRecords() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) return json({ result: "success", records: [] });
  const records = sheet.getRange(2, 1, last - 1, COLS).getValues()
    .filter(r => r[0])
    .map(r => ({
      id: r[0], linea: r[1], partenza: r[2], arrivo: r[3],
      partenza_ora: cellText(r[4]), arrivo_ora: cellText(r[5]), minuti: r[6]
    }));
  return json({ result: "success", records: records });
}

function cellText(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  return v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
