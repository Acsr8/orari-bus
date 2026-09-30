/* ============================================================
   API Statistiche Bus — Google Apps Script
   Gestisce: creazione (create), lettura elenco (list),
             aggiornamento (update), eliminazione (delete)
   ============================================================
   Colonne del foglio "Foglio1":
   A: Linea | B: Direzione | C: Fermata Partenza | D: Fermata Arrivo
   E: Inizio Tratta | F: Fine Tratta | G: Minuti | H: ID (nascosto)
   ============================================================ */

const SHEET_NAME = "Foglio1";

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "list") {
    return listRecords();
  }
  return jsonResponse({ result: "error", error: "Azione GET non riconosciuta." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "create";

    if (action === "create") {
      return createRecord(data);
    } else if (action === "update") {
      return updateRecord(data);
    } else if (action === "delete") {
      return deleteRecord(data);
    } else {
      return jsonResponse({ result: "error", error: "Azione non riconosciuta: " + action });
    }
  } catch (err) {
    return jsonResponse({ result: "error", error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function createRecord(data) {
  const sheet = getSheet();
  const nextRow = sheet.getLastRow() + 1;
  const id = Utilities.getUuid();

  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    data.linea,
    data.direzione,
    data.partenza,
    data.arrivo,
    data.inizio,
    data.fine,
    data.minuti,
    id
  ]]);

  return jsonResponse({ result: "success", row: nextRow, id: id });
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      return i + 2; // +2 perché getRange parte da riga 2 e gli indici sono 0-based
    }
  }
  return -1;
}

function updateRecord(data) {
  const sheet = getSheet();
  const row = findRowById(sheet, data.id);
  if (row === -1) {
    return jsonResponse({ result: "error", error: "Rilevazione non trovata (id: " + data.id + ")." });
  }

  sheet.getRange(row, 1, 1, 7).setValues([[
    data.linea,
    data.direzione,
    data.partenza,
    data.arrivo,
    data.inizio,
    data.fine,
    data.minuti
  ]]);

  return jsonResponse({ result: "success", row: row });
}

function deleteRecord(data) {
  const sheet = getSheet();
  const row = findRowById(sheet, data.id);
  if (row === -1) {
    return jsonResponse({ result: "error", error: "Rilevazione non trovata (id: " + data.id + ")." });
  }

  sheet.deleteRow(row);
  return jsonResponse({ result: "success" });
}

function listRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ result: "success", records: [] });
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const records = values
    .filter(row => row[7]) // scarta righe senza ID (es. righe vuote residue)
    .map(row => ({
      linea: row[0],
      direzione: row[1],
      partenza: row[2],
      arrivo: row[3],
      inizio: formatCellValue(row[4]),
      fine: formatCellValue(row[5]),
      minuti: row[6],
      id: row[7]
    }));

  return jsonResponse({ result: "success", records: records });
}

function formatCellValue(value) {
  // Se Google Sheets ha convertito la stringa in un oggetto Date, riformatta come testo
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }
  return value;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
