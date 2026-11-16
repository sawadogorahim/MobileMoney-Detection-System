/**
 * ============================================================
 *  Orange Money Sync — Google Apps Script Web App
 *  Déployer comme : Exécuter en tant que "Moi"
 *                   Accès : "Tout le monde"
 * ============================================================
 *
 *  Remplacer SHEET_ID par l'ID de votre Google Spreadsheet.
 *  Ex : https://docs.google.com/spreadsheets/d/SHEET_ID/edit
 */

const SHEET_ID   = "VOTRE_SPREADSHEET_ID_ICI";
const SHEET_NAME = "Transactions";

// En-têtes de la feuille (créés automatiquement si absents)
const HEADERS = [
  "Date", "Type", "Sens", "Montant (FCFA)", "Numéro",
  "Nom / Agent", "Solde après (FCFA)", "ID Transaction", "Reçu le"
];

// ── Point d'entrée POST ───────────────────────────────────────────────────────

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || "append";

    if (action === "batch") {
      const transactions = body.transactions || [];
      transactions.forEach(t => appendRow(getSheet(), t));
      return jsonOk("Batch de " + transactions.length + " transaction(s) enregistré");
    }

    if (action === "append") {
      appendRow(getSheet(), body.transaction || body);
      return jsonOk("Transaction enregistrée");
    }

    return jsonError("Action inconnue : " + action);

  } catch (err) {
    Logger.log("Erreur doPost : " + err.message);
    return jsonError(err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Créer les en-têtes si la feuille est vide
  if (sheet.getLastRow() === 0) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#FF6600");
    headerRange.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);   // Date
    sheet.setColumnWidth(4, 130);   // Montant
    sheet.setColumnWidth(6, 200);   // Nom
    sheet.setColumnWidth(8, 200);   // ID Transaction
  }

  return sheet;
}

function appendRow(sheet, t) {
  // Vérifier les doublons par ID de transaction
  const transId = t.transId || "";
  if (transId && isDuplicate(sheet, transId)) {
    Logger.log("Doublon ignoré : " + transId);
    return;
  }

  const row = [
    t.date     || "",
    t.type     || "",
    t.sens     || "",
    t.montant  || 0,
    t.numero   || "",
    t.nom      || "",
    t.solde    || 0,
    transId,
    new Date().toLocaleString("fr-FR")
  ];

  sheet.appendRow(row);

  // Colorer la ligne selon le sens
  const lastRow   = sheet.getLastRow();
  const rowRange  = sheet.getRange(lastRow, 1, 1, HEADERS.length);
  const sens      = (t.sens || "").toUpperCase();

  if (sens === "CRÉDIT" || sens === "CREDIT") {
    rowRange.setBackground("#E8F5E9");   // Vert clair
  } else if (sens === "DÉBIT" || sens === "DEBIT") {
    rowRange.setBackground("#FFEBEE");   // Rouge clair
  }
}

function isDuplicate(sheet, transId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  // Colonne 8 = ID Transaction
  const ids = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  return ids.some(row => row[0] === transId);
}

function jsonOk(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "error", message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Test manuel (à exécuter depuis l'éditeur Apps Script) ────────────────────
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: "append",
        transaction: {
          transId:  "PP260625.1702.TEST001",
          date:     "2026-06-25",
          type:     "Reçu",
          sens:     "CRÉDIT",
          montant:  14900,
          numero:   "66385699",
          nom:      "AMOS",
          solde:    15027.13
        }
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

