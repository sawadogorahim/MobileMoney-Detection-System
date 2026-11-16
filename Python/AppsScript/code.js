function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Transactions");
    if (!sheet) {
      sheet = ss.insertSheet("Transactions");
      sheet.appendRow(["Date","Heure","Type","Nom","Téléphone","Montant","Solde","Opérateur","TID"]);
    }

    const action = e.parameter.action || "sync";
    const rawData = e.parameter.batchData || "[]";

    if (action == "vider") {
      sheet.clearContents();
      sheet.appendRow(["Date","Heure","Type","Nom","Téléphone","Montant","Solde","Opérateur","TID"]);
      return ContentService.createTextOutput("OK - Feuille vidée.")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    let data = [];
    try {
      data = JSON.parse(rawData);
    } catch(jsonErr) {
      return ContentService.createTextOutput("ERREUR JSON: " + jsonErr.toString())
        .setMimeType(ContentService.MimeType.TEXT);
    }

    if (action == "clear") {
      sheet.clearContents();
      sheet.appendRow(["Date","Heure","Type","Nom","Téléphone","Montant","Solde","Opérateur","TID"]);
    }

    const tids = {};
    if (sheet.getLastRow() > 1) {
      const values = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues();
      values.forEach(function(r) { if(r[0]) tids[r[0]] = true; });
    }

    let ajoutees = 0;
    data.forEach(function(t) {
      if (!t.tid) return;
      if (tids[t.tid]) return;
      sheet.appendRow([t.date, t.heure, t.type, t.nom, t.telephone, t.montant, t.solde, t.operateur, t.tid]);
      tids[t.tid] = true;
      ajoutees++;
    });

    // Trier par date (colonne A) et heure (colonne B) du plus récent au plus ancien
    if (sheet.getLastRow() > 2) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
      range.sort([
        {column: 1, ascending: false},
        {column: 2, ascending: false}
      ]);
    }

    return ContentService.createTextOutput("OK - " + ajoutees + " transaction(s) ajoutée(s).")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput("ERREUR: " + err.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}