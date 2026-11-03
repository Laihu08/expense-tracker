const SHEET_NAME = 'Sheet1';

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = JSON.parse(e.postData.contents);

    // ─── DELETE action ───
    if (data.action === "delete") {
      const row = data.row; // Row number in the sheet (1-indexed)
      if (row && row >= 2) { // Don't delete the header row
        sheet.deleteRow(row);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Row deleted" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid row" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ─── ADD action (default) ───
    sheet.appendRow([data.date, data.item, data.category, data.amount]);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // Skip the first row (headers) and organize the rest
    const rows = data.slice(1);
    const result = rows.map(row => {
      return {
        date: row[0],
        item: row[1],
        category: row[2],
        amount: row[3]
      };
    });

    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
