/**
 * Web app entry point. Serves sheet data as JSON.
 *
 * Usage:
 *   GET <deployUrl>?sheet=Customers
 *   GET <deployUrl>?sheet=Customers&limit=10
 *   GET <deployUrl>                         → lists available sheets
 *
 * Deploy: Publish → Deploy as web app → Execute as "Me", access "Anyone".
 */
function doGet(e) {
  const sheetName = (e.parameter || {}).sheet;
  const limit = parseInt((e.parameter || {}).limit, 10) || 0;

  if (!sheetName) {
    return _jsonResponse(_listSheets());
  }

  return _jsonResponse(_getSheetData(sheetName, limit));
}

function _listSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return {
    sheets: sheets.map(s => ({
      name: s.getName(),
      rows: s.getLastRow(),
      cols: s.getLastColumn(),
    })),
  };
}

function _getSheetData(sheetName, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { error: `Sheet '${sheetName}' not found` };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheet: sheetName, headers: [], rows: [] };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => h.toString());

  const dataRowCount = lastRow - 1;
  if (dataRowCount < 1) {
    return { sheet: sheetName, headers, rows: [] };
  }

  const rowCount = limit > 0 ? Math.min(limit, dataRowCount) : dataRowCount;
  const values = sheet.getRange(2, 1, rowCount, lastCol).getValues();

  const rows = values.map(row => {
    const obj = {};
    for (let col = 0; col < headers.length; col++) {
      const val = row[col];
      obj[headers[col]] = val instanceof Date
        ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : val;
    }
    return obj;
  });

  return { sheet: sheetName, headers, rowCount: rows.length, rows };
}

function _jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
