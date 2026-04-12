/**
 * Web app entry point. Serves sheet data as JSON.
 *
 * Usage:
 *   GET <deployUrl>?sheet=Customers
 *   GET <deployUrl>?sheet=Customers&limit=10
 *   GET <deployUrl>?sheet=Customers&formulas=true  → formulas instead of values
 *   GET <deployUrl>                                → lists available sheets
 *
 * Deploy: Publish → Deploy as web app → Execute as "Me", access "Anyone".
 */
function doGet(e) {
  const params = e.parameter || {};
  const sheetName = params.sheet;
  const limit = parseInt(params.limit, 10) || 0;
  const formulas = params.formulas === 'true';

  if (!sheetName) {
    return _jsonResponse(_listSheets());
  }

  return _jsonResponse(_getSheetData(sheetName, limit, formulas));
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

function _getSheetData(sheetName, limit, formulas) {
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
  const range = sheet.getRange(2, 1, rowCount, lastCol);
  const values = range.getValues();
  const formulaGrid = formulas ? range.getFormulas() : null;

  const rows = values.map((row, rowIdx) => {
    const obj = {};
    for (let col = 0; col < headers.length; col++) {
      const val = row[col];
      const formula = formulaGrid ? formulaGrid[rowIdx][col] : '';
      if (formula) {
        obj[headers[col]] = { value: val instanceof Date
          ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : val, formula };
      } else {
        obj[headers[col]] = val instanceof Date
          ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : val;
      }
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
