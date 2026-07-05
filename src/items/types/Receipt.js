const RECEIPTS_SHEET_NAME = 'Receipts';

const _HEADER_ROW = [
  'IsDraft',
  'Date',
  'CustomerID',
  'CustomerName',
  'Language',
  'Currency',
  'Type',
  'Description',
  'ExternalReference',
  'DueDate',
  'DocumentDownloadURL',
  'DocumentPaymentURL',
  'DocumentID',
  'DocumentNumber',
  'DocumentValue',
  'CompanyValue',
  'IsClosed',
];

/**
 * Syncs receipts from Sumit accounting. Creates new rows in the Receipts sheet for all receipts since the last sync.
 */
function syncReceipts() {
  const lastSyncedDocNumber = _getLastSyncedReceiptNumber();
  const res = listNewSumitReceipts(lastSyncedDocNumber);
  const docs = res['Data']['Documents'];
  if (!docs.length) {
    SOLLibrary.alert('No new receipts', `There are no available new summit docs (lastSyncedDocNumber = ${lastSyncedDocNumber}).`);
    return;
  }

  const sheet = _getReceiptsSheet();
  const rows = SOLLibrary.generateSheetRows(sheetName, docs);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  sheet.getRange(lastRow + 1, 1, rows.length, lastCol).setValues(rows);
  SOLLibrary.sortSheet(RECEIPTS_SHEET_NAME, 'DocumentNumber', false);
  SOLLibrary.alert('Receipts synced', `${docs.length} receipts were synced from Sumit:\n\n` + docs.map(doc => doc['DocumentNumber']).join('\n'));
}

function _getLastSyncedReceiptNumber() {
  const docNumbers = SOLLibrary.getColumnValues(RECEIPTS_SHEET_NAME, "DocumentNumber", false);
  return docNumbers.length ? Math.max(...docNumbers) : 0;
}

function _getReceiptsSheet() {
  const sheet = SOLLibrary.getSheet(RECEIPTS_SHEET_NAME, true);

  // header
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(_HEADER_ROW);
  }
  return sheet;
}
