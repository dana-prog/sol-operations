const _invoiceProps = {
  sheetName: 'Invoices',
  name: 'invoice',
};

addTypeProps(_invoiceProps);

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

function syncInvoices() {
  const lastSyncedDocNumber = _getLastSyncedDocNumber();
  const res = listNewSumitDocs(lastSyncedDocNumber);
  const docs = res['Data']['Documents'];
  if (!docs.length) {
    SOLLibrary.alert('No new invoices', `There are no available new summit docs (lastSyncedDocNumber = ${lastSyncedDocNumber}).`);
    return;
  }

  const sheetName = _invoiceProps.sheetName;
  const rows = SOLLibrary.generateSheetRows(sheetName, docs);
  SOLLibrary.appendRowsToSheet(sheetName, rows);
  SOLLibrary.sortSheet(sheetName, 'DocumentNumber', false);
  SOLLibrary.alert('Invoices synced', `${docs.length} invoices were synced from Sumit:\n\n` + docs.map(doc => doc['DocumentNumber']).join('\n'));
}

function _getLastSyncedDocNumber() {
  const docNumbers = SOLLibrary.getColumnValues(_getInvoicesSheet(), "DocumentNumber", false);
  return docNumbers.length ? Math.max(...docNumbers) : 0;
}

function _getInvoicesSheet() {
  const sheet = SOLLibrary.getSheet(_invoiceProps.sheetName, true);

  // header
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(_HEADER_ROW);
  }
  return sheet;
}
