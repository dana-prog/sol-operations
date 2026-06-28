// Creates the SOL menu.
function createMenu() {
  SpreadsheetApp
    .getUi()
    .createMenu('SOL')
    .addItem('Create Invoices', '_onCreateInvoicesItemClicked')
    .addItem('Sync Invoices', '_onSyncInvoicesItemClicked')
    // .addSeparator()
    // .addItem('Toggle Write To Log File', '_onToggleWriteToLogFileClicked')
    .addToUi();
}

function _onCreateInvoicesItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const expected = _paymentProps.sheetName;

  createInvoices();
  syncInvoices();
}

function _onSyncInvoicesItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const expected = _invoiceProps.sheetName;
  if (sheet.getName() !== expected) {
    SOLLibrary.alert('Not supported sheet', `Syncing invoices is only supported in the '${expected}' sheet.`);
    return;
  }

  syncInvoices();
}

function _onToggleWriteToLogFileClicked() {
  SOLLibrary.toggleWriteToLogFile();
}