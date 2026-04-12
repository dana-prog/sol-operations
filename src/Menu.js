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
  const expected = getTypeProps('income').sheetName;
  if (sheet.getName() !== expected) {
    SOLLibrary.alert('Not supported sheet', `Creating invoices is only supported in the '${expected}' sheet.`);
    return;
  }

  createInvoices();
  syncInvoices();
}

function _onSyncInvoicesItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const expected = getTypeProps('invoice').sheetName;
  if (sheet.getName() !== expected) {
    SOLLibrary.alert('Not supported sheet', `Syncing invoices is only supported in the '${expected}' sheet.`);
    return;
  }

  syncInvoices();
}

function _onToggleWriteToLogFileClicked() {
  SOLLibrary.toggleWriteToLogFile();
}