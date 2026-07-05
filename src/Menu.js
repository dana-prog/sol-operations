// Creates the SOL menu.
function createMenu() {
  SpreadsheetApp
    .getUi()
    .createMenu('SOL')
    .addItem('Create Receipts', '_onCreateReceiptsItemClicked')
    .addItem('Sync Receipts', '_onSyncReceiptsItemClicked')
    // .addSeparator()
    // .addItem('Toggle Write To Log File', '_onToggleWriteToLogFileClicked')
    .addToUi();
}

function _onCreateReceiptsItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const expected = RECEIPTS_SHEET_NAME;

  createReceipts();
  syncReceipts();
}

function _onSyncReceiptsItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const expected = RECEIPTS_SHEET_NAME;
  if (sheet.getName() !== expected) {
    SOLLibrary.alert('Not supported sheet', `Syncing receipts is only supported in the '${expected}' sheet.`);
    return;
  }

  syncReceipts();
}

function _onToggleWriteToLogFileClicked() {
  SOLLibrary.toggleWriteToLogFile();
}