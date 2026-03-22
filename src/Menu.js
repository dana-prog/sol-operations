// Creates the SOL menu.
function createMenu() {
  SpreadsheetApp
    .getUi()
    .createMenu('SOL')
    .addItem('Create Invoices', '_onCreateInvoicesItemClicked')
    .addItem('Sync Invoices', '_onSyncInvoicesItemClicked')
    .addSeparator()
    .addItem('Toggle Log Alerts', '_onToggleLogAlertsItemClicked')
    .addToUi();
}

function _onCreateInvoicesItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== INCOMES_SHEET_NAME) {
    SOLLibrary.alert('Not supported sheet', `Creating invoices is only supported in the '${INCOMES_SHEET_NAME}' sheet.`);
    return;
  }

  createInvoices();
  syncInvoices();
}

function _onSyncInvoicesItemClicked() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== INVOICES_SHEET_NAME) {
    SOLLibrary.alert('Not supported sheet', `Syncing invoices is only supported in the '${INVOICES_SHEET_NAME}' sheet.`);
  }
}

function _onToggleLogAlertsItemClicked() {
  SOLLibrary.toggleAlertLogs();
}