/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  createMenu();
}

// noinspection JSUnusedGlobalSymbols
function onFormSubmit(e) {
  const rowNum = e.range.getRow();
  onCustomerResponseSheetInsertRow(rowNum);
}