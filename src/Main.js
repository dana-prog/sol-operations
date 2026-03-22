/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  createMenu();
}

/**
 * An event handler for a new row in a sheet.
 *
 * @param sheet the sheet where the new row was inserted.
 * @param newRowNum the row number of the inserted row (1-based index).
 */
function onNewRow(sheet, newRowNum) {
  if (isIncomesSheet(sheet)) {
    onIncomesSheetInsertRow(newRowNum);
  }
}

