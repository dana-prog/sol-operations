const _incomeProps = {
  sheetName: 'Incomes',
  name: 'income',
};

/**
 * Called when a new row is inserted to the Incomes sheet.
 * Assigns the next available id to the inserted row
 * @param rowNum the row number of the inserted row (1-based index)
 */
function onIncomesSheetInsertRow(rowNum){
  const idColNum = SOLLibrary.getColNumByHeader(_getIncomesSheet(), "id");
  const nextId = _getNextId();

  _getIncomesSheet().getRange(rowNum, idColNum).setValue(nextId);
}

/**
 * Scans the Invoices sheet and creates a new invoice for each income row which does not have an invoice.
 */
function createInvoices() {
  const unprocessedIncomes = _getUnprocessedIncomes();
  let results = []; // map of incomeId -> invoiceId

  for (const income of unprocessedIncomes) {
    const bankTransactionId = income['bankTransactionId'];
    const date = Utilities.formatDate(income['date'], Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    const name = income['nameOnInvoice'] || income['customerName'];
    const product = income['productName'];
    const amount = income['amount'];

    try {
      const invoice = createInvoice(bankTransactionId, date, name, product, amount);
      results.push({incomeId: income['id'], invoiceId: invoice['DocumentNumber']});
    } catch (e) {
      const msg = `${buildInvoiceResultsMessage(results)}\n\n`;
      const err = `Invoice creation for income '${income['id']}' failed:\n${e.message || e}`;

      SOLLibrary.log('IncomesSheet', 'createInvoices', msg);
      SOLLibrary.log('IncomesSheet', 'createInvoices', err, SOLLibrary.LOG_LEVEL.ERROR);
      SOLLibrary.alert('Invoices partially created', `${msg}\n\n${err}`);
      return;
    }
  }

  const msg = buildInvoiceResultsMessage(results);
  SOLLibrary.alert('Invoices created', msg);
}

function isIncomesSheet(sheet) {
  return sheet.getSheetName() === _incomeProps.sheetName;
}

function _getIncomesSheet() {
  return SOLLibrary.getSheet(_incomeProps.sheetName, false);
}

/**
 * Returns an array of objects for all the rows which do not have an associated
 * @returns {*[]}
 * @private
 */
function _getUnprocessedIncomes() {
  const sheetName = _incomeProps.sheetName;
  const sheet = SOLLibrary.getSheet(sheetName);
  SOLLibrary.sortSheet(sheetName, 'date');
  const headerMap = SOLLibrary.getHeaderMap(sheet);
  const invoiceIds = SOLLibrary.getColumnValues(sheetName, 'invoice id', false);
  // find the index of 1st row with an invoice id - all the rows before it are not processed yet
  const unprocessedRowsCount = invoiceIds.findIndex(value => typeof value === 'number');

  // get all the rows from the recent income (row 2) until the first income that was not processed
  const rows = sheet.getRange(2, 1, unprocessedRowsCount, sheet.getLastColumn()).getValues();

  const res = [];

  // iterate unprocessed rows
  for (let rowIndex = unprocessedRowsCount - 1; rowIndex >= 0; rowIndex--) {
    const row = rows[rowIndex];
    const obj = SOLLibrary.arrayToObj(row, headerMap);

    // generate invoice only for bank transfers
    if (obj['paymentMethod'] !== 'Bank Transfer') {
      continue;
    }

    if (!obj['bankTransactionId']) {
      // the current income was not matched with a bank transaction id and therefore it is not possible to generate an invoice for it - stop processing
      const msg = `Income '${obj['id']}' has no related Bank Transaction Id - no invoice will be created for this and later incomes.`;
      SOLLibrary.log('Sumit', 'getUnprocessedIncomes', msg);
      SOLLibrary.alert('No Bank Transaction ID', msg);
      break;
    }

    // add the current income row to the result
    res.push(obj);
  }

  return res;
}

function buildInvoiceResultsMessage(results) {
  return `${results.length} invoices were created:\n\n` +
    results.map(res => `Invoice '${res.invoiceId}' for Income '${res.incomeId}'`).join('\n');
}

function _getNextId() {
  const ids = SOLLibrary.getColumnValues(_incomeProps.sheetName, "id", false);
  return ids.length ? Math.max(...ids) + 1 : 1;
}