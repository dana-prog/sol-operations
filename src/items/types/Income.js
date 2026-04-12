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
 * Returns unprocessed bank-transfer incomes sorted oldest-first, ready for
 * invoice creation. Reads all data in memory — does not modify the sheet.
 *
 * An income is "unprocessed" when its invoice id cell is not a number
 * (empty or a marker like "<Cash>"). Cash payments are skipped.
 * Processing stops at the first bank transfer without a bank transaction id
 * to preserve date-sequential receipt ordering.
 *
 * @returns {Object[]} Incomes to invoice, oldest first.
 * @private
 */
function _getUnprocessedIncomes() {
  const sheet = _getIncomesSheet();
  const headerMap = SOLLibrary.getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const allRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const allIncomes = allRows.map(row => SOLLibrary.arrayToObj(row, headerMap));

  // sort by date ascending (oldest first) for sequential processing
  allIncomes.sort((a, b) => {
    const dateA = a['date'] instanceof Date ? a['date'].getTime() : 0;
    const dateB = b['date'] instanceof Date ? b['date'].getTime() : 0;
    return dateA - dateB;
  });

  const res = [];

  for (const income of allIncomes) {
    // already has an invoice — skip
    if (typeof income['invoiceId'] === 'number') continue;

    // only bank transfers get invoices
    if (income['paymentMethod'] !== 'Bank Transfer') continue;

    if (!income['bankTransactionId']) {
      const msg = `Income '${income['id']}' has no related Bank Transaction Id - no invoice will be created for this and later incomes.`;
      SOLLibrary.log('Sumit', 'getUnprocessedIncomes', msg);
      SOLLibrary.alert('No Bank Transaction ID', msg);
      break;
    }

    res.push(income);
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