PAYMENTS_SHEET_NAME = 'Payments';

/**
 * Called when a new row is inserted to the Payments sheet.
 * Assigns the next available id to the inserted row
 * @param rowNum the row number of the inserted row (1-based index)
 */
function onPaymentsSheetInsertRow(rowNum){
  const idColNum = SOLLibrary.getColNumByHeader(_getPaymentsSheet(), "id");
  const nextId = _getNextId();

  _getPaymentsSheet().getRange(rowNum, idColNum).setValue(nextId);
}

/**
 * Scans the Invoices sheet and creates a new invoice for each payment row which does not have an invoice.
 */
function createInvoices() {
  const unprocessedPayments = _getUnprocessedPayments();
  let results = []; // map of paymentId -> invoiceId

  for (const payment of unprocessedPayments) {
    const bankTransactionId = payment['bankTransactionId'];
    const date = Utilities.formatDate(payment['date'], Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    const name = payment['nameOnInvoice'] || payment['customerName'];
    const event = payment['eventName'];
    const amount = payment['amount'];

    try {
      const invoice = createInvoice(bankTransactionId, date, name, event, amount);
      results.push({paymentId: payment['id'], invoiceId: invoice['DocumentNumber']});
    } catch (e) {
      const msg = `${buildInvoiceResultsMessage(results)}\n\n`;
      const err = `Invoice creation for payment '${payment['id']}' failed:\n${e.message || e}`;

      SOLLibrary.log('PaymentsSheet', 'createInvoices', msg);
      SOLLibrary.log('PaymentsSheet', 'createInvoices', err, SOLLibrary.LOG_LEVEL.ERROR);
      SOLLibrary.alert('Invoices partially created', `${msg}\n\n${err}`);
      return;
    }
  }

  const msg = buildInvoiceResultsMessage(results);
  SOLLibrary.alert('Invoices created', msg);
}

function isPaymentsSheet(sheet) {
  return sheet.getSheetName() === INVOICES_SHEET_NAME;
}

function _getPaymentsSheet() {
  return SOLLibrary.getSheet(INVOICES_SHEET_NAME, false);
}

/**
 * Returns unprocessed bank-transfer payments sorted oldest-first, ready for
 * invoice creation. Reads all data in memory — does not modify the sheet.
 *
 * An payment is "unprocessed" when its invoice id cell is not a number
 * (empty or a marker like "<Cash>"). Cash payments are skipped.
 * Processing stops at the first bank transfer without a bank transaction id
 * to preserve date-sequential receipt ordering.
 *
 * @returns {Object[]} Payments to invoice, oldest first.
 * @private
 */
function _getUnprocessedPayments() {
  const sheet = _getPaymentsSheet();
  const headerMap = SOLLibrary.getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const allRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const allPayments = allRows.map(row => SOLLibrary.arrayToObj(row, headerMap));

  // sort by id ascending (oldest first) for sequential processing
  allPayments.sort((a, b) => {
    const idA = typeof a['id'] === 'number' ? a['id'] : 0;
    const idB = typeof b['id'] === 'number' ? b['id'] : 0;
    return idA - idB;
  });

  const res = [];

  for (const payment of allPayments) {
    // already has an invoice — skip
    if (typeof payment['invoiceId'] === 'number') continue;

    // only bank transfers get invoices
    if (payment['paymentMethod'] !== 'Bank Transfer') continue;

    if (!payment['bankTransactionId']) {
      const msg = `Payment '${payment['id']}' has no related Bank Transaction Id - no invoice will be created for this and later payments.`;
      SOLLibrary.log('Sumit', 'getUnprocessedPayments', msg);
      SOLLibrary.alert('No Bank Transaction ID', msg);
      break;
    }

    res.push(payment);
  }

  return res;
}

function buildInvoiceResultsMessage(results) {
  return `${results.length} invoices were created:\n\n` +
    results.map(res => `Invoice '${res.invoiceId}' for Payment '${res.paymentId}'`).join('\n');
}

function _getNextId() {
  const ids = SOLLibrary.getColumnValues(INVOICES_SHEET_NAME, "id", false);
  return ids.length ? Math.max(...ids) + 1 : 1;
}