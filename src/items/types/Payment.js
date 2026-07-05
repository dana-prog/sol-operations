
/**
 * Scans the Payments sheet and creates a new receipt for each payment row that does not have an receipt.
 */
function generateReceiptsForPayments() {
  const unprocessedPayments = _getUnprocessedPayments();
  let results = []; // map of paymentId -> receiptId

  for (const payment of unprocessedPayments) {
    const bankTransactionId = payment['bankTransactionId'];
    const date = Utilities.formatDate(payment['date'], Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    const name = payment['nameOnReceipt'] || payment['customerName'];
    const event = payment['eventName'];
    const amount = payment['amount'];

    try {
      const receipt = createSummitReceipt(bankTransactionId, date, name, event, amount);
      results.push({paymentId: payment['id'], receiptId: receipt['DocumentNumber']});
    } catch (e) {
      const msg = `${buildReceiptResultsMessage(results)}\n\n`;
      const err = `Receipt creation for payment '${payment['id']}' failed:\n${e.message || e}`;

      SOLLibrary.log('PaymentsSheet', 'createReceipts', msg);
      SOLLibrary.log('PaymentsSheet', 'createReceipts', err, SOLLibrary.LOG_LEVEL.ERROR);
      SOLLibrary.alert('Receipts partially created', `${msg}\n\n${err}`);
      return;
    }
  }

  const msg = buildReceiptResultsMessage(results);
  SOLLibrary.alert('Receipts created', msg);
}

function _getPaymentsSheet() {
  return SOLLibrary.getSheet(RECEIPTS_SHEET_NAME, false);
}

/**
 * Returns unprocessed bank-transfer payments sorted oldest-first, ready for
 * receipt creation. Reads all data in memory — does not modify the sheet.
 *
 * An payment is "unprocessed" when its receipt id cell is not a number
 * (empty or a marker like "<Cash>"). Cash payments are skipped.
 * Processing stops at the first bank transfer without a bank transaction id
 * to preserve date-sequential receipt ordering.
 *
 * @returns {Object[]} Payments to receipt, oldest first.
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
    // already has an receipt — skip
    if (typeof payment['receiptId'] === 'number') continue;

    // only bank transfers get receipts
    if (payment['paymentMethod'] !== 'Bank Transfer') continue;

    if (!payment['bankTransactionId']) {
      const msg = `Payment '${payment['id']}' has no related Bank Transaction Id - no receipt will be created for this and later payments.`;
      SOLLibrary.log('Sumit', 'getUnprocessedPayments', msg);
      SOLLibrary.alert('No Bank Transaction ID', msg);
      break;
    }

    res.push(payment);
  }

  return res;
}

function buildReceiptResultsMessage(results) {
  return `${results.length} receipts were created:\n\n` +
    results.map(res => `Receipt '${res.receiptId}' for Payment '${res.paymentId}'`).join('\n');
}
