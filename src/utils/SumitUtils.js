const SUMIT_LIST_DOCUMENTS_API = "https://api.sumit.co.il/accounting/documents/list/";
const SUMIT_CREATE_DOCUMENT_API = "https://api.sumit.co.il/accounting/documents/create/";
const SUMIT_DRAFT_MODE = false;

function getSumitCredentials() {
  const props = PropertiesService.getScriptProperties();
  return {
    CompanyID: props.getProperty('summitCompanyId'),
    APIKey: props.getProperty('summitPrivateApiKey'),
  };
}

function createSummitReceipt(bankTransferId, date, name, eventName, amount) {
  const payload = {
    Details: {
      IsDraft: SUMIT_DRAFT_MODE,
      Date: date,
      Customer: {
        Name: name,
        SearchMode: "Name"
      },
      Language: 1, // English
      Currency: 764, // THB
      Type: 1, // ReceiptAndReceipt
      Description: bankTransferId
    },
    Items: [
      {
        DocumentCurrency_TotalPrice: amount,
        Item: {
          Name: eventName,
          SearchMode: "Name"
        }
      }
    ],
    Payments: [
      {
        DocumentCurrency_Amount: amount,
        Type: 3 // Bank transfer
      }
    ],
    Credentials: getSumitCredentials()
  };

  const response = SOLLibrary.post(SUMIT_CREATE_DOCUMENT_API, payload);
  if (response['Status'] !== 0) {
    throw new Error(`Failed to create receipt: ${response['UserErrorMessage'] || response['TechnicalErrorDetails']}`);
  }

  SOLLibrary.logArgs('Sumit', 'createReceipt', {data: response['Data']});
  return response['Data'];
}

/**
 * Lists new receipts in Sumit accounting (all receipts since lastSyncedReceiptNumber)
 * @param lastSyncedReceiptNumber
 * @returns {any}
 */
function listNewSumitReceipts(lastSyncedReceiptNumber) {
  const payload = {
    Credentials: getSumitCredentials(),
    IncludeDrafts: true,
    DocumentTypes: ['ReceiptAndReceipt'],
    DocumentNumberFrom: lastSyncedReceiptNumber + 1,
    // TODO: implement real paging (or at least an error when page size is not enough)
    Paging: {PageSize: 1000},
  };

  return SOLLibrary.post(SUMIT_LIST_DOCUMENTS_API, payload);
}