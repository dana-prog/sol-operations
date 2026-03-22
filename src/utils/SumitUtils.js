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

function createInvoice(bankTransferId, date, name, productName, amount) {
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
      Type: 1, // InvoiceAndReceipt
      Description: bankTransferId
    },
    Items: [
      {
        DocumentCurrency_TotalPrice: amount,
        Item: {
          Name: productName,
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

  const response = post(SUMIT_CREATE_DOCUMENT_API, payload);
  if (response['Status'] !== 0) {
    throw new Error(`Failed to create invoice: ${response['UserErrorMessage'] || response['TechnicalErrorDetails']}`);
  }

  logArgs('Sumit', 'createInvoice', {data: response['Data']});
  return response['Data'];
}

function listNewSumitDocs(lastSyncedDocNumber) {
  const payload = {
    Credentials: getSumitCredentials(),
    IncludeDrafts: SUMIT_DRAFT_MODE,
    DocumentTypes: ['InvoiceAndReceipt'],
    DocumentNumberFrom: lastSyncedDocNumber + 1,
    // TODO: implement real paging (or at least an error when page size is not enough)
    Paging: {PageSize: 1000},
  };

  return post(SUMIT_LIST_DOCUMENTS_API, payload);
}