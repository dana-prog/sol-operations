const CUSTOMER_RESPONSES_SHEET_NAME = 'Customer Form Responses';
const CUSTOMERS_SHEET_NAME = 'Customers';

function onCustomerResponseSheetInsertRow(rowNum) {
  const customersSheet = SpreadsheetApp.getActive().getSheetByName(CUSTOMERS_SHEET_NAME);

  const responseValues = SOLLibrary.getRowValues(CUSTOMER_RESPONSES_SHEET_NAME, rowNum);
  const customerValues = {};

  customerValues["id"] = Utilities.getUuid();
  customerValues["first name"] = SOLLibrary.capitalize(responseValues["First Name"]);
  customerValues["middle name"] = SOLLibrary.capitalize(responseValues["Middle Name"] || "");
  customerValues["last name"] = SOLLibrary.capitalize(responseValues["Last Name"]);
  customerValues["nickname"] = SOLLibrary.capitalize(responseValues["Nickname"] || "");
  customerValues["whatsapp"] = _digitsOnly(responseValues["WhatsApp Number"]);
  customerValues["phone"] = _digitsOnly(responseValues["Phone Number"]);
  customerValues["email"] = responseValues["Email"];
  customerValues["nationality"] = responseValues["Nationality"];
  customerValues["birth date"] = responseValues["Birth Date"];
  customerValues["allergies & dietary preferences"] = responseValues["Allergies & Dietary Preferences"];
  customerValues["emergency contact name"] = SOLLibrary.capitalize(responseValues["Emergency Contact Full Name"]);
  customerValues["emergency contact whatsapp"] = _digitsOnly(
    responseValues["Emergency Contact WhatsApp"]);
  customerValues["emergency contact relation"] = responseValues["Emergency Contact Relation"];

  const newCustomerRow = SOLLibrary.generateSheetRows(CUSTOMERS_SHEET_NAME, [customerValues])[0];
  customersSheet.appendRow(newCustomerRow);

  // copy full name formula to the new row
  const fullNameColNumber = SOLLibrary.getColNumByHeader(customersSheet, "full name");
  const sourceCell = customersSheet.getRange(2, fullNameColNumber);
  const targetCell = customersSheet.getRange(customersSheet.getLastRow(), fullNameColNumber);
  sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // copy format to the new row
  const sourceRow = customersSheet.getRange(2, 1, 1, customersSheet.getLastColumn());
  const targetRow = customersSheet.getRange(customersSheet.getLastRow(), 1, 1, customersSheet.getLastColumn());
  sourceRow.copyTo(targetRow, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function _digitsOnly(value) {
  return String(value || "").replace(/\D/g, "");
}
