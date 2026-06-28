/**
 * Syncs customer data from _customers (form responses) to Customers sheet.
 *
 * For each record in _customers:
 * - Checks if customer exists by phone or whatsapp
 * - Creates new customer with UUID if not found
 * - Updates existing customer if found
 */
function syncCustomersFromFormResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("_customers");
  const targetSheet = ss.getSheetByName("Customers_");

  if (!sourceSheet || !targetSheet) {
    throw new Error("Required sheets not found: _customers or Customers_");
  }

  // Read all data once
  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  // Build header index maps
  const sourceHeaders = sourceData[0];
  const targetHeaders = targetData[0];
  const sourceIndex = buildHeaderIndex(sourceHeaders);
  const targetIndex = buildHeaderIndex(targetHeaders);

  Logger.log("Source headers: " + sourceHeaders.join(", "));
  Logger.log("Target headers: " + targetHeaders.join(", "));
  Logger.log("Processing " + (sourceData.length - 1) + " source rows...");

  // Process each source row (skip header)
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    const phoneNumber = sourceRow[sourceIndex["Phone number"]] || "";
    const whatsappNumber = sourceRow[sourceIndex["WhatsApp number (ONLY IF DIFFERENT)"]] || "";

    Logger.log(`Processing row ${i + 1}: phone=${phoneNumber}, whatsapp=${whatsappNumber}`);

    // Find existing customer by phone or whatsapp
    let existingRowIndex = -1;
    if (phoneNumber || whatsappNumber) {
      existingRowIndex = findCustomerByPhoneInData(targetData, phoneNumber, whatsappNumber, targetIndex);
    }

    // If not found by phone, try to find by first name + last name
    if (existingRowIndex === -1) {
      const parsed = parseFullName(sourceRow[sourceIndex["Full name and nickname"]] || "");
      if (parsed.firstName && parsed.lastName) {
        existingRowIndex = findCustomerByNameInData(
          targetData,
          capitalizeString(parsed.firstName),
          capitalizeString(parsed.lastName),
          targetIndex
        );
      }
    }

    if (existingRowIndex !== -1) {
      Logger.log(`  → Found existing customer at row ${existingRowIndex + 1}, updating...`);
      updateCustomerInData(targetData, existingRowIndex, sourceRow, sourceIndex, targetIndex);
    } else {
      Logger.log(`  → New customer, creating...`);
      createCustomerInData(targetData, sourceRow, sourceIndex, targetIndex);
    }
  }

  // Write all changes to sheet at once
  Logger.log("Writing " + targetData.length + " rows to sheet...");
  targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);

  // Apply full name formula to rows that were updated/created
  // Get the template formula from row 2
  if (targetData.length > 1) {
    const fullNameCol = targetIndex["full name"] + 1;
    const templateFormula = targetSheet.getRange(2, fullNameCol).getFormula();

    if (templateFormula) {
      // Apply formula to all data rows (skip header at row 1)
      for (let row = 2; row <= targetSheet.getLastRow(); row++) {
        targetSheet.getRange(row, fullNameCol).setFormula(templateFormula);
      }
      Logger.log("Applied full name formula to all customer rows");
    }
  }

  Logger.log("Sync complete!");
}

/**
 * Builds a map of column headers to their indices.
 */
function buildHeaderIndex(headers) {
  const index = {};
  for (let i = 0; i < headers.length; i++) {
    index[headers[i]] = i;
  }
  return index;
}

/**
 * Capitalizes a string: first letter uppercase, rest lowercase.
 * Examples: "JOHN" → "John", "john" → "John", "JoHn" → "John"
 */
function capitalizeString(str) {
  if (!str) return "";
  const trimmed = str.toString().trim();
  if (trimmed.length === 0) return "";
  return trimmed.charAt(0).toUpperCase() + trimmed.substring(1).toLowerCase();
}

/**
 * Normalizes a phone number: removes all non-digits, removes +, replaces leading 0 with 972.
 * Also removes redundant 0 from 9720 and 660 prefixes.
 * Examples:
 *   "+972-50-1234567" → "972501234567"
 *   "05X XXX XXXX" → "972501234567"
 *   "0501234567" → "972501234567"
 *   "9720501234567" → "972501234567"
 *   "66051234567" → "6651234567"
 *   "972501234567" → "972501234567"
 */
function normalizePhoneNumber(phone) {
  if (!phone) return "";

  let normalized = phone.toString().trim();
  // Remove all non-digit characters (this also removes spaces)
  normalized = normalized.replace(/\D/g, "");

  // Replace leading 0 with 972
  if (normalized.startsWith("0")) {
    normalized = "972" + normalized.substring(1);
  }

  // Remove redundant 0 from 9720 prefix
  if (normalized.startsWith("9720")) {
    normalized = "972" + normalized.substring(4);
  }

  // Remove redundant 0 from 660 prefix
  if (normalized.startsWith("660")) {
    normalized = "66" + normalized.substring(3);
  }

  return normalized;
}

/**
 * Finds a customer row by phone or whatsapp number in data array.
 * Searches for the phone/whatsapp in BOTH phone and whatsapp columns.
 * Returns the row index (0-based, excluding header) or -1 if not found.
 */
function findCustomerByPhoneInData(data, phoneNumber, whatsappNumber, headerIndex) {
  const phoneCol = headerIndex["phone"];
  const whatsappCol = headerIndex["whatsapp"];

  if (!phoneNumber && !whatsappNumber) return -1;

  const normalizedPhone = normalizePhoneNumber(phoneNumber);
  const normalizedWhatsapp = normalizePhoneNumber(whatsappNumber);

  // Start from row 1 (skip header at row 0)
  for (let i = 1; i < data.length; i++) {
    const existingPhone = normalizePhoneNumber(data[i][phoneCol] || "");
    const existingWhatsapp = normalizePhoneNumber(data[i][whatsappCol] || "");

    // Check if source phone matches either target phone or target whatsapp
    if (normalizedPhone && (existingPhone === normalizedPhone || existingWhatsapp === normalizedPhone)) {
      return i;
    }

    // Check if source whatsapp matches either target phone or target whatsapp
    if (normalizedWhatsapp && (existingPhone === normalizedWhatsapp || existingWhatsapp === normalizedWhatsapp)) {
      return i;
    }
  }
  return -1;
}

/**
 * Finds a customer row by first name + last name in data array.
 * Only returns a match if the record doesn't have a phone number.
 * Returns the row index (0-based, excluding header) or -1 if not found.
 */
function findCustomerByNameInData(data, firstName, lastName, headerIndex) {
  const firstNameCol = headerIndex["first name"];
  const lastNameCol = headerIndex["last name"];
  const phoneCol = headerIndex["phone"];
  const whatsappCol = headerIndex["whatsapp"];

  if (!firstName || !lastName) return -1;

  // Start from row 1 (skip header at row 0)
  for (let i = 1; i < data.length; i++) {
    const existingFirstName = capitalizeString(data[i][firstNameCol] || "");
    const existingLastName = capitalizeString(data[i][lastNameCol] || "");
    const existingPhone = normalizePhoneNumber(data[i][phoneCol] || "");
    const existingWhatsapp = normalizePhoneNumber(data[i][whatsappCol] || "");

    // Match by first and last name, and only if record has no phone number
    if (existingFirstName === firstName && existingLastName === lastName) {
      if (!existingPhone && !existingWhatsapp) {
        return i;
      } else {
        return -1;
      }
    }
  }
  return -1;
}

/**
 * Parses "Full name and nickname" into components.
 * Format: "FirstName [MiddleName] LastName [(Nickname)]"
 * Returns: {firstName, middleName, lastName, nickname}
 */
function parseFullName(fullNameString) {
  if (!fullNameString) {
    return {firstName: "", middleName: "", lastName: "", nickname: ""};
  }

  let nickname = "";
  let name = fullNameString.toString().trim();

  // Extract nickname from parentheses
  const nicknameMatch = name.match(/\(([^)]+)\)$/);
  if (nicknameMatch) {
    nickname = nicknameMatch[1].trim();
    name = name.substring(0, nicknameMatch.index).trim();
  }

  // Split remaining name into parts
  const parts = name.split(/\s+/).filter(p => p.length > 0);

  if (parts.length === 0) {
    return {firstName: "", middleName: "", lastName: "", nickname: nickname};
  } else if (parts.length === 1) {
    return {firstName: parts[0], middleName: "", lastName: "", nickname: nickname};
  } else if (parts.length === 2) {
    return {firstName: parts[0], middleName: "", lastName: parts[1], nickname: nickname};
  } else {
    // parts.length >= 3: first, middle(s), last
    const firstName = parts[0];
    const lastName = parts[parts.length - 1];
    const middleName = parts.slice(1, parts.length - 1).join(" ");
    return {firstName, middleName, lastName, nickname};
  }
}

/**
 * Creates a new customer record from source data in the data array.
 */
function createCustomerInData(targetData, sourceRow, sourceIndex, targetIndex) {
  const fullNameString = sourceRow[sourceIndex["Full name and nickname"]] || "";
  const parsed = parseFullName(fullNameString);

  // Prepare new row data (with same number of columns as header)
  const newRow = new Array(targetData[0].length).fill("");

  // Generate UUID v4
  newRow[targetIndex["id"]] = Utilities.getUuid();

  // Parse and set name components (capitalized)
  newRow[targetIndex["first name"]] = capitalizeString(parsed.firstName);
  newRow[targetIndex["middle name"]] = capitalizeString(parsed.middleName);
  newRow[targetIndex["last name"]] = capitalizeString(parsed.lastName);
  newRow[targetIndex["nickname"]] = capitalizeString(parsed.nickname);

  // Note: "full name" column has a formula, don't set it

  // Handle phone and whatsapp
  const phoneNumber = sourceRow[sourceIndex["Phone number"]] || "";
  const whatsappNumber = sourceRow[sourceIndex["WhatsApp number (ONLY IF DIFFERENT)"]] || "";

  const normalizedPhone = normalizePhoneNumber(phoneNumber);
  const normalizedWhatsapp = normalizePhoneNumber(whatsappNumber);

  newRow[targetIndex["whatsapp"]] = normalizedWhatsapp;

  if (normalizedPhone) {
    newRow[targetIndex["phone"]] = normalizedPhone;
    // If phone exists and whatsapp doesn't, use phone for both
    if (!normalizedWhatsapp) {
      newRow[targetIndex["whatsapp"]] = normalizedPhone;
    }
  }

  // Set other fields
  newRow[targetIndex["email"]] = sourceRow[sourceIndex["Email"]] || "";
  newRow[targetIndex["nationality"]] = sourceRow[sourceIndex["Nationality"]] || "";
  newRow[targetIndex["birthday"]] = sourceRow[sourceIndex["Birthday"]] || "";
  newRow[targetIndex["allergy/dietary needs"]] = sourceRow[sourceIndex["Any food allergies or dietary needs?"]] || "";
  newRow[targetIndex["emergency contact"]] = sourceRow[sourceIndex["Emergency contact"]] || "";

  // Append new row to data array
  targetData.push(newRow);

  // Note: Formulas will be applied when the data is written back to the sheet in the main function
}

/**
 * Updates an existing customer record with new data from source in the data array.
 */
function updateCustomerInData(targetData, rowIndex, sourceRow, sourceIndex, targetIndex) {

  // Update all fields except id and full name (which has a formula)
  const targetRow = targetData[rowIndex];

  // Always update birthday
  const birthday = sourceRow[sourceIndex["Birthday"]] || "";
  if (birthday) {
    targetRow[targetIndex["birthday"]] = birthday;
  }

  // Handle phone and whatsapp
  const phoneNumber = sourceRow[sourceIndex["Phone number"]] || "";
  const whatsappNumber = sourceRow[sourceIndex["WhatsApp number (ONLY IF DIFFERENT)"]] || "";

  const normalizedPhone = normalizePhoneNumber(phoneNumber);
  const normalizedWhatsapp = normalizePhoneNumber(whatsappNumber);

  if (normalizedWhatsapp) {
    targetRow[targetIndex["whatsapp"]] = normalizedWhatsapp;
  }

  if (normalizedPhone) {
    targetRow[targetIndex["phone"]] = normalizedPhone;
    // If phone exists and whatsapp doesn't, use phone for both
    if (!normalizedWhatsapp) {
      targetRow[targetIndex["whatsapp"]] = normalizedPhone;
    }
  }

  // Update other fields
  const email = sourceRow[sourceIndex["Email"]] || "";
  if (email) {
    targetRow[targetIndex["email"]] = email;
  }

  const nationality = sourceRow[sourceIndex["Nationality"]] || "";
  if (nationality) {
    targetRow[targetIndex["nationality"]] = nationality;
  }

  const allergies = sourceRow[sourceIndex["Any food allergies or dietary needs?"]] || "";
  if (allergies) {
    targetRow[targetIndex["allergy/dietary needs"]] = allergies;
  }

  const emergencyContact = sourceRow[sourceIndex["Emergency contact"]] || "";
  if (emergencyContact) {
    targetRow[targetIndex["emergency contact"]] = emergencyContact;
  }
}
