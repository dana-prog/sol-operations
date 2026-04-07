function saveItem(item) {
  const normalizedItem = {
    ...item,
    values: _normalizeItemValues(item.typeName, item.values)
  };

  _validateItem(normalizedItem);

  normalizedItem.rowNum ? _updateItem(normalizedItem) : _addItem(normalizedItem);
}

function _addItem(item) {
  const typeName = item.typeName;
  const row = _generateRowValues(item);

  const rowNum= getTypeSheet(item.typeName).appendRow(row).getLastRow();

  _updateItemFormat(typeName, rowNum);
  selectItem(typeName, rowNum);

  SOLLibrary.logArgs('Item', '_addItem', {
    item,
    rowNum,
  });
}

function _updateItem(item) {
  const typeProps = getTypeProps(item.typeName);
  const sheetName = typeProps.sheetName;
  const sheet = SOLLibrary.getSheet(sheetName);

  const row = _generateRowValues(item);
  _getItemRowRange(item.typeName, item.rowNum).setValues([row]);

  SOLLibrary.selectRow(sheetName, rowNum);
  SOLLibrary.logArgs('Item', '_updateItem', {
      typeName,
      rowNum,
      newValuesMapByFieldName,
    });
}

function getItem(typeName, rowNum) {
  const typeProps = getTypeProps(typeName);
  const rowValues = SOLLibrary.getRowValues(typeProps.sheetName, rowNum);
  const fields = typeProps.fields;

  const tz = Session.getScriptTimeZone();

  for (const name in fields) {
    const field = fields[name];

    if (field.inputProps?.type === 'date' && rowValues[name]) {
      const d = rowValues[name]; // Date object
      rowValues[name] = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }
  }

  return Object.fromEntries(
    Object.entries(rowValues).map(([key, value]) => [
      SOLLibrary.toCamelCase(key),
      value
    ])
  );
}

function _formatFieldValues() {

}

function getItems(typeName, fieldName, fieldValue) {
  const typeProps = getTypeProps(typeName);
  SOLLibrary.getRowNumbers(typeProps.sheetName, fieldName, fieldValue);
}

function getTypeSheet(typeName) {
  const typeProps = getTypeProps(typeName);
  return SOLLibrary.getSheet(typeProps.sheetName);
}

function selectItem(typeName, rowNum) {
  SOLLibrary.selectRow(getTypeProps(typeName).sheetName, rowNum);
}

/**
 * Generates a row array for an item by merging provided values with default formulas.
 *
 * Retrieves formulas from the first data row (row 2) for the item's type,
 * and uses them where available, otherwise falls back to the item's values.
 *
 * @param {{typeName: string, values: Object<string, *>}} item Item containing typeName and field values.
 * @returns {Array<*>} Row array with formulas and values merged.
 */
function _generateRowValues(item) {
  // get the formulas from the first data row and merge them with the new row values
  const rowFormulas = _getItemRowRange(item.typeName, 2).getFormulas()[0];
  const rowValuesArr = _getItemRowValuesArr(item);

  return rowValuesArr.map((value, index) => rowFormulas[index] || value);
}

/**
 * Builds a sheet row array for an item from its field values.
 *
 * Maps item field names to column headers and generates a row aligned with the sheet structure.
 *
 * @param {{typeName: string, values: Object<string, *>}} item Item containing typeName and field values.
 * @returns {Array<*>} Row values array aligned with the sheet headers.
 */
function _getItemRowValuesArr(item) {
  const typeProps = getTypeProps(item.typeName);
  const valuesMapByColHeader = Object.fromEntries(
    Object.entries(item.values)
      .map(
        ([fieldName, value]) => [typeProps.fields[fieldName].columnHeader, value]
      )
  );

  return SOLLibrary.generateSheetRows(typeProps.sheetName, [valuesMapByColHeader])[0];
}

/**
 * Returns the full row range for a given item type and row number.
 *
 * @param {string} typeName Item type name (e.g. 'customer', 'event').
 * @param {number} rowNum 1-based row number.
 * @returns {GoogleAppsScript.Spreadsheet.Range} Range covering the entire row.
 */
function _getItemRowRange(typeName, rowNum) {
  const typeProps = getTypeProps(typeName);
  const sheet = SOLLibrary.getSheet(typeProps.sheetName);

  return sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
}

function _normalizeItemValues(typeName, values) {
  const typeProps = getTypeProps(typeName);
  const fields = typeProps.fields;
  return typeProps.normalizeFieldValues ? typeProps.normalizeFieldValues(values) : values;
}

function _validateItem(item) {
  const typeProps = getTypeProps(item.typeName);
  typeProps.validateValues && typeProps.validateValues(item.values);
}

function _updateItemFormat(typeName, rowNum) {
  if (rowNum <= 2) {
    SOLLibrary.log('Item', '_updateItemFormat', `Skipping format update for row ${rowNum}`,
      SOLLibrary.LOG_LEVEL.WARN_LEVEL);
    return;
  }

  // copy the formatting from the first row to rowNum
  const firstItemRange = _getItemRowRange(typeName, 2);
  const newItemRange = _getItemRowRange(typeName, rowNum);
  firstItemRange.copyTo(newItemRange, {formatOnly: true});
}