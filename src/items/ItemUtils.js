function saveItem(typeName, data, rowNum = -1) {
  const typeProps = getTypeProps(typeName);
  const valuesMap = typeProps.normalizeFieldValues ? typeProps.normalizeFieldValues(data) : data;
  typeProps.validateFieldValues && typeProps.validateFieldValues(valuesMap, rowNum);

  if (rowNum > 0) {
    _updateRow(typeName, valuesMap, rowNum);
  } else {
    _addRow(typeName, valuesMap);
  }
}

function _addRow(typeName, valuesMapByFieldName) {
  const typeProps = getTypeProps(typeName);
  const sheetName = typeProps.sheetName;
  const sheet = SOLLibrary.getSheet(sheetName);

  const valuesMapByColHeader = Object.fromEntries(
    Object.entries(valuesMapByFieldName)
      .map(
        ([fieldName, value]) => [typeProps.fields[fieldName].columnHeader, value]
      )
  );

  const rowValuesArr = SOLLibrary.generateSheetRows(sheetName, [valuesMapByColHeader])[0];

  // use the first data row (row 2) to copy formulas and formatting from
  const firstDataRowRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());

  // get the formulas from the first data row and merge them with the new row values
  const rowFormulas = firstDataRowRange.getFormulas()[0];
  const rowValuesWithFormulasArr = rowValuesArr.map((value, index) => rowFormulas[index] || value);
  sheet.appendRow(rowValuesWithFormulasArr);

  // copy the formatting from the first data row to the new row
  const rowNum = sheet.getLastRow();
  const newRowRange = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  firstDataRowRange.copyTo(newRowRange, {formatOnly: true});

  SOLLibrary.selectRow(sheetName, rowNum);
  // TODO: consider sorting the rows after add/update
  // SOLLibrary.sortSheet(sheetName, typeProps.fields.fullName.columnHeader);
  // const newRowNum = _selectRowByValues(typeName, valuesMap);

  SOLLibrary.logArgs('ItemUtils', '_addRow', {
      typeName,
      rowNum,
      valuesMapByFieldName,
    }
  )
}

function _updateRow(typeName, newValuesMapByFieldName, rowNum) {
  const typeProps = getTypeProps(typeName);
  const sheetName = typeProps.sheetName;
  const sheet = SOLLibrary.getSheet(sheetName);

  const rowRange = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  const newValuesMapByColHeader = Object.fromEntries(
    Object.entries(newValuesMapByFieldName)
      .map(
        ([fieldName, value]) => [typeProps.fields[fieldName].columnHeader, value]
      )
  );

  const newRowValuesArr = SOLLibrary.generateSheetRows(sheetName, [newValuesMapByColHeader])[0];

  const rowFormulas = rowRange.getFormulas()[0];
  const rowValuesWithFormulasArr = newRowValuesArr.map((value, index) => value || rowFormulas[index]);

  rowRange.setValues([rowValuesWithFormulasArr]);
  SOLLibrary.selectRow(sheetName, rowNum);
  SOLLibrary.logArgs('ItemUtils', '_updateRow', {
      typeName,
      rowNum,
      newValuesMapByFieldName,
    }
  )
}

function getItemValues(typeName, rowNum) {
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

function findItems(typeName, fieldName, fieldValue) {
  const typeProps = getTypeProps(typeName);
  SOLLibrary.findRows(typeProps.sheetName, fieldName, fieldValue);
}
