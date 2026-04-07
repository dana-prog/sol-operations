function testDate() {
  const sheetName = 'Events';
  const rowNum = SOLLibrary.getSelectedRow(sheetName);
  const rowValues = SOLLibrary.getRowValues(sheetName, rowNum);
  const startDate = rowValues['Start Date'];
  const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  SOLLibrary.logArgs('playground', 'testDate', {startDate, formattedStartDate});
}

// // finds the row according to the value in the unique field
// function _selectRowByValues(typeName, values) {
//   const typeProps = getTypeProp(typeName);
//   const uniqueField = getUniqueField(typeName);
//   const rowNum = SOLLibrary.getRowNumbers(
//     sheet,
//     uniqueField.columnHeader,
//     values[uniqueField.columnHeader])[0];
//
//   SOLLibrary.selectRow(typeProps.sheetName, rowNum);
//   return rowNum;
// }

// function getUniqueField(type) {
//   const uniqueFields = Object.values(itemsProps[type].fields).filter(field => field.unique);
//   if (uniqueFields.length !== 1) {
//     SOLLibrary.logArgs(
//       'TypeProps',
//       'getUniqueFields',
//       {
//         type,
//         props: itemsProps[type]
//       },
//       `There should be one unique field for ${type} type, but found ${uniqueFields.length}`
//     );
//   }
//   return uniqueFields[0];
// }