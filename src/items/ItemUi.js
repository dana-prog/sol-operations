function showNewItemDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const typeName = getTypePropsBySheetName(sheet.getName()).name;

  showItemDialog(typeName);
}

function showEditItemDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const typeName = getTypePropsBySheetName(sheet.getName()).name;
  const rowNum = SOLLibrary.getSelectedRow(sheet.getName());

  showItemDialog(typeName, rowNum);
}

function showItemDialog(typeName, rowNum) {
  const typeProps = getTypeProps(typeName);
  const dialogProps = typeProps.itemDialog;
  const initialData = rowNum ? getItem(typeName, rowNum) : null;
  const fields = Object.values(typeProps.fields).filter(field => field.showInForm);

  const template = HtmlService.createTemplateFromFile(`html/item_form.tmpl`);
  template.typeName = typeName;
  template.fields = fields;
  template.rowNum = rowNum;
  template.initialData = initialData;

  SOLLibrary.logArgs('ItemUiUtils', 'showItemDialog', {
    template,
    rowNum,
  });

  const html = template
    .evaluate()
    .setHeight(dialogProps?.height || 400)
    .setWidth(dialogProps?.width || 400);

  const title = `${rowNum ? 'Edit' : 'New'} ${SOLLibrary.capitalize(typeProps.name)}`;
  SpreadsheetApp.getUi().showModalDialog(html, title);
}
