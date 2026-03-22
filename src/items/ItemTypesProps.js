const _itemTypesProps = {};

function addTypeProps(typeProps) {
  _itemTypesProps[typeProps.name] = typeProps;
}

function getTypeProps(typeName) {
  return _itemTypesProps[typeName];
}

function getTypePropsBySheetName(sheetName) {
  return Object.values(_itemTypesProps)
    .find(item => item.sheetName === sheetName) || null;
}