const _customerProps = {
  sheetName: 'Customers',
  name: 'customer',
  fields: {
    firstName: {
      name: 'firstName',
      columnHeader: 'First Name',
      required: true,
      showInForm: true,
    },
    middleName: {
      name: 'middleName',
      columnHeader: 'Middle Name',
      showInForm: true
    },
    lastName: {
      name: 'lastName',
      columnHeader: 'Last Name',
      required: true,
      showInForm: true
    },
    nickname: {
      name: 'nickname',
      columnHeader: 'Nickname',
      showInForm: true
    },
    whatsapp: {
      name: 'whatsapp',
      columnHeader: 'WhatsApp',
      required: true,
      showInForm: true,
      unique: true,
      inputProps: {
        type: 'tel',
        inputmode: 'numeric',
        pattern: '^\+?\d*$',
        maxlength: 14,
        oninput: "this.value = this.value.replace(/[^\\d+]/g, '').replace(/(?!^)\\+/g, '')"
      }
    },
    fullName: {
      name: 'fullName',
      columnHeader: 'Full Name',
      showInForm: false
    },
  },
  itemDialog: {
    height: 390,
    width: 320,
  },
  normalizeFieldValues: normalizeCustomerFieldValues,
  validateValues: validateCustomerFieldValues,
};

addTypeProps(_customerProps);

// TODO: replace with a validateValue prop on the field object of the type
function validateCustomerFieldValues(valuesMapByFieldName, rowNum) {

  const fields = _customerProps.fields;

  const whatsappField = fields.whatsapp;
  const whatsappValue = valuesMapByFieldName[whatsappField.name];

  const rowNums = SOLLibrary.getRowNumbers(_customerProps.sheetName, whatsappField.columnHeader, whatsappValue);
  const duplicateRowNum = rowNums.find(r => r !== rowNum);

  if (duplicateRowNum) {
    const rowValues = SOLLibrary.getRowValues(_customerProps.sheetName, duplicateRowNum);

    throw new Error(JSON.stringify({
      code: 'DUPLICATE_WHATSAPP',
      fieldName: whatsappField.name,
      message: `WhatsApp number +${whatsappValue} is already used by another customer: ${rowValues[fields.fullName.columnHeader]}`,
      row: duplicateRowNum
    }));
  }
}

// TODO: replace with a normalizeValue prop on the field object of the type
function normalizeCustomerFieldValues(values) {
  const fields = _customerProps.fields;

  return {
    [fields.firstName.name]: SOLLibrary.capitalize(values[fields.firstName.name]),
    [fields.middleName.name]: SOLLibrary.capitalize(values[fields.middleName.name]),
    [fields.lastName.name]: SOLLibrary.capitalize(values[fields.lastName.name]),
    [fields.nickname.name]: SOLLibrary.capitalize(values[fields.nickname.name]),
    [fields.whatsapp.name]: Number(values[fields.whatsapp.name].replace(/\D/g, '')),
  };
}