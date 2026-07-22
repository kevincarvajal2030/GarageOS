//Validation.gs

function runBusinessValidations(sheet, row, config, event) {

  switch (sheet.getName()) {

    case SHEETS.CUSTOMERS:
      runCustomerBusinessValidations(sheet, row, config, event);
      break;

    case SHEETS.VEHICLES:
      runVehicleBusinessValidations(sheet, row, config, event);
      break;

    case SHEETS.WORK_ORDERS:
      runWorkOrderBusinessValidations(sheet, row, config, event);
      break;

  }

}


function formatPhoneNumber(phone) {

  const digits = String(phone).replace(/\D/g, "");

  if (digits.length !== 10) {
    return phone;
  }

  return `(${digits.substring(0, 3)}) ${digits.substring(3, 6)}-${digits.substring(6)}`;

}


function formatPhoneField(sheet, row, config, event) {

  const column = config.fields.Phone;

  if (!column) return;

  const cell = sheet.getRange(row, column);

  const value = cell.getDisplayValue().trim();

  if (!value) return;

  const digits = value.replace(/\D/g, "");

  if (digits.length !== 10) {

    if (event.oldValue !== undefined) {

      cell.setValue(event.oldValue);

    } else {

      cell.clearContent();

    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Phone number must contain exactly 10 digits.",
      APP.NAME,
      3
    );

    return;

  }

  const formatted = formatPhoneNumber(digits);

  if (formatted !== value) {
    cell.setValue(formatted);
  }

}



function validateEmailField(sheet, row, config, event) {

  const column = config.fields.Email;

  if (!column) return;

  const cell = sheet.getRange(row, column);

  const value = cell.getDisplayValue().trim();

  if (!value) return;

  const emailRegex =
    /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

  if (!emailRegex.test(value)) {

    if (event.oldValue !== undefined) {

      cell.setValue(event.oldValue);

    } else {

      cell.clearContent();

    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Invalid email address.",
      APP.NAME,
      3
    );

  }

}


function validateDuplicateFields(sheet, row, config, event) {

  if (!config.duplicateFields || config.duplicateFields.length === 0) {
    return;
  }

  config.duplicateFields.forEach(field => {

    const column = config.fields[field];

    if (!column) return;

    if (event.range.getColumn() !== column) return;

    const value = sheet
      .getRange(row, column)
      .getDisplayValue()
      .trim();

    if (!value) return;

    const lastRow = sheet.getLastRow();

    const values = sheet
      .getRange(TABLE.FIRST_DATA_ROW, column, lastRow - TABLE.FIRST_DATA_ROW + 1)
      .getDisplayValues();

    let duplicated = false;

    for (let i = 0; i < values.length; i++) {

      const currentRow = TABLE.FIRST_DATA_ROW + i;

      if (currentRow === row) continue;

      if (values[i][0].trim() === value) {
        duplicated = true;
        break;
      }

    }

    if (!duplicated) return;

    const cell = sheet.getRange(row, column);

    if (event.oldValue !== undefined) {
      cell.setValue(event.oldValue);
    } else {
      cell.clearContent();
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      field + " already exists.",
      APP.NAME,
      3
    );

  });

}