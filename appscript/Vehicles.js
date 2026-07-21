/**
 * Synchronizes Customer ID from Customer Name.
 */
function syncCustomerReference(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerNameColumn = config.fields.CustomerName;
  const customerIdColumn = config.fields.CustomerID;

  const customerName = sheet
    .getRange(row, customerNameColumn)
    .getDisplayValue()
    .trim();

  if (customerName === "") {

    sheet.getRange(row, customerIdColumn).clearContent();

    return;

  }

  const customerId = findCustomerIdByName(customerName);

  if (!customerId) {

    sheet.getRange(row, customerIdColumn).clearContent();

    return;

  }

  sheet
    .getRange(row, customerIdColumn)
    .setValue(customerId);

}


/**
 * Refreshes the Customer Name dropdown using only ACTIVE customers.
 */
function refreshCustomerDropdown() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const customerColumn =
    MODULE_CONFIG.VEHICLES.columns.customerName;

  const customers = getActiveCustomerNames();

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(customers, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(
      TABLE.FIRST_DATA_ROW,
      customerColumn,
      sheet.getMaxRows() - TABLE.FIRST_DATA_ROW + 1,
      1
    )
    .setDataValidation(rule);

}
