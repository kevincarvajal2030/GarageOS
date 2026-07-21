/**
 * Synchronizes Customer ID from Customer Name
 * and validates Customer Status.
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

  const customerIdCell = sheet.getRange(row, customerIdColumn);

  if (customerName === "") {

    customerIdCell.clearContent();
    return;

  }

  const customerId = findCustomerIdByName(customerName);

  if (!customerId) {

    customerIdCell.clearContent();
    return;

  }

  const customerStatus = findCustomerStatusById(customerId);

  if (
    customerStatus === "Blocked" ||
    customerStatus === "Inactive"
  ) {

    sheet.getRange(row, customerNameColumn).clearContent();
    customerIdCell.clearContent();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Customer is ${customerStatus}. Activate the customer before assigning a vehicle.`,
      APP.NAME,
      5
    );

    return;

  }

  customerIdCell.setValue(customerId);

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

  const customers = getCustomerNames();

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
