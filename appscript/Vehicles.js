/**
 * Synchronizes Customer ID from Customer Name
 * and validates Customer Status.
 */
function syncCustomerReference(event, sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerNameColumn = config.fields.CustomerName;
  const customerIdColumn = config.fields.CustomerID;

  const customerNameCell = sheet.getRange(row, customerNameColumn);
  const customerIdCell = sheet.getRange(row, customerIdColumn);

  const customerName = customerNameCell
    .getDisplayValue()
    .trim();

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

    if (event.oldValue !== undefined) {

      customerNameCell.setValue(event.oldValue);

      const previousCustomerId =
        findCustomerIdByName(event.oldValue);

      if (previousCustomerId) {
        customerIdCell.setValue(previousCustomerId);
      } else {
        customerIdCell.clearContent();
      }

    } else {

      customerNameCell.clearContent();
      customerIdCell.clearContent();

    }

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
