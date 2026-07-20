/**
 * Synchronizes Customer ID from Customer Name.
 */
function syncCustomerId(sheet, row) {

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
