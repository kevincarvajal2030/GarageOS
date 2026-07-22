//Customers.gs


/**
 * Synchronizes Customer ID from Customer Name
 * and validates Customer Status.
 */
function syncCustomerReference(event, sheet, row) {

  debug("syncCustomerReference", "START");

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

  debug("Customer Name", customerName);

  const customerId = findCustomerIdByName(customerName);

  debug("Customer ID", customerId);

  if (!customerId) {

    customerIdCell.clearContent();
    return;

  }

  const customerStatus = findCustomerStatusById(customerId);

  debug("Customer Status", customerStatus);

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


function runCustomerBusinessValidations(sheet, row, config, event) {
  if (
    event.range.getColumn() !== config.fields.Status
  ) {
    return;
  }

  validateCustomerStatusChange(sheet, row, config, event);
}


function validateCustomerStatusChange(sheet, row, config, event) {

  Logger.log({
    oldValue: event.oldValue,
    value: event.value,
    column: event.range.getColumn(),
    customerId: sheet.getRange(row, config.fields.CustomerID).getValue()
  });

  const previousStatus = String(event.oldValue || "").trim();
  const newStatus = String(event.value || "").trim();

  if (previousStatus !== "Active") {
    return;
  }

  if (
    newStatus !== "Blocked" &&
    newStatus !== "Inactive"
  ) {
    return;
  }

  const customerId = sheet
    .getRange(row, config.fields.CustomerID)
    .getDisplayValue()
    .trim();

  if (!customerId) {
    return;
  }

  const ss = SpreadsheetApp.getActive();

  const workOrdersConfig = ModuleConfig.get(SHEETS.WORK_ORDERS);

  const workOrdersSheet = ss.getSheetByName(SHEETS.WORK_ORDERS);

  if (!workOrdersSheet) {
    return;
  }

  const lastRow = workOrdersSheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return;
  }

  const workOrders = workOrdersSheet
    .getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      workOrdersSheet.getLastColumn()
    )
    .getValues();

  for (const workOrder of workOrders) {

    const workOrderCustomerId = String(
      workOrder[workOrdersConfig.fields.CustomerID - 1] || ""
    ).trim();

    const workOrderStatus = String(
      workOrder[workOrdersConfig.fields.Status - 1] || ""
    ).trim();

    Logger.log({
      workOrderCustomerId,
      workOrderStatus,
      customerId
    });

    if (workOrderCustomerId !== customerId) {
      continue;
    }

    if (
      workOrderStatus === "In Progress" ||
      workOrderStatus === "Waiting Parts"
    ) {

      event.range.setValue(previousStatus);

      const action =
        newStatus === "Blocked"
          ? "blocked"
          : "set as inactive";

      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Customer cannot be ${action} because there is an active Work Order.`,
        APP.NAME,
        5
      );

      return;
    }

  }

}


/**
* Returns every customer name.
*
* @returns {string[]}
*/
function getCustomerNames() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.CUSTOMERS);

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return [];
  }

  // B = First Name
  // C = Last Name
  // K = Status

  const values = sheet.getRange(
    TABLE.FIRST_DATA_ROW,
    2,
    lastRow - TABLE.FIRST_DATA_ROW + 1,
    10
  ).getDisplayValues();

  return values
    .filter(row => {

      const firstName = String(row[0]).trim();
      const lastName = String(row[1]).trim();
      const status = String(row[9]).trim();

      return (
        firstName &&
        lastName &&
        status === "Active"
      );

    })
    .map(row => `${row[0].trim()} ${row[1].trim()}`);

}


/**
* Updates the CustomerReference list.
*/
function updateCustomerReferenceList() {

  const ss = SpreadsheetApp.getActive();

  const sheet = ss.getSheetByName(
    SHEETS.REFERENCEDATA
  );

  const customers = getCustomerNames();

  const START_ROW = 2;
  const COLUMN = 17; // Q

  sheet
    .getRange(
      START_ROW,
      COLUMN,
      sheet.getMaxRows(),
      1
    )
    .clearContent();

  if (customers.length === 0) return;

  sheet
    .getRange(
      START_ROW,
      COLUMN,
      customers.length,
      1
    )
    .setValues(
      customers.map(name => [name])
    );

}


/**
   * Finds a Customer ID by full customer name.
   *
   * @param {string} customerName
   * @returns {string|null}
   */
function findCustomerIdByName(customerName) {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.CUSTOMERS);

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return null;
  }

  const values = sheet.getRange(
    TABLE.FIRST_DATA_ROW,
    1,
    lastRow - TABLE.FIRST_DATA_ROW + 1,
    3
  ).getValues();

  customerName = String(customerName).trim();

  for (const row of values) {

    const customerId = String(row[0]).trim();
    const firstName = String(row[1]).trim();
    const lastName = String(row[2]).trim();

    const fullName = firstName + " " + lastName;

    if (fullName === customerName) {
      return customerId;
    }

  }

  return null;

}


/**
* Finds the Customer Status by Customer ID.
*
* @param {string} customerId
* @returns {string|null}
*/
function findCustomerStatusById(customerId) {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.CUSTOMERS);

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return null;
  }

  const values = sheet.getRange(
    TABLE.FIRST_DATA_ROW,
    1,
    lastRow - TABLE.FIRST_DATA_ROW + 1,
    11
  ).getValues();

  customerId = String(customerId).trim();

  for (const row of values) {

    if (String(row[0]).trim() === customerId) {
      return String(row[10]).trim();
    }

  }

  return null;

}