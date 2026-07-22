//Customers.gs


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