/**
 * ============================================================================
 * GarageOS
 * AutoRowEngine.gs
 * ============================================================================
 */

const AutoRowEngine = (() => {

  function processEdit(event) {

    if (!event || !event.range) return;

    const sheet = event.range.getSheet();
    const config = ModuleConfig.get(sheet.getName());

    if (!config) return;

    const row = event.range.getRow();

    if (row < TABLE.FIRST_DATA_ROW) return;

    processRow(sheet, row, config, event);

  }

  function processRow(sheet, row, config, event) {

    if (
      sheet.getName() === SHEETS.VEHICLES &&
      event.range.getColumn() === config.fields.CustomerName
    ) {

      syncCustomerReference(event, sheet, row);

    }

    if (
      sheet.getName() === SHEETS.WORK_ORDERS &&
      event.range.getColumn() === config.fields.CustomerName
    ) {

      syncWorkOrderCustomer(sheet, row);

      syncWorkOrderVehicleValidation(sheet, row);

    }

    if (
      sheet.getName() === SHEETS.WORK_ORDERS &&
      event.range.getColumn() === config.fields.VehicleName
    ) {

      syncWorkOrderVehicle(sheet, row);

    }

    if (
      sheet.getName() === SHEETS.WORK_ORDERS &&
      event.range.getColumn() === config.fields.MechanicName
    ) {

      syncWorkOrderMechanic(sheet, row);

    }

    generateRecordId(sheet, row, config);

    formatPhoneField(sheet, row, config, event);

    validateEmailField(sheet, row, config, event);

    validateDuplicateFields(sheet, row, config, event);

    runBusinessValidations(sheet, row, config, event);

    updateReferenceLists(sheet, config, event);
  }

  function hasAllRequiredFields(sheet, row, config) {

    for (const fieldName of config.requiredFields) {

      const column = config.fields[fieldName];

      const value = sheet.getRange(row, column).getDisplayValue().toString().trim();

      if (value === "") return false;

    }

    return true;

  }

  function generateRecordId(sheet, row, config) {

    if (!config.autoGenerateId) return;

    const idField = getIdFieldName(config);
    if (!idField) return;

    const idCell = sheet.getRange(row, config.fields[idField]);

    const currentId = idCell.getDisplayValue().trim();

    // Si el registro ya tiene un ID válido,
    // nunca volver a generarlo ni eliminarlo.
    if (isValidRecordId(currentId, config)) {
      return;
    }

    // Solo generar ID cuando la fila nueva
    // tenga todos los campos requeridos.
    if (!hasAllRequiredFields(sheet, row, config)) {
      return;
    }

    idCell.setValue(generateModuleId(config));

  }

  function isValidRecordId(id, config) {

    return !!id && id.startsWith(config.idPrefix + "-");

  }

  function getIdFieldName(config) {

    return Object.keys(config.fields).find(name => name.endsWith("ID")) || null;

  }

  function generateModuleId(config) {

    switch (config.idPrefix) {

      case "CUS": return generateCustomerID();
      case "VEH": return generateVehicleID();
      case "WO": return generateWorkOrderID();
      case "SER": return generateServiceID();
      case "PRT": return generatePartID();
      case "SUP": return generateSupplierID();
      case "PUR": return generatePurchaseID();
      case "PAY": return generatePaymentID();
      case "MEC": return generateMechanicID();

      default:
        throw new Error("Unknown module prefix: " + config.idPrefix);

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

  function updateReferenceLists(sheet, config, event) {

    const column = event.range.getColumn();

    switch (sheet.getName()) {

      case SHEETS.CUSTOMERS:

        if (
          column === config.fields.FirstName ||
          column === config.fields.LastName ||
          column === config.fields.Status
        ) {

          updateCustomerReferenceList();

        }

        break;

      case SHEETS.VEHICLES:

        if (
          column === config.fields.Make ||
          column === config.fields.Model ||
          column === config.fields.Year
        ) {
          updateVehicleReferenceList();
        }

        break;

      case SHEETS.MECHANICS:

        if (
          column === config.fields.FirstName ||
          column === config.fields.LastName
        ) {
          updateMechanicReferenceList();
        }

        break;

    }

  }


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

  function runVehicleBusinessValidations(sheet, row, config, event) {


  }

  function runWorkOrderBusinessValidations(sheet, row, config, event) {

  }


  return {
    processEdit
  };

})();
