/**
 * ============================================================================
 * GarageOS
 * AutoRowEngine.gs
 * ----------------------------------------------------------------------------
 * Central event orchestrator for GarageOS.
 *
 * Responsibilities:
 * - Receives every onEdit event.
 * - Detects which module triggered the edit.
 * - Executes the processing pipeline.
 *
 * Every module is responsible for its own business rules.
 * AutoRowEngine only coordinates execution.
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

      syncCustomerReference(event, sheet, row);

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

    processAutomaticFields(sheet, row, config);

    processValidations(sheet, row, config, event);

    processReferenceUpdates(sheet, config, event);
  }

  function processAutomaticFields(sheet, row, config) {

    generateRecordId(sheet, row, config);

  }

  function processValidations(sheet, row, config, event) {

    formatPhoneField(sheet, row, config, event);

    validateEmailField(sheet, row, config, event);

    validateDuplicateFields(sheet, row, config, event);

    runBusinessValidations(sheet, row, config, event);

  }

  function processReferenceUpdates(sheet, config, event) {

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


  return {
    processEdit
  };

})();
