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

    processRow(sheet, row, config);

  }

  function processRow(sheet, row, config) {

    validateProtectedFields(sheet, row, config);
    validateDuplicateFields(sheet, row, config);
    generateRecordId(sheet, row, config);

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

    if (!hasAllRequiredFields(sheet, row, config)) {
      idCell.clearContent();
      return;
    }

    const currentId = idCell.getDisplayValue().toString().trim();

    if (isValidRecordId(currentId, config)) return;

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

  function validateProtectedFields(sheet, row, config) {
    // TODO
  }

  function validateDuplicateFields(sheet, row, config) {
    // TODO
  }

  return {
    processEdit
  };

})();
