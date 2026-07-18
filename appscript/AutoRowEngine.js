/**
 * ============================================================================
 * GarageOS
 * AutoRowEngine.gs
 * ----------------------------------------------------------------------------
 * Generic engine for automatic row processing.
 *
 * Responsibilities:
 * - Detect edits
 * - Process managed modules
 * - Coordinate validation and ID generation
 *
 * Business rules are NOT defined here.
 * They come from ModuleConfig.gs.
 * ============================================================================
 */

const AutoRowEngine = (() => {

  /**
   * Processes an edit event.
   *
   * @param {GoogleAppsScript.Events.SheetsOnEdit} event
   */
  function processEdit(event) {

    if (!event || !event.range) {
      return;
    }

    const sheet = event.range.getSheet();
    const sheetName = sheet.getName();

    const config = ModuleConfig.get(sheetName);

    if (!config) {
      return;
    }

    const row = event.range.getRow();

    if (row < FIRST_DATA_ROW) {
      return;
    }

    processRow(sheet, row, config);

  }

  /**
   * Processes one data row.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   */
  function processRow(sheet, row, config) {

    validateProtectedFields(
      sheet,
      row,
      config
    );

    validateRequiredFields(
      sheet,
      row,
      config
    );

    validateDuplicateFields(
      sheet,
      row,
      config
    );

    generateRecordId(
      sheet,
      row,
      config
    );

  }


  /**
   * Validates required fields before continuing.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   */
  function validateRequiredFields(sheet, row, config) {

    if (!hasAllRequiredFields(sheet, row, config)) {
      return false;
    }

    return true;

  }

  /**
   * Returns true when all required fields contain values.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   * @returns {boolean}
   */
  function hasAllRequiredFields(sheet, row, config) {

    const fields = config.fields;

    for (const fieldName of config.requiredFields) {

      const column = fields[fieldName];

      const value = sheet
        .getRange(row, column)
        .getValue()
        .toString()
        .trim();

      if (value === "") {
        return false;
      }

    }

    return true;

  }


  /**
   * Generates a record ID when needed.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   */
  function generateRecordId(sheet, row, config) {

    if (!config.autoGenerateId) {
      return;
    }

    const idField = getIdFieldName(config);

    const idColumn = config.fields[idField];

    const currentId = sheet
      .getRange(row, idColumn)
      .getValue()
      .toString()
      .trim();

    if (currentId !== "") {
      return;
    }

    const newId = generateModuleId(config);

    sheet
      .getRange(row, idColumn)
      .setValue(newId);

  }

  /**
   * Returns the ID field name for the current module.
   *
   * @param {Object} config
   * @returns {string}
   */
  function getIdFieldName(config) {

    return Object.keys(config.fields).find(fieldName =>
      fieldName.endsWith("ID")
    );

  }

  /**
   * Generates a new formatted ID.
   *
   * @param {Object} config
   * @returns {string}
   */
  function generateModuleId(config) {

    switch (config.idPrefix) {

      case "CUS":
        return generateCustomerID();

      case "VEH":
        return generateVehicleID();

      case "WO":
        return generateWorkOrderID();

      case "SER":
        return generateServiceID();

      case "PRT":
        return generatePartID();

      case "SUP":
        return generateSupplierID();

      case "PUR":
        return generatePurchaseID();

      case "PAY":
        return generatePaymentID();

      case "MEC":
        return generateMechanicID();

      default:
        throw new Error(
          "Unknown ID prefix: " + config.idPrefix
        );

    }

  }

  /**
   * Validates protected fields.
   *
   * Placeholder for future implementation.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   */
  function validateProtectedFields(sheet, row, config) {

    // TODO:
    // Prevent manual modification of protected fields.

  }

  /**
   * Validates duplicate records.
   *
   * Placeholder for future implementation.
   *
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Object} config
   */
  function validateDuplicateFields(sheet, row, config) {

    // TODO:
    // Detect duplicate records before generating IDs.

  }

  return {

    processEdit

  };

})();