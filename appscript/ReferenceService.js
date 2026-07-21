/**
 * ============================================================
 * GARAGE OS
 * ReferenceService.gs
 *
 * Central service for loading dropdown/reference data
 * from sheet:
 *
 * 10_Reference_Data
 *
 * This service is shared by every module.
 * ============================================================
 */

const ReferenceService = (() => {

  const SHEET_NAME = "10_Reference_Data";

  /**
   * Returns the reference sheet.
   */
  function getSheet() {

    const ss = SpreadsheetApp.getActive();

    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error(
        "Reference sheet not found: " + SHEET_NAME
      );
    }

    return sheet;

  }

  /**
   * Reads a single column until the first empty row.
   */
  function readColumn(startRow, column) {

    const sheet = getSheet();
    const lastRow = sheet.getLastRow();

    const values = sheet
      .getRange(startRow, column, lastRow - startRow + 1, 1)
      .getValues();

    const result = [];

    for (const row of values) {

      const value = String(row[0]).trim();

      if (value === "") {
        break;
      }

      result.push({
        value: value
      });

    }

    return result;

  }

  /**
   * Reads States table.
   *
   * A = Name
   * B = Code
   */
  function getStates() {

    const sheet = getSheet();

    const values = sheet
      .getRange(21, 1, 50, 2)
      .getValues();

    const states = [];

    values.forEach(row => {

      const name = String(row[0]).trim();
      const code = String(row[1]).trim();

      if (name !== "") {

        states.push({

          name: name,

          code: code

        });

      }

    });

    return states;

  }

  function getCustomerStatus() {

    return readColumn(6, 1);

  }

  function getPreferredContact() {

    return readColumn(6, 2);

  }

  function getVehicleStatus() {

    return readColumn(6, 3);

  }

  function getFuelTypes() {

    return readColumn(6, 4);

  }

  function getTransmission() {

    return readColumn(6, 5);

  }

  function getWorkOrderStatus() {

    return readColumn(6, 6);

  }

  function getPriority() {

    return readColumn(6, 7);

  }

  function getPaymentMethods() {

    return readColumn(6, 8);

  }

  function getPaymentStatus() {

    return readColumn(6, 9);

  }

  function getSupplierStatus() {

    return readColumn(6, 10);

  }

  function getMechanicStatus() {

    return readColumn(6, 11);

  }


  /**
   * Customer dropdown data
   */
  function getCustomerReferenceData() {

    return {

      states: getStates(),

      customerStatus: getCustomerStatus(),

      preferredContact: getPreferredContact()

    };

  }

  /**
   * Vehicle dropdown data
   */
  function getVehicleReferenceData() {

    return {

      vehicleStatus: getVehicleStatus(),

      fuelTypes: getFuelTypes(),

      transmission: getTransmission()

    };

  }

  /**
   * Work Order dropdown data
   */
  function getWorkOrderReferenceData() {

    return {

      workOrderStatus: getWorkOrderStatus(),

      priority: getPriority()

    };

  }

  /**
   * Payment dropdown data
   */
  function getPaymentReferenceData() {

    return {

      paymentMethods: getPaymentMethods(),

      paymentStatus: getPaymentStatus()

    };

  }

  /**
   * Supplier dropdown data
   */
  function getSupplierReferenceData() {

    return {

      supplierStatus: getSupplierStatus()

    };

  }

  /**
   * Mechanic dropdown data
   */
  function getMechanicReferenceData() {

    return {

      mechanicStatus: getMechanicStatus()

    };

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

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      2,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      2
    ).getValues();

    const customers = [];

    values.forEach(row => {

      const firstName = String(row[0]).trim();
      const lastName = String(row[1]).trim();

      if (!firstName && !lastName) {
        return;
      }

      customers.push(firstName + " " + lastName);

    });

    return customers;

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
 * Returns every vehicle belonging to a customer.
 *
 * @param {string} customerId
 * @returns {Array}
 */
  function getVehiclesByCustomer(customerId) {

    const sheet = SpreadsheetApp
      .getActive()
      .getSheetByName(SHEETS.VEHICLES);

    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      13
    ).getValues();

    const vehicles = [];

    customerId = String(customerId).trim();

    values.forEach(row => {

      const vehicleId = String(row[0]).trim();

      const rowCustomerId = String(row[2]).trim();

      const make = String(row[4]).trim();

      const model = String(row[5]).trim();

      const year = String(row[6]).trim();

      if (rowCustomerId !== customerId) {
        return;
      }

      vehicles.push({

        id: vehicleId,

        name: `${make} ${model} ${year}`

      });

    });

    return vehicles;

  }

  /**
 * Finds the Vehicle ID from Customer ID and Vehicle Name.
 *
 * @param {string} customerId
 * @param {string} vehicleName
 * @returns {string|null}
 */
  function findVehicleIdByName(customerId, vehicleName) {

    const vehicles = getVehiclesByCustomer(customerId);

    vehicleName = String(vehicleName).trim();

    for (const vehicle of vehicles) {

      if (vehicle.name === vehicleName) {
        return vehicle.id;
      }

    }

    return null;

  }

  /**
 * Finds a Mechanic ID by Mechanic Name.
 *
 * @param {string} mechanicName
 * @returns {string|null}
 */
  function findMechanicIdByName(mechanicName) {

    const sheet = SpreadsheetApp
      .getActive()
      .getSheetByName(SHEETS.MECHANICS);

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

    mechanicName = String(mechanicName).trim();

    for (const row of values) {

      const mechanicId = String(row[0]).trim();

      const firstName = String(row[1]).trim();

      const lastName = String(row[2]).trim();

      const fullName = firstName + " " + lastName;

      if (fullName === mechanicName) {
        return mechanicId;
      }

    }

    return null;

  }

  /**
   * Public API
   */
  return {

    getStates,

    getCustomerNames,

    findCustomerIdByName,

    getVehiclesByCustomer,

    findVehicleIdByName,

    findMechanicIdByName,

    getCustomerStatus,

    getPreferredContact,

    getVehicleStatus,

    getFuelTypes,

    getTransmission,

    getWorkOrderStatus,

    getPriority,

    getPaymentMethods,

    getPaymentStatus,

    getSupplierStatus,

    getMechanicStatus,

    getCustomerReferenceData,

    getVehicleReferenceData,

    getWorkOrderReferenceData,

    getPaymentReferenceData,

    getSupplierReferenceData,

    getMechanicReferenceData

  };

})();

/**
 * ============================================================
 * GAS Wrappers
 * These functions are callable from HTML via google.script.run
 * ============================================================
 */

function getCustomerReferenceData() {
  return ReferenceService.getCustomerReferenceData();
}

function getVehicleReferenceData() {
  return ReferenceService.getVehicleReferenceData();
}

function getWorkOrderReferenceData() {
  return ReferenceService.getWorkOrderReferenceData();
}

function getPaymentReferenceData() {
  return ReferenceService.getPaymentReferenceData();
}

function getSupplierReferenceData() {
  return ReferenceService.getSupplierReferenceData();
}

function getMechanicReferenceData() {
  return ReferenceService.getMechanicReferenceData();
}

function findCustomerIdByName(customerName) {
  return ReferenceService.findCustomerIdByName(customerName);
}

function getVehiclesByCustomer(customerId) {
  return ReferenceService.getVehiclesByCustomer(customerId);
}

function findVehicleIdByName(customerId, vehicleName) {
  return ReferenceService.findVehicleIdByName(
    customerId,
    vehicleName
  );
}

function findMechanicIdByName(mechanicName) {
  return ReferenceService.findMechanicIdByName(mechanicName);
}

function getCustomerNames() {
  return ReferenceService.getCustomerNames();
}
