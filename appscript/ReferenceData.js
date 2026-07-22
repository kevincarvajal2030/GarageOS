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


  return {

    // Generic

    getStates,

    // Customer

    getCustomerStatus,

    getPreferredContact,

    getCustomerReferenceData,

    // Vehicle

    getVehicleStatus,

    getFuelTypes,

    getTransmission,

    getVehicleReferenceData,

    // Work Orders

    getWorkOrderStatus,

    getPriority,

    getWorkOrderReferenceData,

    // Payments

    getPaymentMethods,

    getPaymentStatus,

    getPaymentReferenceData,

    // Suppliers

    getSupplierStatus,

    getSupplierReferenceData,

    // Mechanics

    getMechanicStatus,

    getMechanicReferenceData

  };

})();


