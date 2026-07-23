function openVehicleViewerSidebar() {

  const html = HtmlService
    .createHtmlOutputFromFile("VehicleViewer")
    .setTitle("Vehicle Viewer");

  SpreadsheetApp
    .getUi()
    .showSidebar(html);

}


/**
 * ============================================================
 * Creates a standardized Vehicle object from a sheet row.
 *
 * @param {Array} row Row values from 02_Vehicles.
 * @returns {Object} Immutable vehicle object.
 * ============================================================
 */
function createVehicleObject(row) {

  const fields = ModuleConfig
    .get(SHEETS.VEHICLES)
    .fields;

  const get = (field) => {

    const value = row[fields[field] - 1];

    return typeof value === "string"
      ? value.trim()
      : value;

  };

  // Read values only once
  const vehicleId = get("VehicleID");

  const customerId = get("CustomerID");

  const customerName = get("CustomerName");

  const licensePlate = get("LicensePlate");

  const make = get("Make");

  const model = get("Model");

  const year = get("Year");

  const transmission = get("Transmission");

  const color = get("Color");

  const fuelType = get("FuelType");

  const status = get("Status");

  const vehicleName = get("VehicleName");
  
  const imageFileId = get("ImageFileID") || "";

  debug("VehicleID", vehicleId);
  debug("VehicleName", vehicleName);
  debug("ImageFileID", imageFileId);

  return Object.freeze({

    vehicleId,
    customerId,
    customerName,
    licensePlate,
    make,
    model,
    year,

    transmission,
    color,
    fuelType,

    status,

    vehicleName,

    displayName: `${make} ${model} ${year}`.trim(),

    imageFileId

  });

}


/**
 * ============================================================
 * Returns the currently selected vehicle.
 *
 * Selection priority:
 *
 * 1. Vehicles sheet
 * 2. Work Orders sheet
 * 3. First vehicle in database
 *
 * @returns {Object|null}
 * ============================================================
 */
function getSelectedVehicle() {

  const ss = SpreadsheetApp.getActive();

  const sheet = ss.getActiveSheet();

  const sheetName = sheet.getName();

  let vehicle = null;

  switch (sheetName) {

    case "02_Vehicles":
      vehicle = getSelectedVehicleFromVehicles_(sheet);
      break;

    case "03_Work_Orders":
      vehicle = getSelectedVehicleFromWorkOrders_(sheet);
      break;

    default:
      vehicle = getDefaultVehicle();

  }

  return vehicle;

}



/**
 * ============================================================
 * Returns the selected vehicle from 02_Vehicles.
 *
 * @param {Sheet} sheet
 * @returns {Object|null}
 * ============================================================
 */
function getSelectedVehicleFromVehicles_(sheet) {

  const row = sheet.getActiveRange().getRow();

  const FIRST_DATA_ROW = TABLE.FIRST_DATA_ROW;

  if (row < FIRST_DATA_ROW) {
    return getDefaultVehicle();
  }

  const lastColumn = sheet.getLastColumn();

  const values = sheet
    .getRange(row, 1, 1, lastColumn)
    .getValues()[0];

  return createVehicleObject(values);

}


/**
 * ============================================================
 * Returns the selected vehicle from 03_Work_Orders.
 *
 * @param {Sheet} sheet
 * @returns {Object|null}
 * ============================================================
 */
function getSelectedVehicleFromWorkOrders_(sheet) {

  return getDefaultVehicle();

}


/**
 * ============================================================
 * Returns the first available vehicle.
 *
 * @returns {Object|null}
 * ============================================================
 */
function getDefaultVehicle() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("02_Vehicles");

  const FIRST_DATA_ROW = TABLE.FIRST_DATA_ROW;

  const lastRow = sheet.getLastRow();

  if (lastRow < FIRST_DATA_ROW) {
    return null;
  }

  const lastColumn = sheet.getLastColumn();

  const row = sheet
    .getRange(FIRST_DATA_ROW, 1, 1, lastColumn)
    .getValues()[0];

  return createVehicleObject(row);

}


/**
 * ============================================================
 * Returns the data required by Vehicle Viewer.
 *
 * This function acts as the bridge between Apps Script
 * and the HTML sidebar.
 *
 * @returns {Object|null}
 * ============================================================
 */
function getVehicleViewerData() {

  return getSelectedVehicle();

}










