/**
 * Vehicles.gs  
 * Returns every ACTIVE vehicle belonging to a customer.
 *
 * @param {string} customerId
 * @returns {Array}
 */
function getVehiclesByCustomer(customerId) {

  customerId = String(customerId || "").trim();

  if (!customerId) {
    return [];
  }

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

  values.forEach(row => {

    const vehicleId = String(row[0]).trim();
    const rowCustomerId = String(row[2]).trim();

    const make = String(row[4]).trim();
    const model = String(row[5]).trim();
    const year = String(row[6]).trim();

    // Columna M = Status
    const status = String(row[10]).trim().toLowerCase();

    if (rowCustomerId !== customerId) {
      return;
    }

    // Solo vehículos activos
    if (status !== "active") {
      return;
    }

    vehicles.push({
      id: vehicleId,
      name: `${make} ${model} ${year}`
    });

  });

  debug("Vehicles Result", vehicles);

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


function runVehicleBusinessValidations(sheet, row, config, event) {


}