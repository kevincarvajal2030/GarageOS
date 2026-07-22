//Vehicles.cs

/**
 * Synchronizes Customer ID from Customer Name
 * and validates Customer Status.
 */
function syncCustomerReference(event, sheet, row) {

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

  const customerId = findCustomerIdByName(customerName);

  if (!customerId) {

    customerIdCell.clearContent();
    return;

  }

  const customerStatus = findCustomerStatusById(customerId);

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


function runVehicleBusinessValidations(sheet, row, config, event) {


}