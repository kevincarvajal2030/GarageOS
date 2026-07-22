/**
 * Synchronizes Customer ID from Customer Name.
 *
 * Work Orders:
 * Customer Name  -> Customer ID
 */
function syncWorkOrderCustomer(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerNameColumn = config.fields.CustomerName;
  const customerIdColumn = config.fields.CustomerID;

  const customerName = sheet
    .getRange(row, customerNameColumn)
    .getDisplayValue()
    .trim();

  if (customerName === "") {

    sheet
      .getRange(row, customerIdColumn)
      .clearContent();

    return;

  }

  const customerId = findCustomerIdByName(customerName);

  if (!customerId) {

    sheet
      .getRange(row, customerIdColumn)
      .clearContent();

    return;

  }

  sheet
    .getRange(row, customerIdColumn)
    .setValue(customerId);

}


/**
 * Creates the Vehicle dropdown filtered by Customer ID.
 *
 * Work Orders:
 * Customer ID -> Vehicle dropdown
 */
function syncWorkOrderVehicleValidation(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerIdColumn = config.fields.CustomerID;
  const vehicleNameColumn = config.fields.VehicleName;
  const vehicleIdColumn = config.fields.VehicleID;

  const customerId = sheet
    .getRange(row, customerIdColumn)
    .getDisplayValue()
    .trim();

  const vehicleNameCell = sheet.getRange(row, vehicleNameColumn);
  const vehicleIdCell = sheet.getRange(row, vehicleIdColumn);

  // Clear previous vehicle selection before rebuilding the dropdown.
  vehicleNameCell.clearContent();
  vehicleIdCell.clearContent();

  if (customerId === "") {

    vehicleNameCell.clearDataValidations();

    return;

  }

  const vehicles = getVehiclesByCustomer(customerId);

  const vehicleNames = vehicles.map(vehicle => vehicle.name);

  if (vehicleNames.length === 0) {

    vehicleNameCell.clearDataValidations();

    return;

  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vehicleNames, true)
    .setAllowInvalid(false)
    .build();

  vehicleNameCell.setDataValidation(rule);

}



/**
 * Synchronizes Vehicle ID from Vehicle Name.
 *
 * Work Orders:
 * Vehicle Name -> Vehicle ID
 */
function syncWorkOrderVehicle(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerIdColumn = config.fields.CustomerID;
  const vehicleNameColumn = config.fields.VehicleName;
  const vehicleIdColumn = config.fields.VehicleID;

  const customerId = sheet
    .getRange(row, customerIdColumn)
    .getDisplayValue()
    .trim();

  const vehicleName = sheet
    .getRange(row, vehicleNameColumn)
    .getDisplayValue()
    .trim();

  const vehicleIdCell = sheet.getRange(row, vehicleIdColumn);

  if (customerId === "" || vehicleName === "") {

    vehicleIdCell.clearContent();

    return;

  }

  const vehicleId = findVehicleIdByName(
    customerId,
    vehicleName
  );

  if (!vehicleId) {

    vehicleIdCell.clearContent();

    return;

  }

  vehicleIdCell.setValue(vehicleId);

}


/**
 * Synchronizes Mechanic ID from Mechanic Name.
 *
 * Work Orders:
 * Mechanic Name -> Mechanic ID
 */
function syncWorkOrderMechanic(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const mechanicNameColumn = config.fields.MechanicName;
  const mechanicIdColumn = config.fields.MechanicID;

  const mechanicName = sheet
    .getRange(row, mechanicNameColumn)
    .getDisplayValue()
    .trim();

  const mechanicIdCell = sheet.getRange(row, mechanicIdColumn);

  if (mechanicName === "") {

    mechanicIdCell.clearContent();

    return;

  }

  const mechanicId = findMechanicIdByName(mechanicName);

  if (!mechanicId) {

    mechanicIdCell.clearContent();

    return;

  }

  mechanicIdCell.setValue(mechanicId);

}




function runWorkOrderBusinessValidations(sheet, row, config, event) {

}


