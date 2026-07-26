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

  vehicleNameCell.clearContent();
  vehicleIdCell.clearContent();
  vehicleNameCell.clearDataValidations();

  if (!customerId) {
    return;
  }

  const vehicles = getVehiclesByCustomer(customerId);

  debug("Vehicles Found", vehicles);

  if (vehicles.length === 0) {

    showToast(
      "This customer has no active vehicles.",
      "GarageOS"
    );

    return;

  }

  const vehicleNames = vehicles.map(v => v.name);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vehicleNames, true)
    .setAllowInvalid(false)
    .build();

  vehicleNameCell.setDataValidation(rule);

  // Auto-select first vehicle
  vehicleNameCell.setValue(vehicleNames[0]);

  // Auto-fill Vehicle ID
  vehicleIdCell.setValue(vehicles[0].id);

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


/**
 * Validates mechanic selection when a user selects a mechanic from the dropdown.
 * Reuses the same validation pattern as Vehicles → Customer Name.
 * Rejects selection if mechanic is Busy, Vacation, or Inactive.
 *
 * @param {Sheet} sheet - The sheet being edited.
 * @param {number} row - The row being edited.
 * @param {Object} config - The module configuration.
 * @param {Object} event - The edit event object.
 */
function validateMechanicSelection(sheet, row, config, event) {

  const mechanicNameColumn = config.fields.MechanicName;
  const mechanicIdColumn = config.fields.MechanicID;

  // Only act if the Mechanic Name column was edited
  if (event.range.getColumn() !== mechanicNameColumn) {
    return;
  }

  const mechanicName = String(event.value || "").trim();
  const mechanicNameCell = sheet.getRange(row, mechanicNameColumn);
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

  // Check mechanic availability using the service (includes self-healing)
  const availability = MechanicAssignmentService.isMechanicAvailable(mechanicId);

  if (!availability.available) {

    // Build specific toast message based on the reason
    let toastMessage = "This mechanic cannot be assigned because they are currently unavailable.";

    if (availability.reason) {

      const reasonLower = availability.reason.toLowerCase();

      if (reasonLower.includes("busy")) {
        toastMessage = "This mechanic cannot be assigned because they are currently busy with other work orders.";
      } else if (reasonLower.includes("vacation")) {
        toastMessage = "This mechanic cannot be assigned because they are currently on vacation.";
      } else if (reasonLower.includes("inactive")) {
        toastMessage = "This mechanic cannot be assigned because their status is inactive.";
      } else if (reasonLower.includes("no mechanic selected")) {
        toastMessage = "No mechanic has been selected.";
      } else if (reasonLower.includes("not found")) {
        toastMessage = "This mechanic was not found in the system.";
      }

    }

    // Restore previous value
    if (event.oldValue !== undefined) {

      mechanicNameCell.setValue(event.oldValue);

      const previousMechanicId = findMechanicIdByName(String(event.oldValue).trim());

      if (previousMechanicId) {
        mechanicIdCell.setValue(previousMechanicId);
      } else {
        mechanicIdCell.clearContent();
      }

    } else {

      mechanicNameCell.clearContent();
      mechanicIdCell.clearContent();

    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      toastMessage,
      APP.NAME,
      5
    );

    return;

  }

  // Mechanic is available, set the ID
  mechanicIdCell.setValue(mechanicId);

}



/**
 * Automatically sets the Completion Date when Status changes to Completed.
 * Clears the Completion Date when Status changes from Completed to any other status.
 *
 * Work Orders:
 * Status = Completed -> Completion Date = Current Date
 * Status != Completed (was Completed) -> Completion Date = empty
 *
 * @param {Sheet} sheet - The sheet being edited.
 * @param {number} row - The row being edited.
 * @param {Object} config - The module configuration.
 * @param {Object} event - The edit event object.
 */
function syncWorkOrderCompletionDate(sheet, row, config, event) {

  const statusColumn = config.fields.Status;
  const completionDateColumn = config.fields.CompletionDate;

  // Only act if the Status column was edited
  if (event.range.getColumn() !== statusColumn) {
    return;
  }

  const newStatus = String(event.value || "").trim();
  const oldStatus = String(event.oldValue || "").trim();
  const completionDateCell = sheet.getRange(row, completionDateColumn);

  // If status changed to Completed, set completion date to today
  if (newStatus === "Completed") {
    completionDateCell.setValue(
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"
      )
    );
  }
  // If status changed FROM Completed to something else, clear completion date
  else if (oldStatus === "Completed" && newStatus !== "Completed") {
    completionDateCell.clearContent();
  }

}


/**
 * Synchronizes mechanic status when a Work Order is assigned, changed, or completed.
 *
 * @param {Sheet} sheet - The sheet being edited.
 * @param {number} row - The row being edited.
 * @param {Object} config - The module configuration.
 * @param {Object} event - The edit event object.
 */
function syncWorkOrderMechanicAssignment(sheet, row, config, event) {

  const mechanicNameColumn = config.fields.MechanicName;
  const mechanicIdColumn = config.fields.MechanicID;
  const statusColumn = config.fields.Status;

  const editedColumn = event.range.getColumn();

  // Check if mechanic was changed or status was changed
  const mechanicChanged = editedColumn === mechanicNameColumn || editedColumn === mechanicIdColumn;
  const statusChanged = editedColumn === statusColumn;

  if (!mechanicChanged && !statusChanged) {
    return;
  }

  const mechanicId = sheet
    .getRange(row, mechanicIdColumn)
    .getDisplayValue()
    .trim();

  if (!mechanicId) {
    return;
  }

  // Sync the mechanic's status based on their active work orders
  MechanicAssignmentService.syncMechanicStatus(mechanicId);

  // Handle reassignment: if mechanic changed, also sync the old mechanic
  if (mechanicChanged && event.oldValue) {

    const oldMechanicId = findMechanicIdByName(String(event.oldValue).trim());

    if (oldMechanicId && oldMechanicId !== mechanicId) {
      MechanicAssignmentService.syncMechanicStatus(oldMechanicId);
    }

  }

}


function runWorkOrderBusinessValidations(sheet, row, config, event) {

  // Validate mechanic selection (Feature 5)
  validateMechanicSelection(sheet, row, config, event);

  // Sync completion date (Feature 1)
  syncWorkOrderCompletionDate(sheet, row, config, event);

  // Sync mechanic assignment (Feature 4)
  syncWorkOrderMechanicAssignment(sheet, row, config, event);

}



/**
 * Automatically creates the Open Date when
 * every required foreign key has been assigned.
 *
 * Work Orders:
 * Customer ID
 * Vehicle ID
 * Mechanic ID
 * ->
 * Open Date
 */
function syncWorkOrderOpenDate(sheet, row) {

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  const customerId = sheet
    .getRange(row, config.fields.CustomerID)
    .getDisplayValue()
    .trim();

  const vehicleId = sheet
    .getRange(row, config.fields.VehicleID)
    .getDisplayValue()
    .trim();

  const mechanicId = sheet
    .getRange(row, config.fields.MechanicID)
    .getDisplayValue()
    .trim();

  const openDateCell = sheet.getRange(
    row,
    config.fields.OpenDate
  );

  const currentOpenDate = openDateCell
    .getDisplayValue()
    .trim();

  // Never overwrite an existing date.
  if (currentOpenDate !== "") {
    return;
  }

  // Wait until every foreign key exists.
  if (
    customerId === "" ||
    vehicleId === "" ||
    mechanicId === ""
  ) {
    return;
  }

  openDateCell.setValue(
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MM/dd/yyyy"
    )
  );

}




