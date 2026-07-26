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
 * Validates mechanic status changes to prevent changing from Busy to Available/Vacation/Inactive
 * when the mechanic has active Work Orders.
 *
 * @param {Sheet} sheet - The sheet being edited.
 * @param {number} row - The row being edited.
 * @param {Object} config - The module configuration.
 * @param {Object} event - The edit event object.
 */
function validateMechanicStatusChange(sheet, row, config, event) {

  const statusColumn = config.fields.Status;
  const mechanicIdColumn = config.fields.MechanicID;

  // Only act if the Status column was edited
  if (event.range.getColumn() !== statusColumn) {
    return;
  }

  const previousStatus = String(event.oldValue || "").trim();
  const newStatus = String(event.value || "").trim();

  // Only validate when changing FROM Busy TO Available, Vacation, or Inactive
  if (previousStatus !== "Busy") {
    return;
  }

  if (
    newStatus !== "Available" &&
    newStatus !== "Vacation" &&
    newStatus !== "Inactive"
  ) {
    return;
  }

  const mechanicId = sheet
    .getRange(row, mechanicIdColumn)
    .getDisplayValue()
    .trim();

  if (!mechanicId) {
    return;
  }

  // Check if mechanic has active Work Orders
  const hasActive = MechanicAssignmentService.hasActiveWorkOrders(mechanicId);

  if (hasActive) {

    // Restore previous value
    event.range.setValue(previousStatus);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "This mechanic has active work orders and their status cannot be changed.",
      APP.NAME,
      5
    );

    return;

  }

}


function runMechanicBusinessValidations(sheet, row, config, event) {

  validateMechanicStatusChange(sheet, row, config, event);

}


