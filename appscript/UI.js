/**
 * ============================================================================
 * GARAGE OS
 * UI Module
 * ============================================================================
 *
 * Handles all user interface actions.
 *
 * Responsibilities:
 *   - Toast messages.
 *   - Alerts.
 *   - Confirmations.
 *   - Placeholder dialogs.
 * ============================================================================
 */


/**
 * Opens the Edit Customer screen.
 */
function editCustomer() {

  SpreadsheetApp.getUi().alert(
    "Select a customer first."
  );

}


/**
 * Deletes a customer.
 */
function removeCustomer() {

  SpreadsheetApp.getUi().alert(
    "Select a customer first."
  );

}


/******************************************************************************
 * VEHICLES
 ******************************************************************************/

/**
 * Opens the New Vehicle screen.
 */
function newVehicle() {

  showAlert("New Vehicle - Coming Soon");

}


/**
 * Opens the Edit Vehicle screen.
 */
function editVehicle() {

  showAlert("Edit Vehicle - Coming Soon");

}


/**
 * Deletes a vehicle.
 */
function removeVehicle() {

  showAlert("Delete Vehicle - Coming Soon");

}


/******************************************************************************
 * WORK ORDERS
 ******************************************************************************/

/**
 * Opens the New Work Order screen.
 */
function newWorkOrder() {

  showAlert("New Work Order - Coming Soon");

}


/**
 * Opens the Edit Work Order screen.
 */
function editWorkOrder() {

  showAlert("Edit Work Order - Coming Soon");

}


/**
 * Deletes a work order.
 */
function removeWorkOrder() {

  showAlert("Delete Work Order - Coming Soon");

}


/******************************************************************************
 * INVENTORY
 ******************************************************************************/

/**
 * Opens the New Part screen.
 */
function newPart() {

  showAlert("New Part - Coming Soon");

}


/**
 * Opens the Edit Part screen.
 */
function editPart() {

  showAlert("Edit Part - Coming Soon");

}


/**
 * Deletes a part.
 */
function removePart() {

  showAlert("Delete Part - Coming Soon");

}


/******************************************************************************
 * PAYMENTS
 ******************************************************************************/

/**
 * Opens the New Payment screen.
 */
function newPayment() {

  showAlert("New Payment - Coming Soon");

}


/**
 * Opens the Edit Payment screen.
 */
function editPayment() {

  showAlert("Edit Payment - Coming Soon");

}


/**
 * Deletes a payment.
 */
function removePayment() {

  showAlert("Delete Payment - Coming Soon");

}


/******************************************************************************
 * GENERAL UI
 ******************************************************************************/

/**
 * Displays a success message.
 *
 * @param {string} message
 */
function showSuccess(message) {

  showToast(
    message,
    "Success"
  );

}


/**
 * Displays an error message.
 *
 * @param {string} message
 */
function showError(message) {

  showAlert(message);

}


/**
 * Displays an information message.
 *
 * @param {string} message
 */
function showInfo(message) {

  showToast(
    message,
    APP.NAME
  );

}