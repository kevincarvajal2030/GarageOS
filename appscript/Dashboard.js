/**
 * ============================================================================
 * GARAGE OS
 * Dashboard Module
 * ============================================================================
 *
 * Handles dashboard calculations and updates.
 *
 * Responsibilities:
 *   - Update dashboard.
 *   - Calculate KPIs.
 *   - Refresh summary information.
 *
 * ============================================================================
 */


/******************************************************************************
 * DASHBOARD
 ******************************************************************************/

/**
 * Refreshes the entire dashboard.
 */
function refreshDashboard() {

  updateCustomerCount();

  updateVehicleCount();

  updateWorkOrderCount();

  updateInventoryCount();

  updatePaymentCount();

}


/******************************************************************************
 * CUSTOMERS
 ******************************************************************************/

/**
 * Updates total customers.
 */
function updateCustomerCount() {

  const totalCustomers =
    Math.max(
      getLastRow(SHEETS.CUSTOMERS) - TABLE.HEADER_ROW,
      0
    );

  getSheet(SHEETS.DASHBOARD)
    .getRange("B2")
    .setValue(totalCustomers);

}


/******************************************************************************
 * VEHICLES
 ******************************************************************************/

/**
 * Updates total vehicles.
 */
function updateVehicleCount() {

  const totalVehicles =
    Math.max(
      getLastRow(SHEETS.VEHICLES) - TABLE.HEADER_ROW,
      0
    );

  getSheet(SHEETS.DASHBOARD)
    .getRange("B3")
    .setValue(totalVehicles);

}


/******************************************************************************
 * WORK ORDERS
 ******************************************************************************/

/**
 * Updates total work orders.
 */
function updateWorkOrderCount() {

  const totalWorkOrders =
    Math.max(
      getLastRow(SHEETS.WORK_ORDERS) - TABLE.HEADER_ROW,
      0
    );

  getSheet(SHEETS.DASHBOARD)
    .getRange("B4")
    .setValue(totalWorkOrders);

}


/******************************************************************************
 * INVENTORY
 ******************************************************************************/

/**
 * Updates total inventory items.
 */
function updateInventoryCount() {

  const totalParts =
    Math.max(
      getLastRow(SHEETS.PARTS) - TABLE.HEADER_ROW,
      0
    );

  getSheet(SHEETS.DASHBOARD)
    .getRange("B5")
    .setValue(totalParts);

}


/******************************************************************************
 * PAYMENTS
 ******************************************************************************/

/**
 * Updates total payments.
 */
function updatePaymentCount() {

  const totalPayments =
    Math.max(
      getLastRow(SHEETS.PAYMENTS) - TABLE.HEADER_ROW,
      0
    );

  getSheet(SHEETS.DASHBOARD)
    .getRange("B6")
    .setValue(totalPayments);

}


/******************************************************************************
 * DASHBOARD ACTIONS
 ******************************************************************************/

/**
 * Refreshes all dashboard information.
 */
function updateDashboard() {

  refreshDashboard();

  showToast(
    "Dashboard updated successfully.",
    APP.NAME
  );

}
