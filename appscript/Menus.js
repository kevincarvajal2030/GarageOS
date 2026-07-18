/**
 * ============================================================================
 * GARAGE OS
 * Menus Module
 * ============================================================================
 *
 * Creates all custom menus for the application.
 * ============================================================================
 */


/******************************************************************************
 * ON OPEN
 ******************************************************************************/

/**
 * Creates the GarageOS menu when the spreadsheet opens.
 */
function onOpen() {

  createGarageOSMenu();

}


/******************************************************************************
 * MAIN MENU
 ******************************************************************************/

/**
 * Creates the main GarageOS menu.
 */
function createGarageOSMenu() {

  SpreadsheetApp.getUi()

    .createMenu(APP.NAME)

    .addItem(
      "Refresh Dashboard",
      "updateDashboard"
    )

    .addSeparator()

    .addSubMenu(
      createCustomersMenu()
    )

    .addSubMenu(
      createVehiclesMenu()
    )

    .addSubMenu(
      createWorkOrdersMenu()
    )

    .addSubMenu(
      createInventoryMenu()
    )

    .addSubMenu(
      createPaymentsMenu()
    )

    .addSeparator()

    .addItem(
      "Setup",
      "runSetup"
    )

    .addToUi();

}


/******************************************************************************
 * CUSTOMERS MENU
 ******************************************************************************/

function createCustomersMenu() {

  return SpreadsheetApp.getUi()

    .createMenu("Customers")

    .addItem(
      "New Customer",
      "newCustomer"
    )

    .addItem(
      "Edit Customer",
      "editCustomer"
    )

    .addItem(
      "Delete Customer",
      "removeCustomer"
    );

}


/******************************************************************************
 * VEHICLES MENU
 ******************************************************************************/

function createVehiclesMenu() {

  return SpreadsheetApp.getUi()

    .createMenu("Vehicles")

    .addItem(
      "New Vehicle",
      "newVehicle"
    )

    .addItem(
      "Edit Vehicle",
      "editVehicle"
    )

    .addItem(
      "Delete Vehicle",
      "removeVehicle"
    );

}


/******************************************************************************
 * WORK ORDERS MENU
 ******************************************************************************/

function createWorkOrdersMenu() {

  return SpreadsheetApp.getUi()

    .createMenu("Work Orders")

    .addItem(
      "New Work Order",
      "newWorkOrder"
    )

    .addItem(
      "Edit Work Order",
      "editWorkOrder"
    )

    .addItem(
      "Delete Work Order",
      "removeWorkOrder"
    );

}


/******************************************************************************
 * INVENTORY MENU
 ******************************************************************************/

function createInventoryMenu() {

  return SpreadsheetApp.getUi()

    .createMenu("Inventory")

    .addItem(
      "New Part",
      "newPart"
    )

    .addItem(
      "Edit Part",
      "editPart"
    )

    .addItem(
      "Delete Part",
      "removePart"
    );

}


/******************************************************************************
 * PAYMENTS MENU
 ******************************************************************************/

function createPaymentsMenu() {

  return SpreadsheetApp.getUi()

    .createMenu("Payments")

    .addItem(
      "New Payment",
      "newPayment"
    )

    .addItem(
      "Edit Payment",
      "editPayment"
    )

    .addItem(
      "Delete Payment",
      "removePayment"
    );

}
