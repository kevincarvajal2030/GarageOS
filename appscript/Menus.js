/**
 * ============================================================================
 * GarageOS
 * Menus.gs
 * ============================================================================
 *
 * Creates the GarageOS custom menu.
 * 
/******************************************************************************
 * ON OPEN
 ******************************************************************************/

function onOpen() {

  createGarageOSMenu();

}


/******************************************************************************
 * MAIN MENU
 ******************************************************************************/

function createGarageOSMenu() {

  SpreadsheetApp.getUi()

    .createMenu(APP.NAME)

    .addItem(
      "Refresh Dashboard",
      "updateDashboard"
    )

    .addSeparator()

    .addItem(
      "Vehicles Viewer",
      "openVehicleViewerSidebar"
    )

    .addItem(
      "About GarageOS",
      "aboutGarageOS"
    )

    .addToUi();

}