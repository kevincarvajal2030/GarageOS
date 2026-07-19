/**
 * ============================================================================
 * GarageOS
 * Menus.gs
 * ============================================================================
 *
 * Creates the GarageOS custom menu.
 *
 * This menu only contains administrative actions.
 * All CRUD operations are performed directly
 * from the Google Sheets interface.
 * ============================================================================
 */


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
      "Run Setup",
      "runSetup"
    )

    .addItem(
      "About GarageOS",
      "aboutGarageOS"
    )

    .addToUi();

}