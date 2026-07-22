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
      "Vehicles Details", //Proximamente menu que mostrara la imagen del vehiculo que de la celda que selecciono el usuario.
      "runSetup"
    )

    .addItem(
      "About GarageOS",
      "aboutGarageOS"
    )

    .addToUi();

}