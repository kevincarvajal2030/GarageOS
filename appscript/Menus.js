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

    .createMenu("☰ GarageOS")

    .addItem(
      "Open Navigation",
      "openGarageOSSidebar"
    )

    .addToUi();

}


/******************************************************************************
 * SIDEBAR MANAGEMENT
 ******************************************************************************/

function openGarageOSSidebar() {

  const html = HtmlService
    .createHtmlOutputFromFile("GarageOSMenu")
    .setTitle("GarageOS Menu");

  SpreadsheetApp.getUi().showSidebar(html);

}

function closeGarageOSSidebar() {

  SpreadsheetApp.getUi().closeSidebar();

}