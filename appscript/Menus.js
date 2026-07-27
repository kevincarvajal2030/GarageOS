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
      "Toggle Navigation",
      "toggleGarageOSSidebar"
    )

    .addToUi();

}


/******************************************************************************
 * SIDEBAR TOGGLE
 ******************************************************************************/

let sidebarIsOpen = false;

function toggleGarageOSSidebar() {

  if (sidebarIsOpen) {
    closeGarageOSSidebar();
  } else {
    openGarageOSSidebar();
  }

}

function openGarageOSSidebar() {

  const html = HtmlService
    .createHtmlOutputFromFile("GarageOSMenu")
    .setTitle("GarageOS Menu");

  SpreadsheetApp.getUi().showSidebar(html);
  sidebarIsOpen = true;

}

function closeGarageOSSidebar() {

  SpreadsheetApp.getUi().closeSidebar();
  sidebarIsOpen = false;

}