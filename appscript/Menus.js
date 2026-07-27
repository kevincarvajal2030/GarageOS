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

  // Initialize Dashboard layout verification and data refresh
  initializeDashboard();
}


function initializeDashboard() {
  try {
    // Verify dashboard layout exists, build missing sections only
    verifyDashboardLayout();

    // Refresh dashboard data (formulas will auto-recalculate)
    // This is lightweight since KPIs use formulas that auto-update
    const sheet = getSheet(SHEETS.DASHBOARD);
    if (sheet) {
      // Update timestamp
      const lastUpdated = sheet.getRange("I1");
      lastUpdated.setFormula('=TEXT(NOW(),"MM/dd/yyyy HH:mm")');
    }
  } catch (e) {
    // Silently fail - dashboard will be built when user navigates to it
    console.log("Dashboard initialization deferred: " + e.message);
  }
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