/**
 * ============================================================================
 * GARAGE OS
 * Setup Module
 * ============================================================================
 *
 * Handles the initial application setup.
 *
 * Responsibilities:
 *   - Verify required sheets.
 *   - Initialize the application.
 *   - Reset demo data.
 *   - Refresh dashboard.
 *
 * ============================================================================
 */


/******************************************************************************
 * SETUP
 ******************************************************************************/

/**
 * Runs the initial GarageOS setup.
 */
function runSetup() {

  verifySheets();

  showToast(
    "GarageOS is ready.",
    APP.NAME
  );

}


/******************************************************************************
 * SHEET VALIDATION
 ******************************************************************************/

/**
 * Verifies that all required sheets exist.
 */
function verifySheets() {

  const requiredSheets = [

    SHEETS.DASHBOARD,

    SHEETS.CUSTOMERS,

    SHEETS.VEHICLES,

    SHEETS.WORK_ORDERS,

    SHEETS.SERVICES,

    SHEETS.PARTS,

    SHEETS.SUPPLIERS,

    SHEETS.PURCHASES,

    SHEETS.PAYMENTS,

    SHEETS.MECHANICS,

    SHEETS.REFERENCEDATA,

    SHEETS.SYSTEM

  ];

  requiredSheets.forEach(function(sheetName) {

    if (!getSheet(sheetName)) {

      throw new Error(
        "Missing sheet: " + sheetName
      );

    }

  });

}



/******************************************************************************
 * REFRESH
 ******************************************************************************/

/**
 * Refreshes the application.
 */
function refreshApplication() {

  updateDashboard();

  showToast(
    "Application refreshed.",
    APP.NAME
  );

}


/******************************************************************************
 * ABOUT
 ******************************************************************************/

/**
 * Displays application information.
 */
function aboutGarageOS() {

  SpreadsheetApp.getUi().alert(

    APP.NAME +
    "\n\nVersion: " +
    APP.VERSION

  );

}




