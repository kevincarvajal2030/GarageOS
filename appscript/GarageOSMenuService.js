
/**
 * ============================================================================
 * GarageOS
 * GarageOSMenuService.gs
 * ============================================================================
 *
 * Server-side functions for the GarageOS navigation sidebar.
 *
 */

/******************************************************************************
 * PREFERENCES - Auto Open Vehicle Viewer
 ******************************************************************************/

const PREF_KEYS = {
  AUTO_OPEN_VEHICLE_VIEWER: "AUTO_OPEN_VEHICLE_VIEWER"
};

/**
 * Gets the auto-open Vehicle Viewer preference.
 * @returns {boolean} True if auto-open is enabled.
 */
function getAutoOpenVehicleViewerPreference() {

  const props = PropertiesService.getUserProperties();
  const value = props.getProperty(PREF_KEYS.AUTO_OPEN_VEHICLE_VIEWER);

  return value === "true";

}

/**
 * Sets the auto-open Vehicle Viewer preference.
 * @param {boolean} enabled - Whether auto-open should be enabled.
 */
function setAutoOpenVehicleViewerPreference(enabled) {

  const props = PropertiesService.getUserProperties();
  props.setProperty(PREF_KEYS.AUTO_OPEN_VEHICLE_VIEWER, enabled ? "true" : "false");

}


/******************************************************************************
 * DASHBOARD
 ******************************************************************************/

/**
 * Opens the Dashboard by activating the sheet and refreshing data.
 * Does NOT rebuild the layout - only refreshes existing dashboard data.
 */
function showDashboardSidebar() {

  // Activate the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("00_Dashboard");

  if (sheet) {
    ss.setActiveSheet(sheet);
    // Refresh dashboard data without rebuilding layout
    refreshDashboardData();

  } else {
    // If dashboard sheet doesn't exist, build it first
    buildDashboard();

    const newSheet = ss.getSheetByName("00_Dashboard");

    if (newSheet) {
      ss.setActiveSheet(newSheet);
    }

  }

}

/******************************************************************************
 * NAVIGATION
 ******************************************************************************/

/**
 * Activates a sheet by name.
 * @param {string} sheetName - The name of the sheet to activate.
 */
function activateSheetByName(sheetName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    throw new Error("Sheet not found: " + sheetName);
  }

}

/**
 * Gets the name of the currently active sheet.
 * @returns {string} The active sheet name.
 */
function getActiveSheetName() {

  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

}


/******************************************************************************
 * APP INFO
 ******************************************************************************/

/**
 * Gets the application version.
 * @returns {string} The version string.
 */
function getAppVersion() {

  return APP.VERSION || "1.0.1";

}