/**
 * Database.gs
 * Returns the active spreadsheet.
 */
function getDatabase() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


/**
 * Returns a sheet by name.
 *
 * @param {string} sheetName
 * @returns {Sheet}
 */
function getSheet(sheetName) {

  return getDatabase().getSheetByName(sheetName);

}

/**
 * Returns the last row containing data.
 *
 * @param {string} sheetName
 * @returns {number}
 */
function getLastRow(sheetName) {

  return getSheet(sheetName).getLastRow();

}

/**
 * Returns the last column containing data.
 *
 * @param {string} sheetName
 * @returns {number}
 */
function getLastColumn(sheetName) {

  return getSheet(sheetName).getLastColumn();

}

