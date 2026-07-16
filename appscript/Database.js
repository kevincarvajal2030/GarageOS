/**
 * Database.gs
 * Returns the active spreadsheet.
 *
 * @returns {Spreadsheet}
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

/**
 * Returns all data from a sheet.
 *
 * @param {string} sheetName
 * @returns {Array}
 */
function getData(sheetName) {

  return getSheet(sheetName)
    .getDataRange()
    .getValues();

}


/**
 * Appends one row.
 *
 * @param {string} sheetName
 * @param {Array} row
 */
function appendRow(sheetName, row) {

  getSheet(sheetName)
    .appendRow(row);

}


/**
 * Clears all records but preserves headers.
 *
 * @param {string} sheetName
 */
function clearTable(sheetName) {

  const sheet = getSheet(sheetName);

  const lastRow = sheet.getLastRow();

  if (lastRow >= TABLE.FIRST_DATA_ROW) {

    sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    ).clearContent();

  }

}


