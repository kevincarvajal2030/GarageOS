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

/**
 * Returns the spreadsheet row that matches an ID.
 *
 * @param {string} sheetName
 * @param {string} id
 * @returns {number}
 */
function findRowById(sheetName, id) {

  const data = getData(sheetName);

  for (let row = TABLE.FIRST_DATA_ROW - 1; row < data.length; row++) {

    if (data[row][0] === id) {

      return row + 1;

    }

  }

  return -1;

}

/**
 * Determines whether a record exists.
 *
 * @param {string} sheetName
 * @param {string} id
 * @returns {boolean}
 */
function recordExists(sheetName, id) {

  return findRowById(sheetName, id) !== -1;

}


/**
 * Returns an entire spreadsheet row.
 *
 * @param {string} sheetName
 * @param {string} id
 * @returns {Array|null}
 */
function getRecord(sheetName, id) {

  const row = findRowById(sheetName, id);

  if (row === -1) {

    return null;

  }

  return getSheet(sheetName)
    .getRange(
      row,
      1,
      1,
      getLastColumn(sheetName)
    )
    .getValues()[0];

}


/**
 * Replaces a complete record.
 *
 * @param {string} sheetName
 * @param {string} id
 * @param {Array} record
 */
function updateRecord(sheetName, id, record) {

  const row = findRowById(sheetName, id);

  if (row === -1) {

    throw new Error("Record not found.");

  }

  getSheet(sheetName)
    .getRange(
      row,
      1,
      1,
      record.length
    )
    .setValues([record]);

}


/**
 * Deletes a record.
 *
 * @param {string} sheetName
 * @param {string} id
 */
function deleteRecord(sheetName, id) {

  const row = findRowById(sheetName, id);

  if (row === -1) {

    throw new Error("Record not found.");

  }

  getSheet(sheetName)
    .deleteRow(row);

}


/**
 * Appends a complete record.
 *
 * @param {string} sheetName
 * @param {Array} record
 */
function appendRecord(sheetName, record) {

  appendRow(
    sheetName,
    record
  );

}
