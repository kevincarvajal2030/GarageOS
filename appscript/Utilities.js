/**
 * Utilities.gs
 * Returns true if value is empty.
 *
 * @param {*} value
 * @returns {boolean}
 */
function isEmpty(value) {

  return value === "" ||
         value === null ||
         value === undefined;

}


/**
 * Returns true if value is not empty.
 *
 * @param {*} value
 * @returns {boolean}
 */
function isNotEmpty(value) {

  return !isEmpty(value);

}


/**
 * Returns current timestamp.
 *
 * @returns {Date}
 */
function getTimestamp() {

  return new Date();

}


/**
 * Formats a date.
 *
 * @param {Date} date
 * @param {string} format
 * @returns {string}
 */
function formatDate(date, format) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    format
  );

}


/**
 * Displays a toast message.
 *
 * @param {string} message
 * @param {string} title
 */
function showToast(message, title) {

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(message, title);

}


/**
 * Displays an alert dialog.
 *
 * @param {string} message
 */
function showAlert(message) {

  SpreadsheetApp
    .getUi()
    .alert(message);

}


/**
 * Capitalizes the first letter.
 *
 * @param {string} text
 * @returns {string}
 */
function capitalize(text) {

  if (isEmpty(text)) return "";

  return text.charAt(0).toUpperCase() +
         text.slice(1).toLowerCase();

}


/**
 * Formats a number as currency.
 *
 * @param {number} value
 * @returns {string}
 */
function toCurrency(value) {

  return "$" + Number(value).toFixed(2);

}


/**
 * Logs a message.
 *
 * @param {*} message
 */
function log(message) {

  Logger.log(message);

}


/**
 * Logs an error.
 *
 * @param {Error} error
 */
function logError(error) {

  Logger.log(error);

}





