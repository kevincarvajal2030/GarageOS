




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


