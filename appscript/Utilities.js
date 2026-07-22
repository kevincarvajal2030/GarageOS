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


function debug(label, value) {

  console.log(
    JSON.stringify({
      step: label,
      value: value
    })
  );

}