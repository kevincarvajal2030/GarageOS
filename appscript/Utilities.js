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


/**
 * Copies all formatting from source range to target range.
 * Reusable utility for synchronizing cell appearance across modules.
 *
 * @param {Range} sourceRange - The source range to copy formatting from.
 * @param {Range} targetRange - The target range to apply formatting to.
 */
function copyCellFormatting(sourceRange, targetRange) {

  // Copy background color
  const background = sourceRange.getBackground();
  targetRange.setBackground(background);

  // Copy font color
  const fontColor = sourceRange.getFontColor();
  targetRange.setFontColor(fontColor);

  // Copy font weight (bold or not)
  const fontWeight = sourceRange.getFontWeight();
  targetRange.setFontWeight(fontWeight);

  // Copy font style (italic or not)
  const fontStyle = sourceRange.getFontStyle();
  targetRange.setFontStyle(fontStyle);

  // Copy horizontal alignment
  const horizontalAlignment = sourceRange.getHorizontalAlignment();
  targetRange.setHorizontalAlignment(horizontalAlignment);

  // Copy vertical alignment
  const verticalAlignment = sourceRange.getVerticalAlignment();
  targetRange.setVerticalAlignment(verticalAlignment);

}


/**
 * Applies Payment Status formatting to a cell.
 * Reusable utility for consistent Payment Status visual appearance.
 *
 * @param {Range} cell - The cell to format.
 * @param {string} status - The payment status value.
 */
function applyPaymentStatusFormatting(cell, status) {

  // Reset to default first
  cell.setBackground("#ffffff");
  cell.setFontColor("#000000");
  cell.setFontWeight("normal");

  switch (status) {

    case "Pending":
      cell.setBackground("#ffe5a0"); // Yellow
      cell.setFontColor("#473821");
      break;

    case "Paid":
      cell.setBackground("#d4edbc"); // Green
      cell.setFontColor("#11734b");
      break;

    case "Refunded":
      cell.setBackground("#ffcfc9"); // Red
      cell.setFontColor("#b10202");
      break;

    default:
      // Keep default formatting for other statuses
      break;

  }

}