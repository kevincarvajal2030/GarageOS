 /**
 * Finds a Mechanic ID by Mechanic Name.
 *
 * @param {string} mechanicName
 * @returns {string|null}
 */
  function findMechanicIdByName(mechanicName) {

    const sheet = SpreadsheetApp
      .getActive()
      .getSheetByName(SHEETS.MECHANICS);

    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return null;
    }

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      3
    ).getValues();

    mechanicName = String(mechanicName).trim();

    for (const row of values) {

      const mechanicId = String(row[0]).trim();

      const firstName = String(row[1]).trim();

      const lastName = String(row[2]).trim();

      const fullName = firstName + " " + lastName;

      if (fullName === mechanicName) {
        return mechanicId;
      }

    }

    return null;

  }
