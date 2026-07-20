function onEdit(event) {

  restoreProtectedId(event);

  AutoRowEngine.processEdit(event);

}


/**
 * Restores protected IDs if the user edits them manually.
 */
function restoreProtectedId(event) {

  if (!event || !event.range) return;

  const sheet = event.range.getSheet();

  const config = ModuleConfig.get(sheet.getName());

  if (!config) return;

  if (event.range.getRow() < TABLE.FIRST_DATA_ROW) return;

  if (
    event.range.getNumRows() > 1 ||
    event.range.getNumColumns() > 1
  ) {
    return;
  }

  const protectedFields = [];

  // Primary Key
  if (config.primaryKey) {

    protectedFields.push({

      name: config.primaryKey,

      column: config.fields[config.primaryKey]

    });

  }

  // Foreign Keys
  if (config.foreignKeys) {

    config.foreignKeys.forEach(field => {

      protectedFields.push({

        name: field,

        column: config.fields[field]

      });

    });

  }

  const edited = protectedFields.find(
    field => field.column === event.range.getColumn()
  );

  if (!edited) return;

  if (event.oldValue === undefined) return;

  event.range.setValue(event.oldValue);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    edited.name + " cannot be edited.",
    APP.NAME,
    3
  );

}