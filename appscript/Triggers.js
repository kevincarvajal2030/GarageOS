function onEdit(event) {

  restoreProtectedId(event);

  AutoRowEngine.processEdit(event);

}


/**
 * Simple trigger: fires automatically whenever the user's selection
 * changes on any sheet. Writes a cheap signature (sheet name + row) to
 * the user's cache so the Vehicle Viewer sidebar can poll a near-free
 * function (getVehicleViewerSelectionSignature) instead of re-running
 * the full Drive/image lookup every 900ms regardless of clicks.
 *
 * IMPORTANT: simple triggers must never throw, or Sheets will silently
 * disable them. Everything here is wrapped defensively.
 */
function onSelectionChange(event) {

  try {

    if (!event || !event.range) return;

    const sheet = event.range.getSheet();
    const row = event.range.getRow();

    const signature = sheet.getName() + "|" + row;

    CacheService.getUserCache().put("vehicleViewerSelection", signature, 120);

  } catch (err) {
    // Swallow. A broken cache write should never break normal editing.
  }

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