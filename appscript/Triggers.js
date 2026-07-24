//Triggers.gs

function onEdit(e) {
  if (!e || !e.range || !e.value) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName !== SHEETS.VEHICLES && sheetName !== SHEETS.WORK_ORDERS) {
    return;
  }

  if (e.range.getRow() < TABLE.FIRST_DATA_ROW) return;

  let shouldUpdate = false;

  if (sheetName === SHEETS.WORK_ORDERS) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const vehicleColIndex = headers.indexOf("Vehicle") + 1;
    const customerColIndex = headers.indexOf("Customer") + 1;

    if (e.range.getColumn() === vehicleColIndex || e.range.getColumn() === customerColIndex) {
      shouldUpdate = true;
    }
  }
  else if (sheetName === SHEETS.VEHICLES) {
    shouldUpdate = true;
  }

  if (shouldUpdate) {
    const range = e.range;
    const signature = sheetName + "|" + range.getRow() + "|" + range.getColumn();

    const cache = CacheService.getUserCache();
    const now = new Date().getTime().toString();

    cache.put('vehicleViewerSelection', signature, 60);
    cache.put('vehicleViewerTimestamp', now, 60);

    Logger.log("EDIT DETECTADO: Force update triggered for " + signature);
  }
}


/**
 * Configura los triggers para el Vehicle Viewer.
 * NOTA: Los triggers simples (onEdit, onSelectionChange) no necesitan instalación,
 * se ejecutan automáticamente si están en el proyecto con ese nombre exacto.
 * Esta función ahora solo sirve para limpiar triggers antiguos si los hubiera.
 */
function setupVehicleViewerTrigger() {
  // Limpiamos triggers instalables antiguos si existieran para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'forceVehicleViewerRefresh') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log("Triggers limpiados. Los eventos onEdit y onSelectionChange son nativos.");
}



/**
 * SE DISPARA AL MOVER EL CURSOR (Cambio de fila)
 * Funciona para: Cambiar de Work Order o de Vehicle Row.
 */
function onSelectionChange(e) {
  if (!e || !e.source) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName !== SHEETS.VEHICLES && sheetName !== SHEETS.WORK_ORDERS) {
    return;
  }

  const range = e.range;
  if (range.getRow() < TABLE.FIRST_DATA_ROW) return;

  const signature = sheetName + "|" + range.getRow() + "|" + range.getColumn();
  
  const cache = CacheService.getUserCache();
  const now = new Date().getTime().toString();
  
  cache.put('vehicleViewerSelection', signature, 60);
  cache.put('vehicleViewerTimestamp', now, 60);
  
  Logger.log("SELECCIÓN CAMBIADA: " + signature);
}



/**
 * Helper opcional para obtener índices de columnas si no los tienes globales
 * (Ajusta esto según tu ModuleConfig.gs real)
 */
function getWorkOrderColumnIndex_(headerName) {
  // Implementación simplificada: busca el header en la fila 1
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.WORK_ORDERS);
  if (!sheet) return -1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName) + 1;
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