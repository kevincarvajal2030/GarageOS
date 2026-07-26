/**
 * TRIGGERS PRINCIPALES PARA GARAGEOS
 * Archivo: Triggers.gs
 */

// Constantes de debounce para evitar ejecuciones masivas
const DEBOUNCE_MS = 800;

/**
 * ON EDIT: Restaura funcionalidad de AutoRowEngine y maneja actualizaciones del Viewer
 */
function onEdit(e) {
  // 1. SEGURIDAD: Si no hay evento válido, salir
  if (!e || !e.range || !e.source) return;

  // 2. CRÍTICO: Ejecutar primero el motor de validaciones y creación de IDs
  // Esto restaura la generación de CustomerID, validaciones de email/teléfono, etc.
  try {
    AutoRowEngine.processEdit(e);
  } catch (err) {
    Logger.log("Error en AutoRowEngine: " + err.toString());
  }

  // 3. Restaurar IDs protegidos si el usuario intenta editarlos
  try {
    restoreProtectedId(e);
  } catch (err) {
    // Ignorar errores silenciosos aquí para no bloquear la edición normal
  }

  // 4. Lógica específica para el Vehicle Viewer
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Solo nos interesa si es Vehicles o Work Orders
  if (sheetName !== SHEETS.VEHICLES && sheetName !== SHEETS.WORK_ORDERS) {

    return;
  }

  const row = e.range.getRow();
  if (row < TABLE.FIRST_DATA_ROW) return;

  // DEBOUNCE: Evitar múltiples disparos por un solo cambio
  const cache = CacheService.getUserCache();
  const now = new Date().getTime();
  const cacheKey = 'last_edit_' + sheetName;
  const lastEdit = cache.get(cacheKey);

  if (lastEdit && (now - parseInt(lastEdit)) < DEBOUNCE_MS) {
    return;
  }
  cache.put(cacheKey, now.toString(), 10);

  // Forzar actualización del Viewer
  const col = e.range.getColumn();
  const signature = sheetName + "|" + row + "|" + col;

  cache.put('vehicleViewerSelection', signature, 60);
  cache.put('vehicleViewerTimestamp', now.toString(), 60);

  Logger.log("EDIT TRIGGER (Viewer): " + signature);
}

/**
 * ON SELECTION CHANGE: Maneja el cambio de fila para el Viewer
 * CORRECCIÓN CRÍTICA: Validar estrictamente que sea UNA sola celda
 */
function onSelectionChange(e) {
  if (!e || !e.source) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Solo actuar en hojas relevantes
  if (sheetName !== SHEETS.VEHICLES && sheetName !== SHEETS.WORK_ORDERS) {
    return;
  }

  const range = e.range;

  // --- SOLUCIÓN AL ERROR "Número de columnas debe ser 1" ---
  // Si el usuario seleccionó más de una celda (ej. toda la fila o arrastró),
  // NO hacemos nada. El viewer solo funciona con selección de celda única.
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    return;
  }
  // ---------------------------------------------------------

  const row = range.getRow();
  if (row < TABLE.FIRST_DATA_ROW) return;

  // DEBOUNCE para selección (más rápido que edit)
  const cache = CacheService.getUserCache();
  const now = new Date().getTime();
  const cacheKey = 'last_select_' + sheetName;
  const lastSelect = cache.get(cacheKey);

  if (lastSelect && (now - parseInt(lastSelect)) < 400) {
    return;
  }
  cache.put(cacheKey, now.toString(), 10);

  const col = range.getColumn();
  const signature = sheetName + "|" + row + "|" + col;

  cache.put('vehicleViewerSelection', signature, 60);
  cache.put('vehicleViewerTimestamp', now.toString(), 60);

  // Logger.log("SELECTION TRIGGER: " + signature);
}

/**
 * Limpieza de triggers antiguos (opcional)
 */
function setupVehicleViewerTrigger() {

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    if (trigger.getHandlerFunction() === 'forceVehicleViewerRefresh') {

      ScriptApp.deleteTrigger(trigger);
    }
  });
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


  // Module-specific protected fields (only protect after ID exists)
  if (config.protectedFields && config.protectedFields.length > 0) {

    // Check if record already has an ID (for modules with auto-generated IDs)
    const idFieldName = getIdFieldName(config);
    let hasId = false;

    if (idFieldName) {
      const idColumn = config.fields[idFieldName];
      const idValue = sheet.getRange(event.range.getRow(), idColumn).getDisplayValue().trim();
      hasId = isValidRecordId(idValue, config);
    }

    // Only enforce protected fields if the record already has an ID
    if (hasId || !config.autoGenerateId) {
      config.protectedFields.forEach(field => {

        protectedFields.push({

          name: field,

          column: config.fields[field]

        });

      });
    }

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


/**
 * Helper to get the ID field name from config.
 */
function getIdFieldName(config) {

  return Object.keys(config.fields).find(name => name.endsWith("ID")) || null;

}


/**
 * Helper to validate if an ID is valid (has the correct prefix format).
 */
function isValidRecordId(id, config) {

  return !!id && id.startsWith(config.idPrefix + "-");

}


