//Triggers.gs

/**
 * TRIGGERS PRINCIPALES DEL PROYECTO
 * Maneja: Vehicle Viewer, Validaciones de Clientes, Generación de IDs y Cascadas.
 */

function onEdit(e) {
  if (!e || !e.range || !e.value) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Ignorar encabezados
  if (row < TABLE.FIRST_DATA_ROW) return;

  // --- 1. GESTIÓN DE VALIDACIONES Y IDS (CLIENTES) ---
  // Esto asegura que las validaciones de correo/teléfono y generación de ID se activen
  if (sheetName === SHEETS.CUSTOMERS) {
    // Llamamos a las funciones originales de validación si existen
    // AJUSTA LOS NOMBRES DE ESTAS FUNCIONES SEGÚN TU PROYECTO REAL
    try {
      if (typeof validateCustomerEntry === 'function') {
        validateCustomerEntry(e);
      }
      if (typeof generateCustomerIdIfNeeded === 'function') {
        generateCustomerIdIfNeeded(e);
      }
      // Si tienes una función general de restauración de IDs
      if (typeof restoreProtectedId === 'function') {
        restoreProtectedId(e);
      }
    } catch (err) {
      Logger.log("Error en validaciones de cliente: " + err.toString());
    }
    
    // No hacemos return aquí, permitimos que continúe si hay más lógica
  }

  // --- 2. GESTIÓN DE WORK ORDERS (CASCADA Y VIEWER) ---
  if (sheetName === SHEETS.WORK_ORDERS) {
    const config = ModuleConfig.get(SHEETS.WORK_ORDERS);
    let vehicleColIndex = -1;
    let customerColIndex = -1;
    
    // Obtener índices dinámicamente desde ModuleConfig
    if (config && config.fields) {
      vehicleColIndex = config.fields['vehicleName'] || -1;
      customerColIndex = config.fields['customerName'] || -1;
    }
    
    // Si no se encontró por config, usamos valores por defecto (ajustar si es necesario)
    if (vehicleColIndex === -1) vehicleColIndex = 4; // Ejemplo default
    if (customerColIndex === -1) customerColIndex = 3; // Ejemplo default

    // A. Si cambia el CLIENTE -> Actualizar Dropdown de Vehículos
    if (col === customerColIndex) {
      Logger.log("Cliente cambiado en WO. Actualizando dropdowns...");
      Utilities.sleep(50); // Pequeña pausa para asegurar escritura
      updateVehicleDropdownByCustomer(row, e.value);
    }
    
    // B. Si cambia el VEHÍCULO -> Forzar refresh del Viewer
    if (col === vehicleColIndex || col === customerColIndex) {
      triggerVehicleViewerUpdate(sheetName, row, col);
    }
  }

  // --- 3. GESTIÓN DE VEHICLES (REFRESH DE IMAGEN) ---
  if (sheetName === SHEETS.VEHICLES) {
    // Cualquier edición en Vehicles debe refrescar la imagen si es la fila activa
    triggerVehicleViewerUpdate(sheetName, row, col);
  }
}


/**
 * Helper para disparar la actualización del Viewer sin bloquear el resto del script
 */
function triggerVehicleViewerUpdate(sheetName, row, col) {
  const cache = CacheService.getUserCache();
  const now = new Date().getTime();
  
  // Debounce ligero solo para el viewer (700ms)
  const lastEvent = cache.get('lastViewerEvent_' + sheetName);
  if (lastEvent && (now - parseInt(lastEvent)) < 700) {
    return; 
  }
  
  cache.put('lastViewerEvent_' + sheetName, now.toString(), 10);
  
  const signature = sheetName + "|" + row + "|" + col;
  cache.put('vehicleViewerSelection', signature, 60);
  cache.put('vehicleViewerTimestamp', now.toString(), 60);
  
  Logger.log("Viewer Triggered: " + signature);
}



function setupVehicleViewerTrigger() {
  // Limpieza de seguridad
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'forceVehicleViewerRefresh') ScriptApp.deleteTrigger(t);
  });
}



/**
 * SE DISPARA AL MOVER EL CURSOR (Cambio de fila)
 */
function onSelectionChange(e) {
  if (!e || !e.source) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName !== SHEETS.VEHICLES && sheetName !== SHEETS.WORK_ORDERS) return;
  
  const range = e.range;
  if (range.getRow() < TABLE.FIRST_DATA_ROW) return;
  
  // Debounce
  const cache = CacheService.getUserCache();
  const now = new Date().getTime();
  const lastEvent = cache.get('lastViewerEvent_' + sheetName);
  if (lastEvent && (now - parseInt(lastEvent)) < 400) return;
  
  cache.put('lastViewerEvent_' + sheetName, now.toString(), 10);
  
  const signature = sheetName + "|" + range.getRow() + "|" + range.getColumn();
  cache.put('vehicleViewerSelection', signature, 60);
  cache.put('vehicleViewerTimestamp', now.toString(), 60);
}


/**
 * Helper para obtener nombres de vehículos filtrados
 */
function getVehicleNamesForCustomer_(sheet, customerId, customerName) {
  if (!sheet) return [];
  
  const config = ModuleConfig.get(SHEETS.VEHICLES);
  if (!config) return [];
  
  const nameCol = config.fields['vehicleName']; // Nombre para mostrar
  const custIdCol = config.fields['customerId'];
  const custNameCol = config.fields['customerName'];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < TABLE.FIRST_DATA_ROW) return [];
  
  // Leer solo columnas necesarias
  const maxCol = Math.max(nameCol, custIdCol, custNameCol);
  const data = sheet.getRange(TABLE.FIRST_DATA_ROW, 1, lastRow - TABLE.FIRST_DATA_ROW + 1, maxCol).getValues();
  
  const result = [];
  
  for (let i = 0; i < data.length; i++) {
    const rowVehId = String(data[i][nameCol-1] || ""); // Usamos como ID temporal si no hay otro
    const rowCustId = String(data[i][custIdCol-1] || "").trim();
    const rowCustName = String(data[i][custNameCol-1] || "").trim();
    
    let match = false;
    if (customerId && rowCustId === customerId) match = true;
    else if (customerName && rowCustName === customerName) match = true;
    
    if (match && rowVehId !== "") {
      result.push(rowVehId);
    }
  }
  
  return result.sort(); // Ordenar alfabéticamente
}


/**
 * Actualiza la validación de datos (dropdown) de vehículos basado en el cliente
 */
function updateVehicleDropdownByCustomer(rowNumber, customerName) {
  try {
    const ss = SpreadsheetApp.getActive();
    const woSheet = ss.getSheetByName(SHEETS.WORK_ORDERS);
    const vehSheet = ss.getSheetByName(SHEETS.VEHICLES);
    
    if (!woSheet || !vehSheet) return;
    
    const configWO = ModuleConfig.get(SHEETS.WORK_ORDERS);
    const vehicleCol = configWO ? configWO.fields['vehicleName'] : 4;
    
    // 1. Buscar ID del cliente (asumiendo que customerName es único o usando lógica existente)
    // Usamos una búsqueda simple en la hoja de clientes
    const custSheet = ss.getSheetByName(SHEETS.CUSTOMERS);
    let targetCustomerId = "";
    
    if (custSheet) {
      const custConfig = ModuleConfig.get(SHEETS.CUSTOMERS);
      const nameCol = custConfig ? custConfig.fields['name'] : 2; // Ajustar según config
      const idCol = custConfig ? custConfig.fields['customerId'] : 1;
      
      const lastRow = custSheet.getLastRow();
      if (lastRow >= TABLE.FIRST_DATA_ROW) {
        const data = custSheet.getRange(TABLE.FIRST_DATA_ROW, 1, lastRow - TABLE.FIRST_DATA_ROW + 1, Math.max(nameCol, idCol)).getValues();
        for (let i = 0; i < data.length; i++) {
          // Comparación segura
          if (String(data[i][nameCol-1]).trim() === String(customerName).trim()) {
            targetCustomerId = String(data[i][idCol-1]);
            break;
          }
        }
      }
    }
    
    // 2. Obtener lista de vehículos de ese cliente
    let vehicleNames = [""]; // Opción vacía por defecto
    if (targetCustomerId || customerName) {
      vehicleNames = getVehicleNamesForCustomer_(vehSheet, targetCustomerId, customerName);
    }
    
    // 3. Aplicar Validación de Datos al Dropdown
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(vehicleNames, true)
      .setAllowInvalid(false)
      .build();
      
    woSheet.getRange(rowNumber, vehicleCol).setDataValidation(rule);
    
    // Limpiar celda si el valor actual ya no está en la nueva lista
    const currentVal = woSheet.getRange(rowNumber, vehicleCol).getValue();
    if (currentVal && vehicleNames.indexOf(currentVal) === -1) {
      woSheet.getRange(rowNumber, vehicleCol).clearContent();
    }
    
    // Forzar refresh del viewer tras actualizar dropdown
    triggerVehicleViewerUpdate(SHEETS.WORK_ORDERS, rowNumber, vehicleCol);
    
  } catch (err) {
    Logger.log("Error actualizando dropdown WO: " + err.toString());
  }
}

