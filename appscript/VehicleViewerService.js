//VehicleViewerService.gs

function openVehicleViewerSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("VehicleViewer")
    .setTitle("Vehicle Viewer");

  SpreadsheetApp.getUi().showSidebar(html);
}


function createVehicleObject(row) {
  const fields = ModuleConfig.get(SHEETS.VEHICLES).fields;

  const get = (field) => {
    const value = row[fields[field] - 1];
    return typeof value === "string" ? value.trim() : value;
  };

  return {
    vehicleId: get("VehicleID"),
    customerId: get("CustomerID"),
    customerName: get("CustomerName"),
    licensePlate: get("LicensePlate"),
    make: get("Make"),
    model: get("Model"),
    year: get("Year"),
    transmission: get("Transmission"),
    color: get("Color"),
    fuelType: get("FuelType"),
    status: get("Status"),
    notes: get("Notes"),
    vehicleName: get("VehicleName"),
    displayName: `${get("Make")} ${get("Model")} ${get("Year")}`.trim(),
    imageFileId: get("ImageFileID") || ""
  };
}


function createWorkOrderObject(row) {
  const fields = ModuleConfig.get(SHEETS.WORK_ORDERS).fields;

  const get = (field) => {
    const value = row[fields[field] - 1];
    return typeof value === "string" ? value.trim() : value;
  };

  return {
    workOrderId: get("WorkOrderID"),
    customerName: get("CustomerName"),
    customerId: get("CustomerID"),
    vehicleName: get("VehicleName"),
    vehicleId: get("VehicleID"),
    mechanicName: get("MechanicName"),
    mechanicId: get("MechanicID"),
    openDate: get("OpenDate"),
    status: get("Status"),
    priority: get("Priority"),
    mileage: get("Mileage"),
    complaint: get("Complaint"),
    diagnosis: get("Diagnosis"),
    laborCost: get("LaborCost"),
    partCost: get("PartCost"),
    totalCost: get("TotalCost"),
    completionDate: get("CompletionDate")
  };
}


function getVehicleViewerDataFast() {

  try {

    Logger.log("================================");
    Logger.log("getVehicleViewerDataFast()");
    Logger.log("================================");

    const context = getSelectedVehicleContext_();

    Logger.log(context);

    if (!context || !context.vehicle) {

      Logger.log("NO CONTEXT");

      return { error: "No vehicle context available" };

    }

    Logger.log("Vehicle encontrado:");
    Logger.log(context.vehicle.vehicleId);

    const result = prepareVehicleForViewer_(context);

    Logger.log(JSON.stringify(result));

    return result;

  } catch (err) {

    Logger.log("ERROR en getVehicleViewerDataFast: " + err.message);

    return { error: err.message };

  }

}


/**
 * Obtiene la firma de la selección actual y el timestamp actual.
 * @param {Sheet} [sheet] - Hoja opcional. Si no se pasa, usa la activa.
 * @returns {Object} { signature: string, timestamp: number }
 */
function getVehicleViewerSelectionSignature(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }
  
  const sheetName = sheet.getName();
  const activeRange = sheet.getActiveRange();
  
  let signature = "NONE";
  
  if (activeRange) {
    const row = activeRange.getRow();
    const col = activeRange.getColumn();
    signature = sheetName + "|" + row + "|" + col;
  }

  // CORRECCIÓN: Obtener el servicio de caché correctamente
  const cache = CacheService.getUserCache();
  const cachedTimestamp = cache.get('vehicleViewerTimestamp');
  
  const timestamp = cachedTimestamp ? parseInt(cachedTimestamp, 10) : new Date().getTime();

  return {
    signature: signature,
    timestamp: timestamp
  };
}


function generateSelectedVehicleImage(vehicleId, expectedImageName) {
  const context = findVehicleById_(vehicleId);

  if (!context || !context.vehicle) {
    return {
      success: false,
      error: "Vehicle not found."
    };
  }

  const currentImageName = VehicleImageService.buildImageName(context.vehicle);

  if (currentImageName !== expectedImageName) {
    return {
      success: false,
      stale: true,
      error: "Vehicle changed before image generation completed."
    };
  }

  const result = VehicleImageService.ensureVehicleImage(context.vehicle);

  if (result.success && result.fileId) {
    saveVehicleImageFileId_(context.vehicleRow, result.fileId);
  }

  return result;
}


function getSelectedVehicleContext_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  Logger.log("Hoja activa:");
  Logger.log(sheet.getName());

  Logger.log("Fila activa:");
  Logger.log(sheet.getActiveRange().getRow());

  switch (sheet.getName()) {

    case SHEETS.VEHICLES:

      return getSelectedVehicleFromVehicles_(sheet);

    case SHEETS.WORK_ORDERS:

      return getSelectedVehicleFromWorkOrders_(sheet);

    default:

      return getDefaultVehicleContext_();

  }

}


function getSelectedVehicleFromVehicles_(sheet) {
  const activeRange = sheet.getActiveRange();

  if (!activeRange) return getDefaultVehicleContext_();

  const rowNumber = activeRange.getRow();

  if (rowNumber < TABLE.FIRST_DATA_ROW) return getDefaultVehicleContext_();

  const row = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  return {
    vehicle: createVehicleObject(row),
    vehicleRow: rowNumber,
    workOrder: null
  };
}


/**
 * Obtiene el contexto del vehículo desde Work Orders.
 * CORRECCIÓN: Usa una lógica estricta para evitar seleccionar el vehículo equivocado
 * cuando hay múltiples vehículos con nombres similares para el mismo cliente.
 */
function getSelectedVehicleFromWorkOrders_(sheet) {
  const activeRange = sheet.getActiveRange();
  if (!activeRange) return getDefaultVehicleContext_();
  
  const rowNumber = activeRange.getRow();
  if (rowNumber < TABLE.FIRST_DATA_ROW) return getDefaultVehicleContext_();
  
  // Usar getDisplayValues() para leer el texto visible del dropdown
  const rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const workOrder = createWorkOrderObject(rowValues);
  
  Logger.log("WO Detectada: Fila " + rowNumber + ", Vehículo: '" + workOrder.vehicleName + "'");
  
  // Intentar encontrar el vehículo con lógica estricta
  const vehicleContext = findVehicleForWorkOrder_(workOrder, rowNumber);
  
  if (vehicleContext && vehicleContext.vehicle) {
    vehicleContext.workOrder = workOrder;
    return vehicleContext;
  }
  
  // FALLBACK: Si no se encuentra, crear objeto vacío pero válido
  return {
    vehicle: {
      vehicleId: "",
      displayName: workOrder.vehicleName ? workOrder.vehicleName : "No vehicle selected",
      vehicleName: workOrder.vehicleName || "",
      customerName: workOrder.customerName || "Unknown",
      make: "", model: "", year: "", status: "",
      imageFileId: "", imageStatus: "ready"
    },
    vehicleRow: null,
    workOrder: workOrder
  };
}




function findVehicleForWorkOrder_(workOrder) {

  const vehicleSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEETS.VEHICLES);

  if (!vehicleSheet) return null;

  const lastRow = vehicleSheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) return null;

  const values = vehicleSheet.getRange(
    TABLE.FIRST_DATA_ROW,
    1,
    lastRow - TABLE.FIRST_DATA_ROW + 1,
    vehicleSheet.getLastColumn()
  ).getValues();

  const targetVehicleID = String(workOrder.vehicleID || "").trim();
  const targetCustomerID = String(workOrder.customerID || "").trim();
  const targetVehicleName = String(workOrder.vehicleName || "").trim();
  const targetCustomerName = String(workOrder.customerName || "").trim();

  for (let i = 0; i < values.length; i++) {

    const vehicle = createVehicleObject(values[i]);

    //=====================================================
    // 1. Exact match using Vehicle ID (highest priority)
    //=====================================================

    if (
      targetVehicleID &&
      vehicle.vehicleID === targetVehicleID
    ) {

      return {
        vehicle: vehicle,
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };

    }

  }

  for (let i = 0; i < values.length; i++) {

    const vehicle = createVehicleObject(values[i]);

    //=====================================================
    // 2. Vehicle Name + Customer ID
    //=====================================================

    if (
      targetVehicleName &&
      targetCustomerID &&
      vehicle.vehicleName === targetVehicleName &&
      vehicle.customerID === targetCustomerID
    ) {

      return {
        vehicle: vehicle,
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };

    }

  }

  for (let i = 0; i < values.length; i++) {

    const vehicle = createVehicleObject(values[i]);

    //=====================================================
    // 3. Vehicle Name + Customer Name
    //=====================================================

    if (
      targetVehicleName &&
      targetCustomerName &&
      vehicle.vehicleName === targetVehicleName &&
      vehicle.customerName === targetCustomerName
    ) {

      return {
        vehicle: vehicle,
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };

    }

  }

  return null;

}


function findVehicleById_(vehicleId) {
  vehicleId = String(vehicleId || "").trim();

  Logger.log("========== findVehicleById ==========");

  Logger.log(vehicleId);

  if (!vehicleId) return null;

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const config = ModuleConfig.get(SHEETS.VEHICLES);
  const idColumn = config.fields.VehicleID;

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) return null;

  const finderRange = sheet.getRange(
    TABLE.FIRST_DATA_ROW,
    idColumn,
    lastRow - TABLE.FIRST_DATA_ROW + 1,
    1
  );

  const cell = finderRange
    .createTextFinder(vehicleId)
    .matchEntireCell(true)
    .findNext();

  if (!cell) return null;

  const rowNumber = cell.getRow();

  const row = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];


  const vehicle = createVehicleObject(row);
  debug(
    "VehicleViewer",
    "findVehicleById",
    "FOUND",
    {
      row: rowNumber,
      vehicleId: vehicle.vehicleId,
      imageFileId: vehicle.imageFileId
    }
  );


  return {
    vehicle: createVehicleObject(row),
    vehicleRow: rowNumber
  };
}


function findVehicleByNameAndCustomer_(vehicleName, customerId, customerName) {
  vehicleName = String(vehicleName || "").trim();
  customerId = String(customerId || "").trim();
  customerName = String(customerName || "").trim();

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  if (!sheet) return null;

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) return null;

  const values = sheet
    .getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const vehicle = createVehicleObject(values[i]);

    let sameVehicle = false;
    
    if (vehicleName) {
      sameVehicle =
        vehicle.displayName.toLowerCase() === vehicleName.toLowerCase() ||
        vehicle.vehicleName.toLowerCase() === vehicleName.toLowerCase();
    }

    const sameCustomer =
      !customerId ||
      vehicle.customerId === customerId ||
      vehicle.customerName.toLowerCase() === customerName.toLowerCase();

    // Si no hay vehicleName, buscar solo por cliente
    if (!vehicleName && sameCustomer) {
      return {
        vehicle: vehicle,
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };
    }
    
    // Si hay vehicleName, debe coincidir vehículo Y cliente
    if (sameVehicle && sameCustomer) {
      return {
        vehicle: vehicle,
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };
    }
  }

  return null;
}


function getDefaultVehicleContext_() {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  if (!sheet) {
    // Si no existe la hoja Vehicles, retornar un contexto vacío válido
    return {
      vehicle: {
        vehicleId: "",
        customerId: "",
        customerName: "",
        licensePlate: "",
        make: "",
        model: "",
        year: "",
        transmission: "",
        color: "",
        fuelType: "",
        status: "",
        notes: "",
        vehicleName: "",
        displayName: "No vehicle selected",
        imageFileId: "",
        imageStatus: "ready"
      },
      vehicleRow: null,
      workOrder: null
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    // Si no hay datos, retornar un contexto vacío válido
    return {
      vehicle: {
        vehicleId: "",
        customerId: "",
        customerName: "",
        licensePlate: "",
        make: "",
        model: "",
        year: "",
        transmission: "",
        color: "",
        fuelType: "",
        status: "",
        notes: "",
        vehicleName: "",
        displayName: "No vehicles available",
        imageFileId: "",
        imageStatus: "ready"
      },
      vehicleRow: null,
      workOrder: null
    };
  }

  const row = sheet
    .getRange(TABLE.FIRST_DATA_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  return {
    vehicle: createVehicleObject(row),
    vehicleRow: TABLE.FIRST_DATA_ROW,
    workOrder: null
  };
}


function prepareVehicleForViewer_(context) {

  const vehicle = context.vehicle;

  // Si no hay vehicleRow (vehículo no encontrado en 02_Vehicles), retornar datos básicos sin intentar operaciones de imagen
  if (!context.vehicleRow) {
    return Object.assign({}, vehicle, {
      workOrder: context.workOrder,
      expectedImageName: "",
      imageStatus: "ready"
    });
  }


  const expectedImageName = VehicleImageService.buildImageName(vehicle);

  Logger.log("expectedImageName:");
  Logger.log(expectedImageName);

  const expectedFilename = DriveService.buildImageFilename(expectedImageName);

  Logger.log("expectedFilename:");
  Logger.log(expectedFilename);

  if (vehicle.imageFileId) {
    const file = DriveService.getFileById(vehicle.imageFileId);

    if (file && file.getName() === expectedFilename) {
      return Object.assign({}, vehicle, {
        workOrder: context.workOrder,
        expectedImageName: expectedImageName,
        imageStatus: "ready"
      });
    }

    clearVehicleImageFileId_(context.vehicleRow);
    trashVehicleImageIfUnused_(vehicle.imageFileId, context.vehicleRow);
    vehicle.imageFileId = "";
  }

  const existingFile = DriveService.findImage(expectedImageName);

  if (existingFile) {
    saveVehicleImageFileId_(context.vehicleRow, existingFile.getId());

    return Object.assign({}, vehicle, {
      imageFileId: existingFile.getId(),
      workOrder: context.workOrder,
      expectedImageName: expectedImageName,
      imageStatus: "ready"
    });
  }

  debug(
    "VehicleViewer",
    "prepareVehicleForViewer",
    "RETURN",
    {
      needsGeneration: true,
      imageStatus: "needs-generation",
      expectedImageName: expectedImageName
    }
  );

  return Object.assign({}, vehicle, {
    imageFileId: "",
    workOrder: context.workOrder,
    expectedImageName: expectedImageName,
    imageStatus: "needs-generation"
  });
}


function saveVehicleImageFileId_(vehicleRow, fileId) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const imageColumn = ModuleConfig
    .get(SHEETS.VEHICLES)
    .fields
    .ImageFileID;

  sheet.getRange(vehicleRow, imageColumn).setValue(fileId);
}


function clearVehicleImageFileId_(vehicleRow) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const imageColumn = ModuleConfig
    .get(SHEETS.VEHICLES)
    .fields
    .ImageFileID;

  sheet.getRange(vehicleRow, imageColumn).clearContent();
}


function trashVehicleImageIfUnused_(fileId, currentVehicleRow) {
  if (!fileId) return;

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const imageColumn = ModuleConfig
    .get(SHEETS.VEHICLES)
    .fields
    .ImageFileID;

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    DriveService.trashFile(fileId);
    return;
  }

  const values = sheet
    .getRange(
      TABLE.FIRST_DATA_ROW,
      imageColumn,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      1
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = TABLE.FIRST_DATA_ROW + i;

    if (rowNumber === currentVehicleRow) continue;

    if (String(values[i][0] || "").trim() === fileId) {
      return;
    }
  }

  DriveService.trashFile(fileId);
}