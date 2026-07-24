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

  Logger.log("================================");
  Logger.log("getVehicleViewerDataFast()");
  Logger.log("================================");

  const context = getSelectedVehicleContext_();

  Logger.log(context);

  if (!context || !context.vehicle) {

    Logger.log("NO CONTEXT");

    return null;

  }

  Logger.log("Vehicle encontrado:");
  Logger.log(context.vehicle.vehicleId);

  return prepareVehicleForViewer_(context);

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


function getSelectedVehicleFromWorkOrders_(sheet) {
  const activeRange = sheet.getActiveRange();

  if (!activeRange) return getDefaultVehicleContext_();

  const rowNumber = activeRange.getRow();

  if (rowNumber < TABLE.FIRST_DATA_ROW) return getDefaultVehicleContext_();

  const row = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const workOrder = createWorkOrderObject(row);

  Logger.log("========== WORK ORDER ==========");

  Logger.log(JSON.stringify(workOrder));

  const vehicleContext = findVehicleForWorkOrder_(workOrder);

  Logger.log("Vehicle Context:");

  Logger.log(vehicleContext);

  if (!vehicleContext) return null;

  return {
    vehicle: vehicleContext.vehicle,
    vehicleRow: vehicleContext.vehicleRow,
    workOrder: workOrder
  };
}


function findVehicleForWorkOrder_(workOrder) {

  Logger.log("========== findVehicleForWorkOrder ==========");

  Logger.log("VehicleID:");

  Logger.log(workOrder.vehicleId);

  Logger.log("VehicleName:");

  Logger.log(workOrder.vehicleName);

  Logger.log("CustomerID:");

  Logger.log(workOrder.customerId);

  Logger.log("CustomerName:");

  Logger.log(workOrder.customerName);


  if (workOrder.vehicleId) {
    const byId = findVehicleById_(workOrder.vehicleId);
    if (byId) return byId;
  }

  return findVehicleByNameAndCustomer_(
    workOrder.vehicleName,
    workOrder.customerId,
    workOrder.customerName
  );
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

  return {
    vehicle: createVehicleObject(row),
    vehicleRow: rowNumber
  };
}


function findVehicleByNameAndCustomer_(vehicleName, customerId, customerName) {
  vehicleName = String(vehicleName || "").trim();
  customerId = String(customerId || "").trim();
  customerName = String(customerName || "").trim();

  if (!vehicleName) return null;

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

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

    const sameVehicle =
      vehicle.displayName === vehicleName ||
      vehicle.vehicleName === vehicleName;

    const sameCustomer =
      !customerId ||
      vehicle.customerId === customerId ||
      vehicle.customerName === customerName;

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

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) return null;

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
  
  Logger.log("ENTRO A prepareVehicleForViewer");

  const vehicle = context.vehicle;
  const expectedImageName = VehicleImageService.buildImageName(vehicle);
  const expectedFilename = DriveService.buildImageFilename(expectedImageName);

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

  Logger.log(JSON.stringify(result));

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