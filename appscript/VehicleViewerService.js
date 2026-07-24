function openVehicleViewerSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("VehicleViewer")
    .setTitle("Vehicle Viewer");

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}


function createVehicleObject(row) {
  const fields = ModuleConfig
    .get(SHEETS.VEHICLES)
    .fields;

  const get = (field) => {
    const value = row[fields[field] - 1];

    return typeof value === "string"
      ? value.trim()
      : value;
  };

  return Object.freeze({
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
  });
}


function createWorkOrderObject(row) {
  const fields = ModuleConfig
    .get(SHEETS.WORK_ORDERS)
    .fields;

  const get = (field) => {
    const value = row[fields[field] - 1];

    return typeof value === "string"
      ? value.trim()
      : value;
  };

  return Object.freeze({
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
  });
}


function getVehicleViewerData() {
  const context = getSelectedVehicleContext_();

  if (!context || !context.vehicle) {
    return null;
  }

  const vehicle = ensureVehicleImageForViewer_(context);

  return vehicle;
}


function getSelectedVehicleContext_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  switch (sheetName) {
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

  if (!activeRange) {
    return getDefaultVehicleContext_();
  }

  const rowNumber = activeRange.getRow();

  if (rowNumber < TABLE.FIRST_DATA_ROW) {
    return getDefaultVehicleContext_();
  }

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

  if (!activeRange) {
    return getDefaultVehicleContext_();
  }

  const rowNumber = activeRange.getRow();

  if (rowNumber < TABLE.FIRST_DATA_ROW) {
    return getDefaultVehicleContext_();
  }

  const workOrderRow = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const workOrder = createWorkOrderObject(workOrderRow);

  if (!workOrder.vehicleId) {
    return null;
  }

  const vehicleContext = findVehicleById_(workOrder.vehicleId);

  if (!vehicleContext) {
    return null;
  }

  return {
    vehicle: vehicleContext.vehicle,
    vehicleRow: vehicleContext.vehicleRow,
    workOrder: workOrder
  };
}


function getDefaultVehicleContext_() {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return null;
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


function findVehicleById_(vehicleId) {
  vehicleId = String(vehicleId || "").trim();

  if (!vehicleId) {
    return null;
  }

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const config = ModuleConfig.get(SHEETS.VEHICLES);
  const idColumn = config.fields.VehicleID;

  const lastRow = sheet.getLastRow();

  if (lastRow < TABLE.FIRST_DATA_ROW) {
    return null;
  }

  const values = sheet
    .getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const rowVehicleId = String(values[i][idColumn - 1] || "").trim();

    if (rowVehicleId === vehicleId) {
      return {
        vehicle: createVehicleObject(values[i]),
        vehicleRow: TABLE.FIRST_DATA_ROW + i
      };
    }
  }

  return null;
}


function ensureVehicleImageForViewer_(context) {
  const vehicle = context.vehicle;

  if (vehicle.imageFileId) {
    return Object.assign({}, vehicle, {
      workOrder: context.workOrder,
      imageStatus: "ready"
    });
  }

  if (!hasEnoughVehicleImageData_(vehicle)) {
    return Object.assign({}, vehicle, {
      workOrder: context.workOrder,
      imageStatus: "missing-data"
    });
  }

  const result = VehicleImageService.ensureVehicleImage(vehicle);

  if (!result.success || !result.fileId) {
    return Object.assign({}, vehicle, {
      workOrder: context.workOrder,
      imageStatus: "error",
      imageError: result.error || result.message || "Image generation failed."
    });
  }

  saveVehicleImageFileId_(context.vehicleRow, result.fileId);

  return Object.assign({}, vehicle, {
    imageFileId: result.fileId,
    workOrder: context.workOrder,
    imageStatus: result.source || "generated"
  });
}


function saveVehicleImageFileId_(vehicleRow, fileId) {
  if (!vehicleRow || !fileId) {
    return;
  }

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VEHICLES);

  const config = ModuleConfig.get(SHEETS.VEHICLES);
  const imageColumn = config.fields.ImageFileID;

  sheet
    .getRange(vehicleRow, imageColumn)
    .setValue(fileId);
}


function hasEnoughVehicleImageData_(vehicle) {
  return !!(
    vehicle &&
    vehicle.make &&
    vehicle.model &&
    vehicle.year &&
    vehicle.color
  );
}