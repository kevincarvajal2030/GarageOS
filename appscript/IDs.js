/**
 * IDs.gs
 * Generates a formatted ID.
 *
 * @param {string} prefix
 * @param {number} number
 * @returns {string}
 */
function formatID(prefix, number) {

  return `${prefix}-${String(number).padStart(6, "0")}`;

}

/**
 * Returns the next available number.
 *
 * @param {string} key
 * @returns {number}
 */
function getCounter(key) {

  const sheet = getSheet(SHEETS.SYSTEM);

  const cell = sheet
    .createTextFinder(key)
    .matchEntireCell(true)
    .findNext();

  if (!cell) {

    throw new Error("Counter not found: " + key);

  }

  return Number(
    sheet.getRange(cell.getRow(), 2).getValue()
  );

}


/**
 * Updates a counter value.
 *
 * @param {string} key
 * @param {number} value
 */
function setCounter(key, value) {

  const sheet = getSheet(SHEETS.SYSTEM);

  const cell = sheet
    .createTextFinder(key)
    .matchEntireCell(true)
    .findNext();

  if (!cell) {

    throw new Error("Counter not found: " + key);

  }

  sheet
    .getRange(cell.getRow(), 2)
    .setValue(value);

}


/**
 * Generates a unique ID.
 *
 * @param {string} counterKey
 * @param {string} prefix
 * @returns {string}
 */
function generateID(counterKey, prefix) {

  const number = getCounter(counterKey);

  setCounter(counterKey, number + 1);

  return formatID(prefix, number);

}


function generateCustomerID() {

  return generateID(
    COUNTERS.CUSTOMER,
    ID_PREFIX.CUSTOMER
  );

}

function generateVehicleID() {

  return generateID(
    COUNTERS.VEHICLE,
    ID_PREFIX.VEHICLE
  );

}

function generateWorkOrderID() {

  return generateID(
    COUNTERS.WORK_ORDER,
    ID_PREFIX.WORK_ORDER
  );

}

function generateServiceID() {

  return generateID(
    COUNTERS.SERVICE,
    ID_PREFIX.SERVICE
  );

}

function generatePartID() {

  return generateID(
    COUNTERS.PART,
    ID_PREFIX.PART
  );

}

function generateSupplierID() {

  return generateID(
    COUNTERS.SUPPLIER,
    ID_PREFIX.SUPPLIER
  );

}

function generatePurchaseID() {

  return generateID(
    COUNTERS.PURCHASE,
    ID_PREFIX.PURCHASE
  );

}

function generatePaymentID() {

  return generateID(
    COUNTERS.PAYMENT,
    ID_PREFIX.PAYMENT
  );

}

function generateMechanicID() {

  return generateID(
    COUNTERS.MECHANIC,
    ID_PREFIX.MECHANIC
  );

}
