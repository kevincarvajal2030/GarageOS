/**
 * ============================================================================
 * GARAGE OS
 * Work Orders Module
 * ============================================================================
 *
 * Handles all work order business logic.
 *
 * Responsibilities:
 *   - Create work orders.
 *   - Update work orders.
 *   - Delete work orders.
 *   - Validate work order information.
 *   - Build work order records.
 *
 * Uses:
 *   - Validation.gs
 *   - Database.gs
 *   - IDs.gs
 * ============================================================================
 */


/******************************************************************************
 * WORK ORDER CRUD
 ******************************************************************************/

/**
 * Creates a new work order.
 *
 * @param {Object} workOrderData
 */
function createWorkOrder(workOrderData) {

  validateWorkOrderData(workOrderData);

  const record = buildWorkOrderRecord(workOrderData);

  appendRecord(
    SHEETS.WORK_ORDERS,
    record
  );

}


/**
 * Updates an existing work order.
 *
 * @param {string} workOrderId
 * @param {Object} workOrderData
 */
function updateWorkOrder(workOrderId, workOrderData) {

  validateWorkOrderData(workOrderData);

  const existingRecord = getRecord(
    SHEETS.WORK_ORDERS,
    workOrderId
  );

  if (!existingRecord) {

    throw new Error("Work order not found.");

  }

  const updatedRecord = [

    workOrderId,

    workOrderData.customerId,

    workOrderData.vehicleId,

    workOrderData.date,

    workOrderData.status,

    workOrderData.problem,

    workOrderData.mechanic,

    workOrderData.total,

    existingRecord[8]

  ];

  updateRecord(
    SHEETS.WORK_ORDERS,
    workOrderId,
    updatedRecord
  );

}


/**
 * Deletes a work order.
 *
 * @param {string} workOrderId
 */
function deleteWorkOrder(workOrderId) {

  deleteRecord(
    SHEETS.WORK_ORDERS,
    workOrderId
  );

}


/******************************************************************************
 * WORK ORDER QUERIES
 ******************************************************************************/

/**
 * Returns a work order by ID.
 *
 * @param {string} workOrderId
 * @returns {Object|null}
 */
function getWorkOrderById(workOrderId) {

  const record = getRecord(
    SHEETS.WORK_ORDERS,
    workOrderId
  );

  if (!record) {

    return null;

  }

  return {

    id: record[0],

    customerId: record[1],

    vehicleId: record[2],

    date: record[3],

    status: record[4],

    problem: record[5],

    mechanic: record[6],

    total: record[7],

    createdAt: record[8]

  };

}


/**
 * Determines whether a work order exists.
 *
 * @param {string} workOrderId
 * @returns {boolean}
 */
function workOrderExists(workOrderId) {

  return recordExists(
    SHEETS.WORK_ORDERS,
    workOrderId
  );

}


/******************************************************************************
 * WORK ORDER VALIDATION
 ******************************************************************************/

/**
 * Validates work order information.
 *
 * @param {Object} workOrderData
 */
function validateWorkOrderData(workOrderData) {

  validateObjectField(
    workOrderData,
    "Work Order"
  );

  validateTextField(
    workOrderData.customerId,
    "Customer",
    {
      required: true
    }
  );

  validateTextField(
    workOrderData.vehicleId,
    "Vehicle",
    {
      required: true
    }
  );

  validateDateField(
    workOrderData.date,
    "Date"
  );

  validateTextField(
    workOrderData.status,
    "Status",
    {
      required: true
    }
  );

  validateTextField(
    workOrderData.problem,
    "Problem",
    {
      required: true,
      minimumLength: 5,
      maximumLength: 500
    }
  );

  validateTextField(
    workOrderData.mechanic,
    "Mechanic",
    {
      required: true
    }
  );

  validatePositiveNumberField(
    workOrderData.total,
    "Total"
  );

}


/******************************************************************************
 * WORK ORDER RECORD
 ******************************************************************************/

/**
 * Builds a work order record.
 *
 * @param {Object} workOrderData
 * @returns {Array}
 */
function buildWorkOrderRecord(workOrderData) {

  return [

    generateWorkOrderID(),

    workOrderData.customerId,

    workOrderData.vehicleId,

    workOrderData.date,

    workOrderData.status,

    workOrderData.problem,

    workOrderData.mechanic,

    workOrderData.total,

    getTimestamp()

  ];

}
