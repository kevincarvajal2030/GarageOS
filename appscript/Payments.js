/**
 * ============================================================================
 * GARAGE OS
 * Payments Module
 * ============================================================================
 *
 * Handles all payment-related business logic.
 *
 * Responsibilities:
 *   - Create payments.
 *   - Update payments.
 *   - Delete payments.
 *   - Validate payment information.
 *   - Build payment records.
 *
 * Uses:
 *   - Validation.gs
 *   - Database.gs
 *   - IDs.gs
 * ============================================================================
 */


/******************************************************************************
 * PAYMENT CRUD
 ******************************************************************************/

/**
 * Creates a new payment.
 *
 * @param {Object} paymentData
 */
function createPayment(paymentData) {

  validatePaymentData(paymentData);

  const record = buildPaymentRecord(paymentData);

  appendRecord(
    SHEETS.PAYMENTS,
    record
  );

}


/**
 * Updates an existing payment.
 *
 * @param {string} paymentId
 * @param {Object} paymentData
 */
function updatePayment(paymentId, paymentData) {

  validatePaymentData(paymentData);

  const existingRecord = getRecord(
    SHEETS.PAYMENTS,
    paymentId
  );

  if (!existingRecord) {

    throw new Error("Payment not found.");

  }

  const updatedRecord = [

    paymentId,

    paymentData.workOrderId,

    paymentData.date,

    paymentData.method,

    paymentData.amount,

    paymentData.status,

    paymentData.notes || "",

    existingRecord[7]

  ];

  updateRecord(
    SHEETS.PAYMENTS,
    paymentId,
    updatedRecord
  );

}


/**
 * Deletes a payment.
 *
 * @param {string} paymentId
 */
function deletePayment(paymentId) {

  deleteRecord(
    SHEETS.PAYMENTS,
    paymentId
  );

}


/******************************************************************************
 * PAYMENT QUERIES
 ******************************************************************************/

/**
 * Returns a payment by ID.
 *
 * @param {string} paymentId
 * @returns {Object|null}
 */
function getPaymentById(paymentId) {

  const record = getRecord(
    SHEETS.PAYMENTS,
    paymentId
  );

  if (!record) {

    return null;

  }

  return {

    id: record[0],

    workOrderId: record[1],

    date: record[2],

    method: record[3],

    amount: record[4],

    status: record[5],

    notes: record[6],

    createdAt: record[7]

  };

}


/**
 * Determines whether a payment exists.
 *
 * @param {string} paymentId
 * @returns {boolean}
 */
function paymentExists(paymentId) {

  return recordExists(
    SHEETS.PAYMENTS,
    paymentId
  );

}


/******************************************************************************
 * PAYMENT VALIDATION
 ******************************************************************************/

/**
 * Validates payment information.
 *
 * @param {Object} paymentData
 */
function validatePaymentData(paymentData) {

  validateObjectField(
    paymentData,
    "Payment"
  );

  validateTextField(
    paymentData.workOrderId,
    "Work Order",
    {
      required: true
    }
  );

  validateDateField(
    paymentData.date,
    "Payment Date"
  );

  validateTextField(
    paymentData.method,
    "Payment Method",
    {
      required: true
    }
  );

  validatePositiveNumberField(
    paymentData.amount,
    "Amount"
  );

  validateTextField(
    paymentData.status,
    "Status",
    {
      required: true
    }
  );

  validateTextField(
    paymentData.notes || "",
    "Notes",
    {
      required: false,
      maximumLength: 500
    }
  );

}


/******************************************************************************
 * PAYMENT RECORD
 ******************************************************************************/

/**
 * Builds a payment record.
 *
 * @param {Object} paymentData
 * @returns {Array}
 */
function buildPaymentRecord(paymentData) {

  return [

    generatePaymentID(),

    paymentData.workOrderId,

    paymentData.date,

    paymentData.method,

    paymentData.amount,

    paymentData.status,

    paymentData.notes || "",

    getTimestamp()

  ];

}
