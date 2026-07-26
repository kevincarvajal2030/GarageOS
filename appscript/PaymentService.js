
/**
 * ============================================================================
 * GarageOS
 * PaymentService.gs
 * ----------------------------------------------------------------------------
 * Central service for managing Payment records.
 *
 * This service handles:
 * - Automatic Payment creation from completed Work Orders
 * - Payment synchronization with Work Orders
 * - Payment date management
 *
 * Payments are automatically created when a Work Order status changes to "Completed".
 * Payments become the source of truth for financial status after creation.
 * ============================================================================
 */

const PaymentService = (() => {

  /**
   * Returns the Payments sheet.
   *
   * @returns {Sheet} The Payments sheet object.
   */
  function getPaymentsSheet() {
    return getSheet(SHEETS.PAYMENTS);
  }

  /**
   * Returns the Work Orders sheet.
   *
   * @returns {Sheet} The Work Orders sheet object.
   */
  function getWorkOrdersSheet() {
    return getSheet(SHEETS.WORK_ORDERS);
  }

  /**
   * Checks if a Payment record already exists for a given Work Order ID.
   *
   * Business Rule 2: Avoid Duplicate Payment Records
   * Before creating a Payment, search the Payments sheet.
   * If another Payment already exists with the same Work Order ID, do NOT create another one.
   *
   * @param {string} workOrderId - The Work Order ID to check.
   * @returns {boolean} True if a Payment exists, false otherwise.
   */
  function paymentExists(workOrderId) {
    const paymentsSheet = getPaymentsSheet();
    const config = ModuleConfig.get(SHEETS.PAYMENTS);

    const lastRow = paymentsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return false;
    }

    const workOrderIdColumn = config.fields.WorkOrderID;
    const values = paymentsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      workOrderIdColumn,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      1
    ).getValues();

    workOrderId = String(workOrderId).trim();

    for (const row of values) {
      if (String(row[0]).trim() === workOrderId) {
        return true;
      }
    }

    return false;
  }

  /**
   * Creates a Payment record from a Work Order.
   *
   * Business Rule 1: Automatic Payment Creation
   * When a Work Order changes to "Completed", automatically create one Payment record.
   *
   * @param {Object} workOrderData - The Work Order data object containing all required fields.
   * @returns {string} The generated Payment ID.
   */
  function createPaymentFromWorkOrder(workOrderData) {
    const paymentsSheet = getPaymentsSheet();
    const config = ModuleConfig.get(SHEETS.PAYMENTS);

    // Generate new Payment ID
    const paymentId = generatePaymentID();

    // Find the first empty row
    const newRow = paymentsSheet.getLastRow() + 1;

    // Populate Payment record
    // Column 1: Payment ID (system-generated)
    paymentsSheet.getRange(newRow, config.fields.PaymentID).setValue(paymentId);

    // Column 2: Customer Name (copied from Work Order)
    paymentsSheet.getRange(newRow, config.fields.CustomerName).setValue(workOrderData.customerName);

    // Column 3: Customer ID (copied from Work Order)
    paymentsSheet.getRange(newRow, config.fields.CustomerID).setValue(workOrderData.customerId);

    // Column 4: Work Order ID (copied from Work Order)
    paymentsSheet.getRange(newRow, config.fields.WorkOrderID).setValue(workOrderData.workOrderId);

    // Column 5: Payment Method (empty by default - user will choose later)
    paymentsSheet.getRange(newRow, config.fields.PaymentMethod).clearContent();

    // Column 6: Total Cost (copied from Work Order)
    paymentsSheet.getRange(newRow, config.fields.TotalCost).setValue(workOrderData.totalCost);

    // Column 7: Payment Status (Pending by default)
    paymentsSheet.getRange(newRow, config.fields.PaymentStatus).setValue("Pending");

    // Column 8: Payment Date (empty by default - payment has not happened yet)
    paymentsSheet.getRange(newRow, config.fields.PaymentDate).clearContent();

    // Column 9: Notes (empty by default)
    paymentsSheet.getRange(newRow, config.fields.Notes).clearContent();

    return paymentId;
  }

  /**
   * Populates a Payment record with Work Order data.
   * Helper function to extract Work Order data and pass to createPaymentFromWorkOrder.
   *
   * @param {number} workOrderRow - The row number of the Work Order in the Work Orders sheet.
   * @returns {string|null} The generated Payment ID, or null if creation failed.
   */
  function populatePaymentRecord(workOrderRow) {
    const workOrdersSheet = getWorkOrdersSheet();
    const woConfig = ModuleConfig.get(SHEETS.WORK_ORDERS);

    const workOrderId = workOrdersSheet.getRange(workOrderRow, woConfig.fields.WorkOrderID).getDisplayValue().trim();
    const customerName = workOrdersSheet.getRange(workOrderRow, woConfig.fields.CustomerName).getDisplayValue().trim();
    const customerId = workOrdersSheet.getRange(workOrderRow, woConfig.fields.CustomerID).getDisplayValue().trim();
    const totalCost = workOrdersSheet.getRange(workOrderRow, woConfig.fields.TotalCost).getValue();

    if (!workOrderId || !customerName || !customerId) {
      return null;
    }

    const workOrderData = {
      workOrderId: workOrderId,
      customerName: customerName,
      customerId: customerId,
      totalCost: totalCost
    };

    return createPaymentFromWorkOrder(workOrderData);
  }

  /**
   * Updates the Payment Status in the Work Order when Payment Status changes.
   *
   * Business Rule 4: Synchronize Payment Status
   * When Payment Status changes in Payments, automatically update Work Orders Payment Status.
   *
   * @param {string} workOrderId - The Work Order ID to update.
   * @param {string} paymentStatus - The new Payment Status value.
   */
  function updateWorkOrderPaymentStatus(workOrderId, paymentStatus) {
    const workOrdersSheet = getWorkOrdersSheet();
    const woConfig = ModuleConfig.get(SHEETS.WORK_ORDERS);

    const lastRow = workOrdersSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return;
    }

    const workOrderIdColumn = woConfig.fields.WorkOrderID;
    const paymentStatusColumn = woConfig.fields.PaymentStatus;

    const values = workOrdersSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      workOrderIdColumn,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      1
    ).getValues();

    workOrderId = String(workOrderId).trim();

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === workOrderId) {
        const targetRow = TABLE.FIRST_DATA_ROW + i;
        workOrdersSheet.getRange(targetRow, paymentStatusColumn).setValue(paymentStatus);
        break;
      }
    }
  }

  /**
   * Synchronizes Payment Status from Payment record to Work Order.
   * Looks up the Work Order ID from the Payment record and updates the Work Order.
   *
   * @param {number} paymentRow - The row number of the Payment in the Payments sheet.
   */
  function synchronizePaymentStatus(paymentRow) {
    const paymentsSheet = getPaymentsSheet();
    const payConfig = ModuleConfig.get(SHEETS.PAYMENTS);

    const workOrderId = paymentsSheet.getRange(paymentRow, payConfig.fields.WorkOrderID).getDisplayValue().trim();
    const paymentStatus = paymentsSheet.getRange(paymentRow, payConfig.fields.PaymentStatus).getDisplayValue().trim();

    if (workOrderId) {
      updateWorkOrderPaymentStatus(workOrderId, paymentStatus);
    }
  }

  /**
   * Updates the Payment Date based on Payment Status.
   *
   * Business Rule 5: Automatic Payment Date
   * - If Payment Status is Pending, Paid, or Refunded: Set Payment Date to today's date.
   * - If Payment Status is empty: Clear Payment Date.
   *
   * Payment Date always represents the most recent payment activity.
   *
   * @param {number} paymentRow - The row number of the Payment in the Payments sheet.
   */
  function updatePaymentDate(paymentRow) {
    const paymentsSheet = getPaymentsSheet();
    const config = ModuleConfig.get(SHEETS.PAYMENTS);

    const paymentStatus = paymentsSheet.getRange(paymentRow, config.fields.PaymentStatus).getDisplayValue().trim();
    const paymentDateColumn = config.fields.PaymentDate;
    const paymentDateCell = paymentsSheet.getRange(paymentRow, paymentDateColumn);

    if (paymentStatus === "") {
      // Clear Payment Date if status is empty
      paymentDateCell.clearContent();
    } else if (paymentStatus === "Pending" || paymentStatus === "Paid" || paymentStatus === "Refunded") {
      // Set Payment Date to today's date for valid payment statuses
      paymentDateCell.setValue(
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "MM/dd/yyyy"
        )
      );
    }
  }

  /**
   * Handles Payment Status changes.
   * Called when Payment Status is modified to:
   * 1. Update Payment Date
   * 2. Synchronize Payment Status to Work Order
   *
   * @param {number} paymentRow - The row number of the Payment in the Payments sheet.
   */
  function handlePaymentStatusChange(paymentRow) {
    updatePaymentDate(paymentRow);
    synchronizePaymentStatus(paymentRow);
  }

  /**
   * Main entry point for creating a Payment from a Work Order.
   * Checks if Payment already exists before creating.
   *
   * Business Rule 1 & 2: Automatic Payment Creation without Duplicates
   *
   * @param {number} workOrderRow - The row number of the Work Order in the Work Orders sheet.
   * @returns {string|null} The generated Payment ID, or null if no Payment was created.
   */
  function processWorkOrderCompletion(workOrderRow) {
    const workOrdersSheet = getWorkOrdersSheet();
    const woConfig = ModuleConfig.get(SHEETS.WORK_ORDERS);

    const workOrderId = workOrdersSheet.getRange(workOrderRow, woConfig.fields.WorkOrderID).getDisplayValue().trim();

    if (!workOrderId) {
      return null;
    }

    // Check if Payment already exists (Business Rule 2)
    if (paymentExists(workOrderId)) {
      return null;
    }

    // Create Payment record (Business Rule 1)
    const paymentId = populatePaymentRecord(workOrderRow);

    if (paymentId) {
      // Business Rule 3: Synchronize Payment Status
      // Update Work Order Payment Status to Pending
      const paymentStatusColumn = woConfig.fields.PaymentStatus;
      workOrdersSheet.getRange(workOrderRow, paymentStatusColumn).setValue("Pending");
    }

    return paymentId;
  }

  return {
    // Public functions
    paymentExists,
    createPaymentFromWorkOrder,
    populatePaymentRecord,
    updateWorkOrderPaymentStatus,
    synchronizePaymentStatus,
    updatePaymentDate,
    handlePaymentStatusChange,
    processWorkOrderCompletion
  };

})();
