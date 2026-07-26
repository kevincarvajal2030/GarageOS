
/**
 * ============================================================================
 * GarageOS
 * MechanicAssignmentService.gs
 * ----------------------------------------------------------------------------
 * Centralized service for synchronizing mechanic availability based on
 * Work Order assignments.
 *
 * This service is the single source of truth for mechanic status.
 * All modules should call this service instead of implementing their own logic.
 * ============================================================================
 */

const MechanicAssignmentService = (() => {

  /**
   * Active Work Order statuses that make a mechanic Busy.
   * @type {string[]}
   */
  const ACTIVE_STATUSES = ["Pending", "In Progress", "Waiting Parts"];

  /**
   * Inactive Work Order statuses that do NOT keep a mechanic Busy.
   * @type {string[]}
   */
  const INACTIVE_STATUSES = ["Completed", "Cancelled"];

  /**
   * Mechanic statuses that indicate unavailability for assignment.
   * @type {string[]}
   */
  const UNAVAILABLE_STATUSES = ["Busy", "Vacation", "Inactive"];

  /**
   * Checks if a mechanic has any active Work Orders.
   *
   * @param {string} mechanicId - The mechanic ID to check.
   * @returns {boolean} True if the mechanic has at least one active Work Order.
   */
  function hasActiveWorkOrders(mechanicId) {

    mechanicId = String(mechanicId || "").trim();

    if (!mechanicId) {
      return false;
    }

    const ss = SpreadsheetApp.getActive();
    const workOrdersSheet = ss.getSheetByName(SHEETS.WORK_ORDERS);

    if (!workOrdersSheet) {
      return false;
    }

    const lastRow = workOrdersSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return false;
    }

    const config = ModuleConfig.get(SHEETS.WORK_ORDERS);

    const values = workOrdersSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      workOrdersSheet.getLastColumn()
    ).getValues();

    for (const row of values) {

      const rowMechanicId = String(row[config.fields.MechanicID - 1] || "").trim();
      const status = String(row[config.fields.Status - 1] || "").trim();

      if (rowMechanicId === mechanicId && ACTIVE_STATUSES.includes(status)) {
        return true;
      }

    }

    return false;

  }

  /**
   * Synchronizes a mechanic's status based on their active Work Orders.
   * If the mechanic has active Work Orders, sets status to Busy.
   * If no active Work Orders exist and status is Busy, sets to Available.
   *
   * @param {string} mechanicId - The mechanic ID to synchronize.
   * @returns {boolean} True if the status was updated, false otherwise.
   */
  function syncMechanicStatus(mechanicId) {

    mechanicId = String(mechanicId || "").trim();

    if (!mechanicId) {
      return false;
    }

    const ss = SpreadsheetApp.getActive();
    const mechanicsSheet = ss.getSheetByName(SHEETS.MECHANICS);

    if (!mechanicsSheet) {
      return false;
    }

    const config = ModuleConfig.get(SHEETS.MECHANICS);
    const lastRow = mechanicsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return false;
    }

    // Find the mechanic row
    const values = mechanicsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      mechanicsSheet.getLastColumn()
    ).getValues();

    let mechanicRow = null;
    let currentStatus = null;

    for (let i = 0; i < values.length; i++) {

      const rowMechanicId = String(values[i][config.fields.MechanicID - 1] || "").trim();

      if (rowMechanicId === mechanicId) {
        mechanicRow = TABLE.FIRST_DATA_ROW + i;
        currentStatus = String(values[i][config.fields.Status - 1] || "").trim();
        break;
      }

    }

    if (!mechanicRow) {
      return false;
    }

    const hasActive = hasActiveWorkOrders(mechanicId);
    let newStatus = null;

    // Determine the correct status
    if (hasActive) {
      newStatus = "Busy";
    } else if (currentStatus === "Busy") {
      // Self-healing: if marked Busy but no active WO, set to Available
      newStatus = "Available";
    }

    // Only update if status needs to change
    if (newStatus && newStatus !== currentStatus) {
      mechanicsSheet.getRange(mechanicRow, config.fields.Status).setValue(newStatus);
      return true;
    }

    return false;

  }

  /**
   * Updates a mechanic's status directly.
   * Use with caution - this bypasses the automatic synchronization logic.
   *
   * @param {string} mechanicId - The mechanic ID to update.
   * @param {string} status - The new status value.
   * @returns {boolean} True if the update was successful.
   */
  function updateMechanicStatus(mechanicId, status) {

    mechanicId = String(mechanicId || "").trim();
    status = String(status || "").trim();

    if (!mechanicId || !status) {
      return false;
    }

    const ss = SpreadsheetApp.getActive();
    const mechanicsSheet = ss.getSheetByName(SHEETS.MECHANICS);

    if (!mechanicsSheet) {
      return false;
    }

    const config = ModuleConfig.get(SHEETS.MECHANICS);
    const lastRow = mechanicsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return false;
    }

    const values = mechanicsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      mechanicsSheet.getLastColumn()
    ).getValues();

    for (let i = 0; i < values.length; i++) {

      const rowMechanicId = String(values[i][config.fields.MechanicID - 1] || "").trim();

      if (rowMechanicId === mechanicId) {
        const targetRow = TABLE.FIRST_DATA_ROW + i;
        mechanicsSheet.getRange(targetRow, config.fields.Status).setValue(status);
        return true;
      }

    }

    return false;

  }

  /**
   * Refreshes mechanic availability by checking all mechanics with Busy status.
   * Corrects any inconsistencies where a mechanic is marked Busy but has no active WOs.
   *
   * @returns {number} The number of mechanics whose status was corrected.
   */
  function refreshMechanicAvailability() {

    const ss = SpreadsheetApp.getActive();
    const mechanicsSheet = ss.getSheetByName(SHEETS.MECHANICS);

    if (!mechanicsSheet) {
      return 0;
    }

    const config = ModuleConfig.get(SHEETS.MECHANICS);
    const lastRow = mechanicsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return 0;
    }

    const values = mechanicsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      mechanicsSheet.getLastColumn()
    ).getValues();

    let correctedCount = 0;

    for (let i = 0; i < values.length; i++) {

      const mechanicId = String(values[i][config.fields.MechanicID - 1] || "").trim();
      const currentStatus = String(values[i][config.fields.Status - 1] || "").trim();

      if (!mechanicId) {
        continue;
      }

      // Self-healing: check if Busy mechanics actually have active WOs
      if (currentStatus === "Busy") {

        const hasActive = hasActiveWorkOrders(mechanicId);

        if (!hasActive) {
          // Correct the inconsistency
          const targetRow = TABLE.FIRST_DATA_ROW + i;
          mechanicsSheet.getRange(targetRow, config.fields.Status).setValue("Available");
          correctedCount++;
        }

      }

    }

    return correctedCount;

  }

  /**
   * Gets all available mechanics (those with Available status).
   *
   * @returns {Array<{id: string, name: string}>} List of available mechanics.
   */
  function getAvailableMechanics() {

    const ss = SpreadsheetApp.getActive();
    const mechanicsSheet = ss.getSheetByName(SHEETS.MECHANICS);

    if (!mechanicsSheet) {
      return [];
    }

    const config = ModuleConfig.get(SHEETS.MECHANICS);
    const lastRow = mechanicsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const values = mechanicsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      mechanicsSheet.getLastColumn()
    ).getValues();

    const availableMechanics = [];

    for (const row of values) {

      const mechanicId = String(row[config.fields.MechanicID - 1] || "").trim();
      const firstName = String(row[config.fields.FirstName - 1] || "").trim();
      const lastName = String(row[config.fields.LastName - 1] || "").trim();
      const status = String(row[config.fields.Status - 1] || "").trim();

      if (status === "Available" && mechanicId && firstName && lastName) {
        availableMechanics.push({
          id: mechanicId,
          name: firstName + " " + lastName
        });
      }

    }

    return availableMechanics;

  }

  /**
   * Checks if a mechanic is available for assignment.
   *
   * @param {string} mechanicId - The mechanic ID to check.
   * @returns {{available: boolean, reason: string|null}} Availability status and reason if unavailable.
   */
  function isMechanicAvailable(mechanicId) {

    mechanicId = String(mechanicId || "").trim();

    if (!mechanicId) {
      return { available: false, reason: "No mechanic selected" };
    }

    // First, self-heal: sync the status based on actual WO data
    syncMechanicStatus(mechanicId);

    const ss = SpreadsheetApp.getActive();
    const mechanicsSheet = ss.getSheetByName(SHEETS.MECHANICS);

    if (!mechanicsSheet) {
      return { available: false, reason: "Mechanics sheet not found" };
    }

    const config = ModuleConfig.get(SHEETS.MECHANICS);
    const lastRow = mechanicsSheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return { available: false, reason: "No mechanics found" };
    }

    const values = mechanicsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      mechanicsSheet.getLastColumn()
    ).getValues();

    for (const row of values) {

      const rowMechanicId = String(row[config.fields.MechanicID - 1] || "").trim();

      if (rowMechanicId === mechanicId) {

        const status = String(row[config.fields.Status - 1] || "").trim();

        if (UNAVAILABLE_STATUSES.includes(status)) {
          return { available: false, reason: "Mechanic is " + status };
        }

        return { available: true, reason: null };

      }

    }

    return { available: false, reason: "Mechanic not found" };

  }

  /**
   * Handles mechanic reassignment from one mechanic to another.
   * Releases the old mechanic and reserves the new one.
   *
   * @param {string} oldMechanicId - The previous mechanic ID.
   * @param {string} newMechanicId - The new mechanic ID.
   */
  function handleMechanicReassignment(oldMechanicId, newMechanicId) {

    // Release old mechanic
    if (oldMechanicId && oldMechanicId !== newMechanicId) {
      syncMechanicStatus(oldMechanicId);
    }

    // Reserve new mechanic
    if (newMechanicId && newMechanicId !== oldMechanicId) {
      syncMechanicStatus(newMechanicId);
    }

  }

  return {

    // Core synchronization functions
    hasActiveWorkOrders,
    syncMechanicStatus,
    updateMechanicStatus,
    refreshMechanicAvailability,

    // Utility functions
    getAvailableMechanics,
    isMechanicAvailable,
    handleMechanicReassignment,

    // Constants (exposed for reference)
    ACTIVE_STATUSES,
    INACTIVE_STATUSES,
    UNAVAILABLE_STATUSES

  };

})();