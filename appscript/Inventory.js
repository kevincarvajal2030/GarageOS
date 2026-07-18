/**
 * ============================================================================
 * GARAGE OS
 * Inventory Module
 * ============================================================================
 *
 * Handles all inventory (parts) business logic.
 *
 * Responsibilities:
 *   - Create parts.
 *   - Update parts.
 *   - Delete parts.
 *   - Validate part information.
 *   - Build part records.
 *
 * Uses:
 *   - Validation.gs
 *   - Database.gs
 *   - IDs.gs
 * ============================================================================
 */


/******************************************************************************
 * PART CRUD
 ******************************************************************************/

/**
 * Creates a new part.
 *
 * @param {Object} partData
 */
function createPart(partData) {

  validatePartData(partData);

  const record = buildPartRecord(partData);

  appendRecord(
    SHEETS.PARTS,
    record
  );

}


/**
 * Updates an existing part.
 *
 * @param {string} partId
 * @param {Object} partData
 */
function updatePart(partId, partData) {

  validatePartData(partData);

  const existingRecord = getRecord(
    SHEETS.PARTS,
    partId
  );

  if (!existingRecord) {

    throw new Error("Part not found.");

  }

  const updatedRecord = [

    partId,

    partData.name,

    partData.category,

    partData.brand,

    partData.quantity,

    partData.unitPrice,

    existingRecord[6]

  ];

  updateRecord(
    SHEETS.PARTS,
    partId,
    updatedRecord
  );

}


/**
 * Deletes a part.
 *
 * @param {string} partId
 */
function deletePart(partId) {

  deleteRecord(
    SHEETS.PARTS,
    partId
  );

}


/******************************************************************************
 * PART QUERIES
 ******************************************************************************/

/**
 * Returns a part by ID.
 *
 * @param {string} partId
 * @returns {Object|null}
 */
function getPartById(partId) {

  const record = getRecord(
    SHEETS.PARTS,
    partId
  );

  if (!record) {

    return null;

  }

  return {

    id: record[0],

    name: record[1],

    category: record[2],

    brand: record[3],

    quantity: record[4],

    unitPrice: record[5],

    createdAt: record[6]

  };

}


/**
 * Determines whether a part exists.
 *
 * @param {string} partId
 * @returns {boolean}
 */
function partExists(partId) {

  return recordExists(
    SHEETS.PARTS,
    partId
  );

}


/******************************************************************************
 * PART VALIDATION
 ******************************************************************************/

/**
 * Validates part information.
 *
 * @param {Object} partData
 */
function validatePartData(partData) {

  validateObjectField(
    partData,
    "Part"
  );

  validateTextField(
    partData.name,
    "Part Name",
    {
      required: true,
      minimumLength: 2,
      maximumLength: 100
    }
  );

  validateTextField(
    partData.category,
    "Category",
    {
      required: true
    }
  );

  validateTextField(
    partData.brand,
    "Brand",
    {
      required: true
    }
  );

  validatePositiveNumberField(
    partData.quantity,
    "Quantity"
  );

  validatePositiveNumberField(
    partData.unitPrice,
    "Unit Price"
  );

}


/******************************************************************************
 * PART RECORD
 ******************************************************************************/

/**
 * Builds a part record.
 *
 * @param {Object} partData
 * @returns {Array}
 */
function buildPartRecord(partData) {

  return [

    generatePartID(),

    partData.name,

    partData.category,

    partData.brand,

    partData.quantity,

    partData.unitPrice,

    getTimestamp()

  ];

}