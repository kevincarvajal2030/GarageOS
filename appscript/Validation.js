//Validation.gs

/**
 * ============================================================================
 * GARAGE OS
 * Validation Library
 * ============================================================================
 *
 * Centralized validation library used across the entire GarageOS application.
 *
 * Responsibilities:
 *   - Validate incoming data.
 *   - Enforce business validation rules.
 *   - Throw standardized validation errors.
 *
 * This module NEVER:
 *   - Reads Google Sheets.
 *   - Writes Google Sheets.
 *   - Generates IDs.
 *   - Performs CRUD operations.
 *
 * Every module in the system (Customers, Vehicles, WorkOrders, Inventory,
 * Payments, etc.) should use this library before saving or updating data.
 *
 * ============================================================================
 */


/******************************************************************************
 * ERROR HANDLING
 ******************************************************************************/

/**
 * Throws a standardized validation error.
 *
 * @param {string} message Error description.
 * @throws {Error}
 */
function throwValidationError(message) {
  throw new Error(`Validation Error: ${message}`);
}


/******************************************************************************
 * GENERAL UTILITIES
 ******************************************************************************/

/**
 * Determines whether a value should be considered empty.
 *
 * Empty values include:
 *   - null
 *   - undefined
 *   - empty strings
 *   - strings containing only whitespace
 *
 * Numbers, booleans and Date objects are never considered empty.
 *
 * @param {*} value Value to evaluate.
 * @returns {boolean}
 */
function isEmpty(value) {

  if (value === null || value === undefined) {
    return true;
  }

  if (typeof value === "string") {
    return value.trim() === "";
  }

  return false;

}


/**
 * Determines whether a value contains useful information.
 *
 * @param {*} value Value to evaluate.
 * @returns {boolean}
 */
function hasValue(value) {
  return !isEmpty(value);
}


/******************************************************************************
 * REQUIRED VALIDATION
 ******************************************************************************/

/**
 * Validates that a required value exists.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateRequiredField(value, fieldName) {

  if (isEmpty(value)) {
    throwValidationError(`${fieldName} is required.`);
  }

}


/******************************************************************************
 * TEXT VALIDATION
 ******************************************************************************/

/**
 * Validates text values.
 *
 * Available options:
 *
 * required
 * minimumLength
 * maximumLength
 * trim
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name.
 * @param {Object} [options={}] Validation options.
 */
function validateTextField(value, fieldName, options = {}) {

  const settings = {
    required: false,
    minimumLength: null,
    maximumLength: null,
    trim: true,
    ...options
  };

  if (settings.required) {
    validateRequiredField(value, fieldName);
  }

  if (isEmpty(value)) {
    return;
  }

  if (typeof value !== "string") {
    throwValidationError(`${fieldName} must be text.`);
  }

  const text = settings.trim
    ? value.trim()
    : value;

  if (
    settings.minimumLength !== null &&
    text.length < settings.minimumLength
  ) {
    throwValidationError(
      `${fieldName} must contain at least ${settings.minimumLength} characters.`
    );
  }

  if (
    settings.maximumLength !== null &&
    text.length > settings.maximumLength
  ) {
    throwValidationError(
      `${fieldName} cannot exceed ${settings.maximumLength} characters.`
    );
  }

}


/******************************************************************************
 * NUMBER VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is a valid number.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateNumberField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (typeof value !== "number" || Number.isNaN(value)) {
    throwValidationError(`${fieldName} must be a valid number.`);
  }

}


/**
 * Validates that a number is greater than zero.
 *
 * @param {number} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validatePositiveNumberField(value, fieldName) {

  validateNumberField(value, fieldName);

  if (value <= 0) {
    throwValidationError(`${fieldName} must be greater than zero.`);
  }

}


/******************************************************************************
 * BOOLEAN VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is a boolean.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateBooleanField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (typeof value !== "boolean") {
    throwValidationError(`${fieldName} must be true or false.`);
  }

}


/******************************************************************************
 * DATE VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is a valid Date object.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateDateField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    throwValidationError(`${fieldName} must be a valid date.`);
  }

}


/******************************************************************************
 * COLLECTION VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is an array.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateArrayField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (!Array.isArray(value)) {
    throwValidationError(`${fieldName} must be an array.`);
  }

}


/**
 * Validates that a value is an object.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateObjectField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (
    typeof value !== "object" ||
    value === null ||
    Array.isArray(value)
  ) {
    throwValidationError(`${fieldName} must be an object.`);
  }

}

/******************************************************************************
 * EMAIL VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is a valid email address.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateEmailField(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

  const emailPattern =
    /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

  if (!emailPattern.test(value.trim())) {
    throwValidationError(`${fieldName} is not a valid email address.`);
  }

}


/******************************************************************************
 * PHONE VALIDATION
 ******************************************************************************/

/**
 * Validates that a phone number contains only valid characters.
 *
 * Accepted characters:
 * Numbers
 * Spaces
 * Parentheses
 * Hyphen
 * Plus sign
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validatePhoneField(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

  const phonePattern = /^[0-9+\-() ]+$/;

  if (!phonePattern.test(value.trim())) {
    throwValidationError(`${fieldName} is not a valid phone number.`);
  }

}


/******************************************************************************
 * VALUE VALIDATION
 ******************************************************************************/

/**
 * Validates that a value is greater than or equal to a minimum value.
 *
 * @param {number} value Value to validate.
 * @param {number} minimumValue Minimum allowed value.
 * @param {string} fieldName Display name of the field.
 */
function validateMinimumFieldValue(value, minimumValue, fieldName) {

  validateNumberField(value, fieldName);

  if (value < minimumValue) {
    throwValidationError(
      `${fieldName} must be greater than or equal to ${minimumValue}.`
    );
  }

}


/**
 * Validates that a value is less than or equal to a maximum value.
 *
 * @param {number} value Value to validate.
 * @param {number} maximumValue Maximum allowed value.
 * @param {string} fieldName Display name of the field.
 */
function validateMaximumFieldValue(value, maximumValue, fieldName) {

  validateNumberField(value, fieldName);

  if (value > maximumValue) {
    throwValidationError(
      `${fieldName} must be less than or equal to ${maximumValue}.`
    );
  }

}


/**
 * Validates that a value exists in a list of allowed values.
 *
 * @param {*} value Value to validate.
 * @param {Array} allowedValues List of allowed values.
 * @param {string} fieldName Display name of the field.
 */
function validateFieldValueInList(value, allowedValues, fieldName) {

  validateRequiredField(value, fieldName);

  if (!allowedValues.includes(value)) {
    throwValidationError(
      `${fieldName} contains an invalid value.`
    );
  }

}

