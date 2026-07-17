/******************************************************************************
 * VEHICLE VALIDATION
 ******************************************************************************/

/**
 * Validates a vehicle VIN.
 *
 * Standard VIN length is 17 characters.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateVehicleVin(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

  const vin = value.trim().toUpperCase();

  if (vin.length !== 17) {
    throwValidationError(`${fieldName} must contain exactly 17 characters.`);
  }

}


/**
 * Validates a vehicle license plate.
 *
 * Only checks that a value exists.
 * Country-specific validation rules can be added later.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateLicensePlate(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

}


/**
 * Validates a vehicle manufacturing year.
 *
 * @param {*} value Value to validate.
 * @param {string} fieldName Display name of the field.
 */
function validateVehicleYear(value, fieldName) {

  validateNumberField(value, fieldName);

  const currentYear = new Date().getFullYear() + 1;

  if (value < 1900 || value > currentYear) {
    throwValidationError(
      `${fieldName} must be between 1900 and ${currentYear}.`
    );
  }

}
