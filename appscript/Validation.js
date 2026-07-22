//Validation.gs

function throwValidationError(message) {
  throw new Error(`Validation Error: ${message}`);
}


function isEmpty(value) {

  if (value === null || value === undefined) {
    return true;
  }

  if (typeof value === "string") {
    return value.trim() === "";
  }

  return false;

}


function validateRequiredField(value, fieldName) {

  if (isEmpty(value)) {
    throwValidationError(`${fieldName} is required.`);
  }

}


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


function validateNumberField(value, fieldName) {

  validateRequiredField(value, fieldName);

  if (typeof value !== "number" || Number.isNaN(value)) {
    throwValidationError(`${fieldName} must be a valid number.`);
  }

}


function validateEmailField(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

  const emailPattern =
    /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

  if (!emailPattern.test(value.trim())) {
    throwValidationError(`${fieldName} is not a valid email address.`);
  }

}


function validatePhoneField(value, fieldName) {

  validateRequiredField(value, fieldName);
  validateTextField(value, fieldName);

  const phonePattern = /^[0-9+\-() ]+$/;

  if (!phonePattern.test(value.trim())) {
    throwValidationError(`${fieldName} is not a valid phone number.`);
  }

}

