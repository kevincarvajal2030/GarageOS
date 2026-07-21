
/******************************************************************************
 * CUSTOMER.GS
 ******************************************************************************/

/**
 * Returns a customer by ID.
 *
 * @param {string} customerId
 * @returns {Object|null}
 */
function getCustomerById(customerId) {

  const record = getRecord(
    SHEETS.CUSTOMERS,
    customerId
  );

  if (!record) {

    return null;

  }

  return {

    id: record[0],

    firstName: record[1],

    lastName: record[2],

    phone: record[3],

    email: record[4],

    address: record[5],

    createdAt: record[6]

  };

}


/**
 * Determines whether a customer exists.
 *
 * @param {string} customerId
 * @returns {boolean}
 */
function customerExists(customerId) {

  return recordExists(
    SHEETS.CUSTOMERS,
    customerId
  );

}


/******************************************************************************
 * CUSTOMER VALIDATION
 ******************************************************************************/

/**
 * Validates all customer information before saving.
 *
 * @param {Object} customerData
 */
function validateCustomerData(customerData) {

  validateObjectField(customerData, "Customer");

  validateTextField(
    customerData.firstName,
    "First Name",
    {
      required: true,
      minimumLength: 2,
      maximumLength: 50
    }
  );

  validateTextField(
    customerData.lastName,
    "Last Name",
    {
      required: true,
      minimumLength: 2,
      maximumLength: 50
    }
  );

  validatePhoneField(
    customerData.phone,
    "Phone"
  );

  validateEmailField(
    customerData.email,
    "Email"
  );

  validateTextField(
    customerData.address,
    "Address",
    {
      required: false,
      maximumLength: 200
    }
  );

}
