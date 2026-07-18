/**
 * ============================================================================
 * GARAGE OS
 * Customer UI Module
 * ============================================================================
 *
 * Handles all Customer user interface operations.
 *
 * Responsibilities:
 *   - Open Customer Form
 *   - Load Customer Data
 *   - Submit Customer Form
 *   - Refresh Customer View
 *
 * Business Logic remains in:
 *   Customers.gs
 * ============================================================================
 */


/******************************************************************************
 * CUSTOMER FORM
 ******************************************************************************/

/**
 * Opens the Customer Form.
 *
 * @param {Object} options
 *
 * Example:
 * showCustomerForm({
 *   mode: "create"
 * });
 *
 * showCustomerForm({
 *   mode: "edit",
 *   customerId: "CUS-000001"
 * });
 */
function showCustomerForm(options) {

  options = options || {};

  const template = HtmlService.createTemplateFromFile("CustomerForm");

  template.options = {
    mode: options.mode || "create",
    customerId: options.customerId || null
  };

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(620);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      options.mode === "edit"
        ? "Edit Customer"
        : "New Customer"
    );

}


/******************************************************************************
 * LOAD CUSTOMER
 ******************************************************************************/

/**
 * Returns customer information for the HTML form.
 *
 * @param {string} customerId
 * @returns {Object}
 */
function loadCustomer(customerId) {

  return getCustomerById(customerId);

}


/******************************************************************************
 * SUBMIT CUSTOMER FORM
 ******************************************************************************/

/**
 * Creates or updates a customer depending on form mode.
 *
 * @param {Object} formData
 * @returns {Object}
 */
function submitCustomerForm(formData) {

  try {

    if (formData.mode === "edit") {

      updateCustomer(
        formData.customerId,
        formData
      );

    } else {

      createCustomer(formData);

    }

    refreshCustomersView();

    return {

      success: true,

      message:
        formData.mode === "edit"
          ? "Customer updated successfully."
          : "Customer created successfully."

    };

  }

  catch (error) {

    return {

      success: false,

      message: error.message

    };

  }

}


/******************************************************************************
 * CUSTOMER VIEW
 ******************************************************************************/

/**
 * Refreshes the Customers sheet.
 */
function refreshCustomersView() {

  SpreadsheetApp.flush();

}


/******************************************************************************
 * DIALOG
 ******************************************************************************/

/**
 * Closes the active dialog.
 */
function closeDialog() {

  SpreadsheetApp
    .getUi()
    .alert("This function should only be called from HTML.");

}



function openNewCustomer() {

  showCustomerForm({
    mode: "create"
  });

}