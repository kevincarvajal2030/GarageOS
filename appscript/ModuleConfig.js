/**
 * ============================================================================
 * GarageOS
 * ModuleConfig.gs
 * ----------------------------------------------------------------------------
 * Central configuration repository for every GarageOS module.
 *
 * This file contains ONLY metadata.
 *
 * No business logic.
 * No Spreadsheet operations.
 * No validations.
 * No ID generation.
 *
 * Every Engine must read configuration from this file.
 * ============================================================================
 */

const ModuleConfig = (() => {

  /**
   * --------------------------------------------------------------------------
   * Module Definitions
   * --------------------------------------------------------------------------
   */

  const MODULES = {

    "01_Customers": {

      sheetName: "01_Customers",

      entityName: "Customer",

      idPrefix: "CUS",

      counterKey: "CUSTOMER_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      /**
       * --------------------------------------------------------
       * Field Mapping
       * --------------------------------------------------------
       */

      fields: {

        CustomerID: 1,

        FirstName: 2,

        LastName: 3,

        Phone: 4,

        Email: 5,

        Address: 6,

        City: 7,

        State: 8,

        ZIPCode: 9,

        PreferredContact: 10,

        Status: 11,

        Notes: 12

      },

      /**
       * --------------------------------------------------------
       * Required Fields
       * --------------------------------------------------------
       */

      requiredFields: [

        "FirstName",

        "LastName",

        "Phone",

        "Email"

      ],

      /**
       * --------------------------------------------------------
       * Protected Fields
       * --------------------------------------------------------
       */

      protectedFields: [

        "CustomerID"

      ],

      /**
       * --------------------------------------------------------
       * Duplicate Detection
       * --------------------------------------------------------
       */

      duplicateFields: [

        "FirstName",

        "LastName",

        "Phone",

        "Email"

      ]

    },



    "02_Vehicles": {

      entityName: "Vehicle",

      idPrefix: "VEH",

      counterKey: "VEHICLE_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        VehicleID: 1,

        CustomerID: 2,

        LicensePlate: 3,

        VIN: 4,

        Make: 5,

        Model: 6,

        Year: 7,

        Color: 8,

        Mileage: 9,

        FuelType: 10,

        Transmission: 11,

        Status: 12,

        Notes: 13

      },

      requiredFields: [

        "CustomerID",

        "LicensePlate",

        "Make",

        "Model"

      ],

      protectedFields: [

        "VehicleID"

      ],

      duplicateFields: [

        "LicensePlate",

        "VIN"

      ]

    },

    "03_Work_Orders": {

      entityName: "Work Order",

      idPrefix: "WO",

      counterKey: "WORK_ORDER_ID",

      autoGenerateId: true,

      allowManualDelete: false,

      allowManualEdition: true,

      fields: {

        WorkOrderID:1,

        CustomerID:2,

        VehicleID:3,

        OpenDate:4,

        CloseDate:5,

        Status:6,

        Priority:7,

        MechanicID:8,

        Total:9,

        Notes:10

      },

      requiredFields:[

        "CustomerID",

        "VehicleID",

        "Status"

      ],

      protectedFields:[

        "WorkOrderID"

      ],

      duplicateFields:[

      ]

    },

    "04_Services": {

      entityName:"Service",

      idPrefix:"SER",

      counterKey:"SERVICE_ID",

      autoGenerateId:true,

      allowManualDelete:true,

      allowManualEdition:true,

      fields:{

        ServiceID:1,

        ServiceName:2,

        Category:3,

        LaborHours:4,

        LaborRate:5,

        Price:6,

        Status:7,

        Notes:8

      },

      requiredFields:[

        "ServiceName"

      ],

      protectedFields:[

        "ServiceID"

      ],

      duplicateFields:[

        "ServiceName"

      ]

    },

    "05_Parts": {

      entityName:"Part",

      idPrefix:"PRT",

      counterKey:"PART_ID",

      autoGenerateId:true,

      allowManualDelete:true,

      allowManualEdition:true,

      fields:{

        PartID:1,

        PartNumber:2,

        Description:3,

        SupplierID:4,

        Cost:5,

        Price:6,

        Stock:7,

        MinimumStock:8,

        Status:9,

        Notes:10

      },

      requiredFields:[

        "PartNumber",

        "Description"

      ],

      protectedFields:[

        "PartID"

      ],

      duplicateFields:[

        "PartNumber"

      ]

    },

    "06_Suppliers": {

      entityName: "Supplier",

      idPrefix: "SUP",

      counterKey: "SUPPLIER_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        SupplierID: 1,

        SupplierName: 2,

        ContactName: 3,

        Phone: 4,

        Email: 5,

        Address: 6,

        City: 7,

        State: 8,

        ZIPCode: 9,

        Status: 10,

        Notes: 11

      },

      requiredFields: [

        "SupplierName"

      ],

      protectedFields: [

        "SupplierID"

      ],

      duplicateFields: [

        "SupplierName"

      ]

    },

    "07_Purchases": {

      entityName: "Purchase",

      idPrefix: "PUR",

      counterKey: "PURCHASE_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        PurchaseID: 1,

        SupplierID: 2,

        PurchaseDate: 3,

        Total: 4,

        Status: 5,

        Notes: 6

      },

      requiredFields: [

        "SupplierID"

      ],

      protectedFields: [

        "PurchaseID"

      ],

      duplicateFields: []

    },

    "08_Payments": {

      entityName: "Payment",

      idPrefix: "PAY",

      counterKey: "PAYMENT_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        PaymentID: 1,

        WorkOrderID: 2,

        PaymentDate: 3,

        PaymentMethod: 4,

        Amount: 5,

        Status: 6,

        Notes: 7

      },

      requiredFields: [

        "WorkOrderID",

        "Amount"

      ],

      protectedFields: [

        "PaymentID"

      ],

      duplicateFields: []

    },

    "09_Mechanics": {

      entityName: "Mechanic",

      idPrefix: "MEC",

      counterKey: "MECHANIC_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        MechanicID: 1,

        FirstName: 2,

        LastName: 3,

        Phone: 4,

        Email: 5,

        Status: 6,

        Notes: 7

      },

      requiredFields: [

        "FirstName",

        "LastName"

      ],

      protectedFields: [

        "MechanicID"

      ],

      duplicateFields: [

        "FirstName",

        "LastName",

        "Phone"

      ]

    }

  };

  /**
   * Returns module configuration by sheet name.
   */
  function get(sheetName) {

    return MODULES[sheetName] || null;

  }

  /**
   * Returns all module definitions.
   */
  function getAll() {

    return MODULES;

  }

  return {

    get,

    getAll

  };

})();