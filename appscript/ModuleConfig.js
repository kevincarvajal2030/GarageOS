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

      primaryKey: "CustomerID",

      foreignKeys: [],

      protectedFields: [],

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

      duplicateFields: [

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
        CustomerName: 2,
        CustomerID: 3,
        LicensePlate: 4,
        Make: 5,
        Model: 6,
        Year: 7,
        Transmission: 8,
        Color: 9,
        FuelType: 10,
        Status: 11,
        Notes: 12

      },

      primaryKey: "VehicleID",

      foreignKeys: [

        "CustomerID"

      ],

      protectedFields: [],

      requiredFields: [

        "CustomerName",
        "CustomerID",
        "LicensePlate",
        "Year",
        "Make"

      ],

      duplicateFields: [

        "LicensePlate"

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

        WorkOrderID: 1,
        CustomerName: 2,
        CustomerID: 3,
        VehicleName: 4,
        VehicleID: 5,
        MechanicName: 6,
        MechanicID: 7,
        OpenDate: 8,
        Status: 9,
        Priority: 10,
        Mileage: 11,
        Complaint: 12,
        Diagnosis: 13,
        LaborCost: 14,
        PartCost: 15,
        TotalCost: 16, 
        CompletionDate: 17

      },

      primaryKey: "WorkOrderID",

      foreignKeys: [
        "CustomerID",
        "VehicleID"
      ],

      protectedFields: [],

      requiredFields: [

        "CustomerName",
        "CustomerID",
        "VehicleName",
        "VehicleID",
        "MechanicName",
        "MechanicID",
        "OpenDate",

      ],

      duplicateFields: [

      ]

    },

    "04_Services": {

      entityName: "Service",

      idPrefix: "SER",

      counterKey: "SERVICE_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        ServiceID: 1,
        ServiceName: 2,
        Category: 3,
        LaborHours: 4,
        LaborRate: 5,
        Price: 6,
        Status: 7,
        Notes: 8

      },

      primaryKey: "ServiceID",

      foreignKeys: [
        
      ],

      protectedFields: [],

      requiredFields: [

        "ServiceName"

      ],


      duplicateFields: [

        "ServiceName"

      ]

    },

    "05_Parts": {

      entityName: "Part",

      idPrefix: "PRT",

      counterKey: "PART_ID",

      autoGenerateId: true,

      allowManualDelete: true,

      allowManualEdition: true,

      fields: {

        PartID: 1,
        PartNumber: 2,
        Description: 3,
        SupplierID: 4,
        Cost: 5,
        Price: 6,
        Stock: 7,
        MinimumStock: 8,
        Status: 9,
        Notes: 10

      },

      primaryKey: "PartID",

      foreignKeys: [
        
      ],

      protectedFields: [],

      requiredFields: [

        "PartNumber",
        "Description"

      ],


      duplicateFields: [

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
        CompanyName: 2,
        ContactPerson: 3,
        Phone: 4,
        Email: 5,
        Address: 6,
        City: 7,
        State: 8,
        Status: 9

      },

      primaryKey: "SupplierID",

      foreignKeys: [
        
      ],

      protectedFields: [],

      requiredFields: [

        "CompanyName",
        "ContactPerson",
        "Phone",
        "Email",

      ],


      duplicateFields: [

        "Phone",
        "Email"

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

      primaryKey: "PurchaseID",

      foreignKeys: [
        
      ],

      protectedFields: [],

      requiredFields: [

        "SupplierID"

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

      primaryKey: "PaymentID",

      foreignKeys: [
        
      ],

      protectedFields: [],

      requiredFields: [

        "WorkOrderID",

        "Amount"

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
        HireDate: 6,
        Specialty: 7,
        Status: 8
      },

      primaryKey: "MechanicID",

      foreignKeys: [],

      protectedFields: [],

      requiredFields: [

        "FirstName",

        "LastName"

      ],

      duplicateFields: [

        "Phone",
        "Email"

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