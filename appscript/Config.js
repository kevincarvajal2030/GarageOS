/**
 * ==========================================
 * GarageOS - Global Configuration | Config.gs
 * ==========================================
 */

/**
 * Sheet Names
 */
const SHEETS = {

  DASHBOARD: "00_Dashboard",

  CUSTOMERS: "01_Customers",

  VEHICLES: "02_Vehicles",

  WORK_ORDERS: "03_Work_Orders",

  SERVICES: "04_Services",

  PARTS: "05_Parts",

  SUPPLIERS: "06_Suppliers",

  PURCHASES: "07_Purchases",

  PAYMENTS: "08_Payments",

  MECHANICS: "09_Mechanics",

  SETTINGS: "10_Settings",

  SYSTEM: "12_System"

};


const ID_PREFIX = {

  CUSTOMER: "CUS",

  VEHICLE: "VEH",

  WORK_ORDER: "WO",

  SERVICE: "SER",

  PART: "PRT",

  SUPPLIER: "SUP",

  PURCHASE: "PUR",

  PAYMENT: "PAY",

  MECHANIC: "MEC"

};


const COUNTERS = {

  CUSTOMER: "NextCustomerID",

  VEHICLE: "NextVehicleID",

  WORK_ORDER: "NextWorkOrderID",

  SERVICE: "NextServiceID",

  PART: "NextPartID",

  SUPPLIER: "NextSupplierID",

  PURCHASE: "NextPurchaseID",

  PAYMENT: "NextPaymentID",

  MECHANIC: "NextMechanicID"

};


const TABLE = {

  HEADER_ROW: 6,

  FIRST_DATA_ROW: 7

};


const APP = {

  NAME: "GarageOS",

  VERSION: "1.0.0"

};