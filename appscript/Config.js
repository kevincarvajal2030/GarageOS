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

  REFERENCEDATA: "10_Reference_Data",

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


COUNTERS = {

  CUSTOMER: "CUSTOMER_ID",

  VEHICLE: "VEHICLE_ID",

  WORK_ORDER: "WORK_ORDER_ID",

  SERVICE: "SERVICE_ID",

  PART: "PART_ID",

  SUPPLIER: "SUPPLIER_ID",

  PURCHASE: "PURCHASE_ID",

  PAYMENT: "PAYMENT_ID",

  MECHANIC: "MECHANIC_ID"

};


const TABLE = {

  HEADER_ROW: 6,

  FIRST_DATA_ROW: 7

};


const APP = {

  NAME: "GarageOS",

  VERSION: "1.0.1" //7-22-2026

};


/**
 * Google Drive
 */
const DRIVE = Object.freeze({

    IMAGE_EXTENSION: "png"

});

/**
 * API OPENAI
 */
const OPENAI = Object.freeze({

    CHAT_MODEL: "gpt-5.1",

    IMAGE_MODEL: "gpt-image-1"

});