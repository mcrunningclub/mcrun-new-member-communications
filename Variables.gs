/**
 * Name of Literals sheet
 * @const {string}
 */
const LITERAL_SHEET_NAME = 'Literals';

/**
 * ID of New member comms spreadsheet
 * @const {string}
 */
const SPREADSHEET_ID = '1PrKth6f81Dx52bB3oPX1t55US-GnNRGve-TN4rU9Wlo';

/**
 * Spreadsheet object of Literals sheet
 * @const {SpreadsheetApp.Sheet}
 */
const LITERAL_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LITERAL_SHEET_NAME);

/**
 * Name of Payment Logs sheet
 * @const {string}
 */
const PAYMENT_LOG_SHEET_NAME = 'Payment Logs';

/**
 * Spreadsheet object of Payment Logs sheet
 * @const {SpreadsheetApp.Sheet}
 */
const PAYMENT_LOG_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYMENT_LOG_SHEET_NAME);

/**
 * ID of email draft with welcome email template
 * @const {string}
 */
const WELCOME_EMAIL_DRAFT_ID = 'r-7747016114606374047';

/**
 * Timezone of script
 * @const {string}
 */
const TIMEZONE = getUserTimeZone_();

/**
 * Club email
 * @const {string}
 */
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';

/**
 * Club name to include in "From" field for emails
 * @const {string}
 */
const CLUB_NAME = 'McGill Students Running Club';

/**
 * Gets Literals sheet using spreadsheet ID if needed.
 * 
 * ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
 * SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
 * 
 * @return {SpreadsheetApp.Sheet} Literals sheet object
 */
const GET_LITERAL_SHEET_ = () => {
  return (LITERAL_SHEET) ?? SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LITERAL_SHEET_NAME);
}

/**
 * Gets Payment Logs sheet using spreadsheet ID if needed.
 * 
 * ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
 * SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
 * 
 * @return {SpreadsheetApp.Sheet}  Payment Logs sheet object
 */
const GET_PAYMENT_LOG_SHEET_ = () => {
  return (PAYMENT_LOG_SHEET) ?? SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PAYMENT_LOG_SHEET_NAME);
}

/**
 * Mapping column letters to numbers
 * @const {Object}
 */
const COL = {
  A: 1,
  B: 2,
  C: 3,
  D: 4,
  E: 5,
  F: 6,
  G: 7,
  H: 8,
  I: 9,
  J: 10,
  K: 11
}

/**
 * Mapping columns in Literals sheet
 * @const {Object}
 */
const LITERALS = {
  EMAIL: COL.A,
  FIRST_NAME: COL.B,
  LAST_NAME: COL.C,
  MEMBER_ID: COL.D,
  MEMBER_STATUS: COL.E,
  FEE_STATUS: COL.F,
  EXPIRY_DATE: COL.G,
  DIGITAL_PASS_URL: COL.H,
  EMAIL_LOG: COL.I,
};

/**
 * Mapping from keys in import object (from Membership Registry)
 * to columns in Literals sheet
 * @const {Object}
 */
const IMPORT_MAP = {
  'email': LITERALS.EMAIL,
  'firstName': LITERALS.FIRST_NAME,
  'lastName': LITERALS.LAST_NAME,
  'memberId': LITERALS.MEMBER_ID,
  'memberStatus': LITERALS.MEMBER_STATUS,
  'feeStatus' : LITERALS.FEE_STATUS,
  'expiry': LITERALS.EXPIRY_DATE,
  'passUrl': LITERALS.DIGITAL_PASS_URL,
}

/**
 * Mapping from fields in payment log object to columns in Payment Log sheet (?)
 * @const {Object}
 */
const PAYMENT_LOG_MAP = {
  'timestamp' : 1,
  'email' : 2,
  'feeStatus' : 3,
}