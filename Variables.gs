const LITERAL_SHEET_NAME = 'Literals';
const LITERAL_SHEET_ID = '1PrKth6f81Dx52bB3oPX1t55US-GnNRGve-TN4rU9Wlo';
const LITERAL_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LITERAL_SHEET_NAME);

const PAYMENT_LOG_SHEET_NAME = 'Payment Logs';
const PAYMENT_LOG_SHEET_ID = '1607079750';
const PAYMENT_LOG_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYMENT_LOG_SHEET_NAME);

// EMAIL DRAFT CONSTANTS
const DRAFT_SUBJECT_LINE = 'Here\'s your post-run report! 🙌';
const DRAFT_ID = ''; 'r-7747016114606374047';

// THIS IS AN ALTERNATE TEMPLATE (CURRENTLY NOT USED)
const WELCOME_EMAIL_TEMPLATE_ID = '13hYRAbGSBVUzPTSzUl3Vao0RQJmjhWHHWoAku09j5iI';

const TIMEZONE = getUserTimeZone_();
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';
const CLUB_NAME = 'McGill Students Running Club';

// ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
// SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
const GET_LITERAL_SHEET_ = () => {
  return (LITERAL_SHEET) ?? SpreadsheetApp.openById(LITERAL_SHEET_ID).getSheetByName(LITERAL_SHEET_NAME);
}

const GET_PAYMENT_LOG_SHEET_ = () => {
  return (PAYMENT_LOG_SHEET) ?? SpreadsheetApp.openById(LITERAL_SHEET_ID).getSheetByName(PAYMENT_LOG_SHEET_NAME);
}

// SHEET MAPPING
const COL_MAP = {
  EMAIL: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  MEMBER_ID: 4,
  MEMBER_STATUS: 5,
  FEE_STATUS: 6,
  EXPIRY_DATE: 7,
  DIGITAL_PASS_URL: 8,
  EMAIL_LOG: 9,
};

// MAPPING FROM MEMBER OBJ TO SHEET HEADER
const IMPORT_MAP = {
  'email': COL_MAP.EMAIL,
  'firstName': COL_MAP.FIRST_NAME,
  'lastName': COL_MAP.LAST_NAME,
  'memberId': COL_MAP.MEMBER_ID,
  'memberStatus': COL_MAP.MEMBER_STATUS,
  'feeStatus' : COL_MAP.FEE_STATUS,
  'expiry': COL_MAP.EXPIRY_DATE,
  'passUrl': COL_MAP.DIGITAL_PASS_URL,
}

const PAYMENT_LOG_MAP = {
  'timestamp' : 1,
  'email' : 2,
  'feeStatus' : 3,
}