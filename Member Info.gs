/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// SHEET CONSTANTS
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


function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}


function getDraftBySubject_(subject = DRAFT_SUBJECT_LINE) {
  return GmailApp
  .getDrafts()
  .filter(
    subjectFilter_(subject)
  )[0];
}

function getDraftById_(id = DRAFT_ID) {
  return GmailApp.getDraft(id);
}


function createNewMemberCommunications(memberObj) {
  const thisSheet = GET_LITERAL_SHEET_();
  console.log('Starting execution now...');

  try {
    // Append member info
    const newRow = appendNewValues_(memberObj, thisSheet);
    console.log('Successfully imported values to row ' + newRow);

    // Create member pass 
    const passUrl = createPassFile_(memberObj);   // Get download url for member pass
    console.log('Successfully created digital pass with url:\n' + passUrl);

    // Save url of digital pass to sheet and `memberObj`
    thisSheet.getRange(newRow, COL_MAP.DIGITAL_PASS_URL).setValue(passUrl);
    memberObj['passUrl'] = passUrl;
    console.log('Successfully saved url to row ' + newRow);

    // Send welcome email and log result
    const returnMessage = sendWelcomeEmail_(memberObj);
    console.log(returnMessage);
    logMessage_(returnMessage, thisSheet, newRow);
  }
  catch(e) {
    logMessage_(e.message, thisSheet, newRow);
    throw e;
  }
}


function appendNewValues_(memberObj, thisSheet = GET_LITERAL_SHEET_()) {
  const importMap = IMPORT_MAP;
  const entries = Object.entries(memberObj)
  const valuesToAppend = Array(entries.length);

  for (let [key, value] of entries) {
    if (key in importMap) {
      let indexInSheet = importMap[key] - 1;   // Set 1-index to 0-index
      valuesToAppend[indexInSheet] = value;
    }
  }

  // Append imported values and return new row index
  const newRow = thisSheet.getLastRow() + 1;
  const colSize = entries.length;
  thisSheet.getRange(newRow, 1, 1, colSize).setValues([valuesToAppend]);
  
  return newRow;
}


function triggerUpdateAndSendPass(row = 2) {
  const thisSheet = GET_PAYMENT_LOG_SHEET_();
  const colSize = thisSheet.getLastColumn() - 1;    // ERROR_STATUS not needed

  const headerKeys = thisSheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = thisSheet.getRange(row, 1, 1, colSize).getDisplayValues()[0];
  
  // Package member information using key-values
  const updated = headerKeys.reduce(
    (obj, key, i) => (obj[toCamelCase(key)]= newMemberValues[i], obj), {}
  );

  console.log(updated);

  // Try to send email and record status
  updateAndSendPass(updated, true);

  function toCamelCase(str) {
    return str
      .toLowerCase()
      .replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
  }
}


function updateAndSendPass(statusObj, isLogged = false) {
  // STEP 1: Add to payment logs
  if (!isLogged) logPaymentStatus_(statusObj);

  // STEP 2: Get existing member data
  const literalsSheet = GET_LITERAL_SHEET_();
  const email = statusObj['email'];
  const targetRow = findRowByEmail_(email);

  const memberData = literalsSheet.getSheetValues(targetRow, 1, 1,  COL_MAP.DIGITAL_PASS_URL)[0];
 
  // STEP 3: Delete previous member pass
  const oldPassUrl = memberData[COL_MAP.DIGITAL_PASS_URL - 1];
  
  if (oldPassUrl) {
    const match = oldPassUrl.match(/\/d\/([^/]+)\/export/);
    const fileId = match[1];
    DriveApp.getFileById(fileId).setTrashed(true);
  }
  
  // STEP 4: Update fee status
  literalsSheet.getRange(targetRow, COL_MAP.FEE_STATUS).setValue(statusObj['feeStatus']);

  // STEP 5: Create new pass and store url
  const newPassUrl = createNewPass(targetRow);

  // STEP 6: Send updated pass email
  sendUpdatedPass({
    'firstName' :  memberData[COL_MAP.FIRST_NAME - 1],
    'email' :  email,
    'passUrl' : newPassUrl,
  });

  logMessage_('Sent updated pass!', literalsSheet, targetRow);
  Logger.log(`[NMC] Completed 'updateAndSendPass' and exiting`);
}


function logPaymentStatus_(status) {
  const sheet = GET_PAYMENT_LOG_SHEET_();
  const updatedRow = [];

  // Map values from the status object to the correct indexes using PAYMENT_LOG_MAP
  Object.entries(PAYMENT_LOG_MAP).forEach(([key, index]) => {
    updatedRow[index - 1] = status[key];    // Turn 1-index to 0-index
  });

  sheet.appendRow(updatedRow);
}


function findRowByEmail_(targetEmail) {
  const sheet = GET_LITERAL_SHEET_();
  const allEmail = sheet.getRange(1, COL_MAP.EMAIL, sheet.getLastRow()).getValues();
  return allEmail.findIndex(row => row[0] === targetEmail) + 1;  // 0 to 1-index
}


function logMessage_(message, thisSheet =  GET_LITERAL_SHEET_(), thisRow = thisSheet.getLastRow()) {
  // Update the status of email for new member
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss]');
  const statusRange = thisSheet.getRange(thisRow, COL_MAP.EMAIL_LOG);

  // Append status to previous value
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  statusRange.setValue(`${previousValue}${currentTime}: ${message}`);
}

