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

// EMAIL DRAFT CONSTANTS
const DRAFT_SUBJECT_LINE = 'Welcome to McRun!';
const DRAFT_ID = 'r-7747016114606374047';

// THIS IS AN ALTERNATE TEMPLATE (CURRENTLY NOT USED)
const WELCOME_EMAIL_TEMPLATE_ID = '13hYRAbGSBVUzPTSzUl3Vao0RQJmjhWHHWoAku09j5iI';

const TIMEZONE = getUserTimeZone_();

// ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
// SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
const GET_LITERAL_SHEET = () => {
  return (LITERAL_SHEET) ? (LITERAL_SHEET) : SpreadsheetApp.openById(LITERAL_SHEET_ID).getSheetByName(LITERAL_SHEET_NAME);
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
  const thisSheet = GET_LITERAL_SHEET();
  console.log('Starting execution now...');

  try {
    // Append member info
    const newRow = appendNewValues(memberObj, thisSheet);
    console.log('Successfully imported values to row ' + newRow);

    // Create member pass 
    const passUrl = createPassFile(memberObj);   // Get download url for member pass
    console.log('Successfully created digital pass with url:\n' + passUrl);

    // Save url of digital pass to sheet and `memberObj`
    thisSheet.getRange(newRow, COL_MAP.DIGITAL_PASS_URL).setValue(passUrl);
    memberObj['passUrl'] = passUrl;
    console.log('Successfully saved url to row ' + newRow);

    // Send welcome email and log result
    const returnMessage = sendWelcomeEmail_(memberObj);
    console.log(returnMessage);
    logMessage(returnMessage, thisSheet, newRow);
  }
  catch(e) {
    logMessage(e.message, thisSheet, newRow);
    throw e;
  }
}


function appendNewValues(memberObj, thisSheet = GET_LITERAL_SHEET()) {
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


function logMessage(message, thisSheet =  GET_LITERAL_SHEET(), thisRow = thisSheet.getLastRow()) {
  // Update the status of email for new member
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss]');
  const statusRange = thisSheet.getRange(thisRow, COL_MAP.EMAIL_LOG);

  // Append status to previous value
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  statusRange.setValue(`${previousValue}${currentTime}: ${message}`);
}


/**
 * Sends email using member information in `row`.
 * Logs email status in column `EMAIL_STATUS`
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {integer} row  Row to target for information
 * @error  Error status of sending email.
*/

function sendWelcomeEmailInRow(row = LITERAL_SHEET.getLastRow()) {
  if(getCurrentUserEmail_() !== 'mcrunningclub@ssmu.ca') {
    throw new Error('Wrong email. Please try using the club\'s account');
  }

  const thisSheet = GET_LITERAL_SHEET();
  const colSize = thisSheet.getLastColumn() - 1;    // ERROR_STATUS no needed

  const headerKeys = thisSheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = thisSheet.getSheetValues(row, 1, 1, colSize)[0];

  const keySize = headerKeys.length;
  
  if (keySize !== newMemberValues.length) {
    throw new Error(`Expected ${keySize} for newMemberValues.length - Received ${newMemberValues.length}`);
  }

  // Package member information using key-values
  const memberInformation = headerKeys.reduce(
    (obj, key, i) => (obj[key]= newMemberValues[i], obj), {}
  );

  // Try to send email and record status
  const returnStatus = sendWelcomeEmail_(memberInformation);
  logMessage(returnStatus, thisSheet, row);
}


function sendWelcomeEmail_(memberInformation) {
  try {
    const TEMPLATE_NAME = 'Welcome Email';
    const CLUB_EMAIL = 'mcrunningclub@ssmu.ca';
    const CLUB_NAME = 'McGill Students Running Club';
    const SUBJECT_LINE = 'Hi from McRUN 👋';

    // Prepare the HTML body from the template
    const template = HtmlService.createTemplateFromFile(TEMPLATE_NAME);
    
    template.THIS_YEAR = new Date().getFullYear();
    template.FIRST_NAME = memberInformation['firstName'];
    template.PASS_URL = memberInformation['passUrl'];

    template.LINKTREE_CID = 'linktreeLogo';
    template.HEADER_CID = 'emailHeader';
    template.STRAVA_CID = 'stravaLogo';

    // Returns string content from populated html template
    const emailBodyHTML = template.evaluate().getContent();

    // Retrieve cached blobs
    const inlineImages = {
      emailHeader: getBlobFromProperties_('emailHeaderBlob'),
      linktreeLogo: getBlobFromProperties_('linktreeLogoBlob'),
      stravaLogo: getBlobFromProperties_('stravaLogoBlob'),
    };

    // Create message object
    const message = {
      to: memberInformation['email'],
      subject: SUBJECT_LINE,
      from: CLUB_EMAIL,
      name: CLUB_NAME,
      replyTo: CLUB_EMAIL,
      htmlBody: emailBodyHTML,
      inlineImages: inlineImages,
    };

    MailApp.sendEmail(message);
    return 'Successfully sent!';  // Return success message

  } catch(e) {
    return e.message;   // Return error message to log
  }
}


function sendSamosaEmailFromHTML(recipient, subject) {
  try {
    const TEMPLATE_NAME = 'Samosa Email';
    const CLUB_EMAIL = 'mcrunningclub@ssmu.ca';
    const CLUB_NAME = 'McGill Students Running Club';

    // Prepare the HTML body from the template
    const template = HtmlService.createTemplateFromFile(TEMPLATE_NAME);

    // Populate placeholders
    template.LINKTREE_CID = 'linktreeLogo';
    template.HEADER_CID = 'emailHeader';
    template.STRAVA_CID = 'stravaLogo';
    template.REGISTRATION_LINK = 'https://mcgill.ca/x/i4T';
    template.THIS_YEAR = new Date().getFullYear();

    // Returns string content from populated html template
    const emailBodyHTML = template.evaluate().getContent();

    // Retrieve cached blobs
    const inlineImages = {
      emailHeader: getBlobFromProperties_('emailHeaderBlob'),
      linktreeLogo: getBlobFromProperties_('linktreeLogoBlob'),
      stravaLogo: getBlobFromProperties_('stravaLogoBlob'),
      
      // DriveApp call too expensive, better to cache in store
      //emailHeader : DriveApp.getFileById('1ctHsQstsoHVyCH7XcbkUNjPEka9zV9L6').getBlob().setName('emailHeaderBlob'),
    };

    // Create message object
    const message = {
      to: recipient,
      subject: subject,
      from: CLUB_EMAIL,
      name: CLUB_NAME,
      replyTo: CLUB_EMAIL,
      htmlBody: emailBodyHTML,
      inlineImages: inlineImages,
    };

    MailApp.sendEmail(message);
    console.log(`Email sent successfully to ${recipient}`);
  } catch (error) {
    console.error(`Error sending email: ${error}`);
  }
}
