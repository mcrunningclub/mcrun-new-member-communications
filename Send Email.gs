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

  const thisSheet = GET_LITERAL_SHEET_();
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
  logMessage_(returnStatus, thisSheet, row);
}


/**
 * Sends a personalized welcome email to a new member using their information.
 *
 * @param {Object} memberInformation  Member information object (must include firstName, passUrl, email).
 * @returns {string}  Status message indicating success or error.
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function sendWelcomeEmail_(memberInformation) {
  try {
    const TEMPLATE_NAME = 'Welcome Email';
    const CLUB_EMAIL = 'mcrunningclub@ssmu.ca';
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


/**
 * Sends an updated digital pass email to a member.
 *
 * @param {Object} member  Member information object (must include firstName, passUrl, email).
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function sendUpdatedPass(member) {
  try {
    const TEMPLATE_NAME = 'Updated Pass Email';
    const CLUB_EMAIL = MCRUN_EMAIL;

    // Prepare the HTML body from the template
    const template = HtmlService.createTemplateFromFile(TEMPLATE_NAME);

    // Populate member data
    template.PASS_URL = member['passUrl'];
    template.FIRST_NAME = member['firstName'];

    // Add CID 
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
      to: member['email'],
      subject: 'Your updated digital pass',
      from: CLUB_EMAIL,
      name: CLUB_NAME,
      replyTo: CLUB_EMAIL,
      htmlBody: emailBodyHTML,
      inlineImages: inlineImages,
    };

    MailApp.sendEmail(message);
    console.log(`[NMC] Email sent successfully to ${member['email']}`);
  } catch (error) {
    console.error(`[NMC] Error sending email: ${error}`);
  }
}


/**
 * Quickly sends an updated pass email for a member at the specified row and logs the action.
 *
 * @param {number} row  Row number to target.
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function quickPassUpdate(row) {
  const sheet = GET_LITERAL_SHEET_();
  const memberData = sheet.getSheetValues(row, 1, 1,  COL_MAP.DIGITAL_PASS_URL)[0];

  sendUpdatedPass({
    'firstName' :  memberData[COL_MAP.FIRST_NAME - 1],
    'email' :  memberData[COL_MAP.EMAIL - 1],
    'passUrl' :  memberData[COL_MAP.DIGITAL_PASS_URL - 1],
  });

  logMessage_('Sent updated pass!', sheet, row);
}
