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
    const CLUB_EMAIL = MCRUN_EMAIL;
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
