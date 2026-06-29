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
function sendWelcomeEmailInRow(row) {
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
  logEmailStatus_(returnStatus, row);
}

/**
 * Sends welcome email to member using template and member info.
 * 
 * Gets member information and image blobs stored in script properties,
 * populates template, and sends email.
 * 
 * @param {Object} memberInformation  Object containing member information from Literals
 * @return {string}  "Successfully sent!" if email sent, otherwise error message
 */
function sendWelcomeEmail_(memberInformation) {
  try {
    // Name of html file (without .html extension) in the apps script project
    const TEMPLATE_NAME = 'Welcome Email';

    const CLUB_EMAIL = 'mcrunningclub@ssmu.ca';
    const SUBJECT_LINE = 'Hi from McRUN 👋';

    // Prepare the HTML body from the template
    const template = HtmlService.createTemplateFromFile(TEMPLATE_NAME);
    
    template.THIS_YEAR = new Date().getFullYear();
    template.FIRST_NAME = memberInformation['firstName'] || memberInformation['FIRST_NAME'];
    template.PASS_URL = memberInformation['passUrl'] || memberInformation['DIGITAL_PASS_URL'];

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
      to: memberInformation['email'] || memberInformation['EMAIL'],
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
 * Sends updated pass email to member using template and member info.
 * 
 * Gets member information and image blobs stored in script properties,
 * populates template, sends email, and log to console if successful or not.
 * 
 * @param {Object} memberInformation  Object containing member information from Literals
 */
function sendUpdatedPass_(memberInformation) {
  try {
    // Name of html file (without .html extension) in the apps script project
    const TEMPLATE_NAME = 'Pass Email';

    const CLUB_EMAIL = MCRUN_EMAIL;

    // Prepare the HTML body from the template
    const template = HtmlService.createTemplateFromFile(TEMPLATE_NAME);

    // Populate member data
    template.PASS_URL = memberInformation['passUrl'];
    template.FIRST_NAME = memberInformation['firstName'];

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
      to: memberInformation['email'],
      subject: 'Your updated digital pass',
      from: CLUB_EMAIL,
      name: CLUB_NAME,
      replyTo: CLUB_EMAIL,
      htmlBody: emailBodyHTML,
      inlineImages: inlineImages,
    };

    MailApp.sendEmail(message);
    console.log(`[NMC] Email sent successfully to ${memberInformation['email']}`);
  } catch (error) {
    console.error(`[NMC] Error sending email: ${error}`);
  }
}

/**
 * Update pass using member information from given row in Literals sheet,
 * and sends an email with the new pass.
 * 
 * @param {number} row  Row of member to update pass for. Defaults to 15 (dunno why)
 */
function quickPassUpdate(row = 15) {
  const sheet = GET_LITERAL_SHEET_();
  createPassFromRow(row);

  const memberData = sheet.getSheetValues(row, 1, 1,  LITERALS.DIGITAL_PASS_URL)[0];

  sendUpdatedPass_({
    'firstName' :  memberData[LITERALS.FIRST_NAME - 1],
    'email' :  memberData[LITERALS.EMAIL - 1],
    'passUrl' :  memberData[LITERALS.DIGITAL_PASS_URL - 1],
  });

  logEmailStatus_('Sent updated pass!', row);
}

/**
 * Sends new pass to member from given row in Payment Logs sheet.
 * 
 * @param {number} row  Row of member to update pass for. Defaults to 2 (???)
 */
function triggerUpdateAndSendPass(row = 2) {
  const thisSheet = GET_PAYMENT_LOG_SHEET_();
  const colSize = thisSheet.getLastColumn() - 1;    // ERROR_STATUS not needed

  const headerKeys = thisSheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = thisSheet.getRange(row, 1, 1, colSize).getDisplayValues()[0];
  
  // Package member information using key-values
  const updated = headerKeys.reduce(
    (obj, key, i) => (obj[toCamelCase_(key)]= newMemberValues[i], obj), {}
  );

  console.log(updated);

  // Try to send email and record status
  updateAndSendPass_(updated, true);
}


/**
 * Sends pass given payment status object.
 * 
 * Finds existing member data from literals sheet, deletes old pass
 * and creates new one, and sends email to member
 * 
 * @param {Object} statusObj  Payment status, including 'email' and 'fee status'
 * @param {boolean} isLogged  Whether the status has been added to Payment Logs sheet.
 *                              Determines whether to add it or not. Default false.
 */
function updateAndSendPass_(statusObj, isLogged = false) {
  // STEP 1: Add to payment logs
  if (!isLogged) logPaymentStatus_(statusObj);

  // STEP 2: Get existing member data
  const literalsSheet = GET_LITERAL_SHEET_();
  const email = statusObj['email'];
  const targetRow = findRowByEmail_(email);

  const memberData = literalsSheet.getSheetValues(targetRow, 1, 1,  LITERALS.DIGITAL_PASS_URL)[0];
 
  // STEP 3: Delete previous member pass
  const oldPassUrl = memberData[LITERALS.DIGITAL_PASS_URL - 1];
  
  if (oldPassUrl) {
    const match = oldPassUrl.match(/\/d\/([^/]+)\/export/);
    const fileId = match[1];
    DriveApp.getFileById(fileId).setTrashed(true);
  }
  
  // STEP 4: Update fee status
  literalsSheet.getRange(targetRow, LITERALS.FEE_STATUS).setValue(statusObj['feeStatus']);

  // STEP 5: Create new pass and store url
  const newPassUrl = createPassFromRow(targetRow);

  // STEP 6: Send updated pass email
  sendUpdatedPass_({
    'firstName' :  memberData[LITERALS.FIRST_NAME - 1],
    'email' :  email,
    'passUrl' : newPassUrl,
  });

  logEmailStatus_('Sent updated pass!', targetRow);
  Logger.log(`[NMC] Completed 'updateAndSendPass' and exiting`);
}


/**
 * Sends email from template in drafts using member information.
 * 
 * Finds draft using subject and gets template from it, then fills
 * in member information and creates a new email to send.
 * Throws error and logs it in console if error occurs during sending. 
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {Object} memberInformation  Information to populate email draft
 * @param {string} draftSubject  Subject line of the email draft to use as template
 * @return {{message:string, isError:bool}}  Status of sending email.
*/
function sendEmail_(memberInformation, draftSubject) {
  // Gets the draft Gmail message to use as a template
  const emailTemplate = getTemplateFromDraft_(draftSubject);

  try {
    const memberEmail = memberInformation['EMAIL'];
    const msgObj = fillInTemplate(emailTemplate.message, memberInformation);

    //DriveApp.createFile('TestFile3b', msgObj.html);

    MailApp.sendEmail(
      'andrey.gonzalez@mail.mcgill.ca',
      msgObj.subject,
      msgObj.text,
      {
        htmlBody: msgObj.html,
        from: 'mcrunningclub@ssmu.ca',
        name: 'McGill Students Running Club',
        replyTo: 'mcrunningclub@ssmu.ca',
        attachments: emailTemplate.attachments,
        inlineImages: emailTemplate.inlineImages
      }
    );

  } catch(e) {
    // Log and return error
    console.log(`(sendEmail) ${e.message}`);
    throw new Error(e);
  }
  // Return success message
  return {message: 'Sent!', isError : false};
}