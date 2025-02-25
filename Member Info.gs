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
const SHEET_NAME = 'Literals';
const SHEET_ID = '1PrKth6f81Dx52bB3oPX1t55US-GnNRGve-TN4rU9Wlo';
const LITERAL_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

// EMAIL DRAFT CONSTANTS
const DRAFT_SUBJECT_LINE = 'Welcome to McRun!';
const DRAFT_ID = 'r-7747016114606374047';

// THIS IS AN ALTERNATE TEMPLATE (CURRENTLY NOT USED)
const WELCOME_EMAIL_TEMPLATE_ID = '13hYRAbGSBVUzPTSzUl3Vao0RQJmjhWHHWoAku09j5iI';

const TIMEZONE = getUserTimeZone_();

// SHEET MAPPING
const COL_MAP = {
  EMAIL: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  DIGITAL_PASS_URL: 4,
  EXPIRY_DATE: 5,
  IS_FEE_PAID: 6,
  PAYMENT_LINK: 7,
  EMAIL_STATUS: 8,
};

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}


function getDraftBySubject(subject = DRAFT_SUBJECT_LINE) {
  return GmailApp
  .getDrafts()
  .filter(
    subjectFilter_(subject)
  )[0];
}

function getDraftById(id = DRAFT_ID) {
  return GmailApp.getDraft(id);
}


function showFiles() {
  const files = DriveApp.getRootFolder().getFiles();

  while (files.hasNext()) {
    let file = files.next();
    console.log(file.getName());
  }

  // Testing access to template document (NOT EMAIL DRAFT)
  const template = DriveApp.getFileById(WELCOME_EMAIL_TEMPLATE_ID);
  console.log(template.getName());
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
  const sheet = LITERAL_SHEET;
  const colSize = sheet.getLastColumn() - 1;    // ERROR_STATUS no needed

  const headerKeys = sheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = sheet.getSheetValues(row, 1, 1, colSize)[0];

  const keySize = headerKeys.length;
  
  if (keySize !== newMemberValues.length) {
    throw new Error(`Expected ${keySize} for newMemberValues.length - Received ${newMemberValues.length}`);
  }

  // Package member information using key-values
  const memberInformation = headerKeys.reduce(
    (obj, key, i) => (obj[key]= newMemberValues[i], obj), {}
  );

  // Try to send email and record status
  try {
    const returnStatus = sendWelcomeEmail_(memberInformation);

    // Update the status of email for new member
    const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss]');
    const statusRange = sheet.getRange(row, COL_MAP.EMAIL_STATUS);

    // Append status to previous value
    const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
    statusRange.setValue(`${previousValue}${currentTime}: ${returnStatus.message}`);
  }
  catch (e) {
    throw new Error (e);
  }
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
    template.PASS_URL = memberInformation.DIGITAL_PASS_URL;
    template.FIRST_NAME = memberInformation.FIRST_NAME;

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
      to: memberInformation.EMAIL,
      subject: SUBJECT_LINE,
      from: CLUB_EMAIL,
      name: CLUB_NAME,
      replyTo: CLUB_EMAIL,
      htmlBody: emailBodyHTML,
      inlineImages: inlineImages,
    };

    MailApp.sendEmail(message);

  } catch(e) {
    // Log and return error
    console.log(`(sendEmail) ${e.message}`);
    throw new Error(e);
  }
  // Return success message
  return {message: 'Successfully sent!', isError : false};
}



/**
 * Sends email using member information.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {{key:value<string>}} memberInformation  Information to populate email draft
 * @return {{message:string, isError:bool}}  Status of sending email.
*/
function sendEmail_(memberInformation) {
  // Gets the draft Gmail message to use as a template
  const subjectLine = DRAFT_SUBJECT_LINE;
  const emailTemplate = getGmailTemplateFromDrafts(subjectLine);

  try {
    const memberEmail = memberInformation['EMAIL'];
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, memberInformation);

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


/**
 * Get a Gmail draft message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {string} subjectLine to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
*/

function getGmailTemplateFromDrafts(subjectLine = DRAFT_SUBJECT_LINE){
  try {
    // Get the target draft, then message object
    const drafts = GmailApp.getDrafts();
    const filteredDrafts = drafts.filter(subjectFilter_(subjectLine));

    if (filteredDrafts.length > 1) {
      throw new Error (`Too many drafts with subject line '${subjectLine}. Please review.`);
    }

    const draft = filteredDrafts[0];
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    //const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();
    //DriveApp.createFile('testFile3a', htmlBody);

    const allInlineImages = draft.getMessage().getAttachments({
      includeInlineImages: true,
      includeAttachments:false
    });

    //Initiate the allInlineImages object
    var inlineImagesObj = {}
    //Regexp to search for all string positions 
    var regexp = RegExp('img data-surl=\"cid:', 'g');
    var indices = htmlBody.matchAll(regexp)

    //Iterate through all matches
    var i = 0;
    for (const match of indices){
      //Get the start position of the CID
      var thisPos = match.index + 19
      //Get the CID
      var thisId = htmlBody.substring(thisPos, thisPos + 15).replace(/"/,"").replace(/\s.*/g, "")
      //Add to object
      inlineImagesObj[thisId] = allInlineImages[i];
      i++
    }

    /* // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const imgObj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    for(const [key, value] of Object.entries(imgObj)) {
      console.log(key + " has value... " + value);
    }

    // Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    // Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => {
      console.log(match);
      inlineImagesObj[match[1]] = imgObj[match[2]];
    })

    console.log(matches); */
       

    const draftObj = {
      message: {
        subject: subjectLine, 
        text: msg.getPlainBody(), 
        html:htmlBody
      }, 
      //attachments: attachments, 
      inlineImages: inlineImagesObj 
    };

    console.log(inlineImagesObj);
    for(const [key, value] of Object.entries(inlineImagesObj)) {
      console.log(key + " has value... " + value);
    }


    return draftObj;
     
  } catch(e) {
    throw new Error("Oops - can't create template from draft. Error: " + e.message);
  }
}


/**
 * Filter draft objects with the matching subject line message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {string} subjectLine to search for draft message
 * @return {object} GmailDraft object
*/

function subjectFilter_(subjectLine){
  return function(element) {
    if (element.getMessage().getSubject() === subjectLine) {
      return element;
    }
  }
}

/**
 * Fill template string with data object and add current year.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * @see https://stackoverflow.com/a/378000/1027723
 * 
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} JSON-formatted message replaced with data
*/

function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let templateStr = JSON.stringify(template);

  // Add year for copyright message
  data['THIS_YEAR'] = String(new Date().getFullYear());

  templateStr = templateStr.replace(/{{[^{}]+}}/g, key => { 
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });

  return JSON.parse(templateStr);


  /* for (const [key, value] of Object.entries(data)) {
    const safeValue = escapeData_(value) || "";
    const regex = new RegExp(`{{${key}}}`, 'g');
    templateStr = templateStr.replace(regex, safeValue);
  } */


}

/**
 * Escape cell data to make JSON safe.
 * 
 * @author  Martin Hawksey (2022)
 * @see https://stackoverflow.com/a/9204218/1027723
 * 
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/

function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}
