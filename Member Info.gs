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
const DRAFT_SUBJECT_LINE = 'Welcome!';
const DRAFT_ID = 'r-6320835261613870889';

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

const logThenReturn = (x) => { console.log(x); return x };

function printIt() {
  const x = getDraftById();
  console.log(x);
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


function sendWelcomeEmail() {
  const sheet = LITERAL_SHEET;
  const colSize = sheet.getLastColumn();
  const newRow = sheet.getLastRow();

  const headerKeys = sheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = sheet.getSheetValues(newRow, 1, 1, colSize)[0];

  const keySize = headerKeys.length;
  
  if (keySize != newMemberValues.length) {
    throw new Error(`Expected ${keySize} for newMemberValues. Received ${newMemberValues.length}`);
  }

  const removeRegex = /^[,\s\n\r\t-]+|[,\s\n\r\t-]+$/g;

  // Package member information using key-values
  const memberInformation = headerKeys.reduce(
    (obj, key, i) => (obj[key]= newMemberValues[i],obj), {}
  );

  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[HH:mm:ss]');
  const returnMsg =  sendEmail(memberInformation);

  // Update the status of email for new member
  sheet.getRange(newRow, COL_MAP.EMAIL_STATUS).setValue(`${currentTime}: ${returnMsg}`);
  
  if (typeof(returnMsg) === Error) {
    throw new Error (returnMsg);
  }

}


/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
*/
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
    const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();

    // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {
      message: {
        subject: subject_line,
        text: msg.getPlainBody(),
        html: htmlBody
      },
      attachments: attachments,
      inlineImages: inlineImagesObj
    };
  } catch (e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
  */
  function subjectFilter_(subject_line) {
    return function (element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
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
};

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}
