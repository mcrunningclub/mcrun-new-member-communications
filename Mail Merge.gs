/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

THIS FILE HAS BEEN MODIFIED BY ANDREY GONZALEZ AS FOLLOWING:
- Reconfigured data retrieval of new member registrations
- Removed `onOpen()` and menu creation for script executions
- Deleted redundant comments
- Renamed variables using camelCase

- Modifications of `sendEmails`:
  * Removed browser prompt
  * Modified parameters to target row instead of sheet-wide
  * Changed scope of helper functions to project-wide
*/

/**
 * Sends an email with inline images as an example.
 */

function inlineImage_() {
  const googleLogoUrl = 'https://www.gstatic.com/images/branding/googlelogo/1x/googlelogo_color_74x24dp.png';
  const youtubeLogoUrl = 'https://developers.google.com/youtube/images/YouTube_logo_standard_white.png';
  const googleLogoBlob = UrlFetchApp.fetch(googleLogoUrl).getBlob().setName('googleLogoBlob');
  const youtubeLogoBlob = UrlFetchApp.fetch(youtubeLogoUrl).getBlob().setName('youtubeLogoBlob');
  MailApp.sendEmail({
    to: 'andreysebastian10.g@gmail.com',
    subject: 'Logos',
    htmlBody: 'inline Google Logo<img src=\'cid:googleLogo\'> images! <br>' +
      'inline YouTube Logo <img src=\'cid:youtubeLogo\'>',
    inlineImages: {
      googleLogo: googleLogoBlob,
      youtubeLogo: youtubeLogoBlob,
    },
  });
}


/**
 * Saves the HTML content of a Gmail draft (by subject) as a file in Google Drive.
 */

function saveDraftAsHtml() {
  const subjectLine = 'Here\'s your post-run report! 🙌';
  generateHtmlFromDraft_(subjectLine);
}


/**
 * Generates and saves the HTML version of an email draft found by subject line.
 *
 * @param {string} subjectLine  Subject line of the target draft.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function generateHtmlFromDraft_(subjectLine) {
  const datetime = Utilities.formatDate(new Date(), TIMEZONE, 'MMM-dd\'T\'hh.mm');
  const baseName = subjectLine.replace(/ /g, '-').toLowerCase();

  // Create filename for html file
  const fileName = `${baseName}-html-${datetime}`;

  // Find template in drafts and get email objects
  const emailTemplate = getGmailTemplateFromDrafts(subjectLine);
  const msgObj = fillInTemplateFromObject_(emailTemplate.message, {});

  // Save html file in drive
  DriveApp.createFile(fileName, msgObj.html);
}


/**
 * Caches a Drive file's blob to script properties for faster access.
 */

function cacheBlobToStore() {
  //cacheBlobToProperties_('1ctHsQstsoHVyCH7XcbkUNjPEka9zV9L6', 'emailHeaderBlob');
  //cacheBlobToProperties_('1Im1c4-20Sx1xLlGgWKkTxXU9OXTKct8I', 'linktreeLogoBlob');
  //cacheBlobToProperties_('1rg72NxBtCAzQsKhCRx_Fb0azzoD8ztZ-', 'stravaLogoBlob');
  cacheBlobToProperties_('1v8bSVxgM9rr5u1vjKB7qLEuaSu5xjgf2', 'runMapBlob');
}


/**
 * Encodes a Drive file's blob and stores it in script properties.
 *
 * @param {string} fileId  The Drive file ID.
 * @param {string} blobName  The key to store the blob under.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function cacheBlobToProperties_(fileId, blobName) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const encodedBlob = Utilities.base64Encode(blob.getBytes());
  console.log(encodedBlob);
  return;
  
  PropertiesService.getScriptProperties().setProperty(blobName, encodedBlob);
  console.log(`${blobName} cached in properties!`);
}


/**
 * Retrieves a blob from script properties by key.
 *
 * @param {string} blobKey  The key for the stored blob.
 * @returns {Blob} The decoded blob.
 * @throws {Error} If the blob is not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 */

function getBlobFromProperties_(blobKey) {
  const encodedBlob = PropertiesService.getScriptProperties().getProperty(blobKey);
  if (encodedBlob) {
    return Utilities.newBlob(Utilities.base64Decode(encodedBlob), 'image/png', blobKey);
  }
  throw new Error(`Blob ${blobKey} not found. Please add to script properties.`);
}


/**
 * Tests the runtime of sending an email using cached blobs vs DriveApp.
 */

function testRuntime() {
  const recipient = 'andrey.gonzalez@mail.mcgill.ca';
  const startTime = new Date().getTime();

  // Runtime if using DriveApp call : 1200ms
  // If caching images in script properties once: 550 ms
  sendSamosaEmailFromHTML(recipient, 'Test 5 samosa sale');

  //sendSamosaEmail();    // around 3000ms
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}



/**
 * Sends an email using member information and a Gmail draft as a template.
 *
 * @author Martin Hawksey (2022)
 * @update [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * @param {Object.<string, string>} memberInformation  Information to populate the email draft.
 * @returns {{message: string, isError: boolean}}  Status of sending email.
 * @throws {Error} If sending fails.
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
 * Retrieves a Gmail draft message by matching the subject line.
 *
 * @author Martin Hawksey (2022)
 * @update Andrey Gonzalez (<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * @param {string} [subjectLine=DRAFT_SUBJECT_LINE]  Subject line to search for draft message.
 * @returns {Object} Object containing the subject, plain and HTML message body, and inline images.
 * @throws {Error} If multiple or no drafts are found, or on other errors.
 */
function getGmailTemplateFromDrafts(subjectLine = DRAFT_SUBJECT_LINE) {
  // Verify if McRUN draft to search
  if (Session.getActiveUser().getEmail() != MCRUN_EMAIL) {
    return Logger.log('Change Gmail Account');
  }

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
 * Returns a filter function to match Gmail drafts by subject line.
 *
 * @param {string} subjectLine  The subject line to match.
 * @returns {function(Object): boolean}  Filter function for Gmail drafts.
 */
function subjectFilter_(subjectLine) {
  return function(element) {
    if (element.getMessage().getSubject() === subjectLine) {
      return element;
    }
  }
}

/**
 * Fills a template string with data from an object and adds the current year.
 *
 * @author Martin Hawksey (2022)
 * @update Andrey Gonzalez (<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template  String containing {{}} markers to be replaced.
 * @param {Object} data  Object used to replace {{}} markers.
 * @returns {Object}  JSON-formatted message with replaced data.
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
 * Escapes cell data to make it JSON safe.
 *
 * @author Martin Hawksey (2022)
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str  String to escape JSON special characters from.
 * @returns {string} Escaped string.
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
