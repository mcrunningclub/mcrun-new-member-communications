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
*/


// Example function provided by Google
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
 * User function to execute `generateHtmlFromDraft_`.
 * 
 * Must updated subject line as needed.
 * 
 */

function saveDraftAsHtml() {
  const subjectLine = 'Welcome to Our Running Club';
  generateHtmlFromDraft_(subjectLine);
}


/**
 * Generate html version of email found in draft using its subject line.
 * 
 * Saves html under {draft subject}-{datetime}-html in Google Drive.
 * 
 * @param {string} subjectLine  Subject line of target draft.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 */

function generateHtmlFromDraft_(subjectLine) {
  const datetime = Utilities.formatDate(new Date(), TIMEZONE, 'MMM-dd\'T\'hh.mm');
  const baseName = subjectLine.replace(/ /g, '-').toLowerCase();

  // Create filename for html file
  const fileName = `${baseName}-${datetime}-html`;

  // Find template in drafts and get email objects
  const emailTemplate = getTemplateFromDraft_(subjectLine);
  const msgObj = fillInTemplate(emailTemplate.message, {});

  // Save html file in drive
  DriveApp.createFile(fileName, msgObj.html);
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

function getTemplateFromDraft_(subjectLine){
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
 * Filter draft objects with the matching subject line message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {string} subject  Line to search for draft message
 * @return {object} Gmail  Draft object
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

function fillInTemplate(template, data) {
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