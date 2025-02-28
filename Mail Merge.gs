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


// Example function provided by Google
// See 
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


function cacheBlobToStore() {
  //cacheBlobToProperties_('1ctHsQstsoHVyCH7XcbkUNjPEka9zV9L6', 'emailHeaderBlob');
  cacheBlobToProperties_('1Im1c4-20Sx1xLlGgWKkTxXU9OXTKct8I', 'linktreeLogoBlob');
  //cacheBlobToProperties_('1rg72NxBtCAzQsKhCRx_Fb0azzoD8ztZ-', 'stravaLogoBlob');
}


function cacheBlobToProperties_(fileId, blobName) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const encodedBlob = Utilities.base64Encode(blob.getBytes());
  PropertiesService.getScriptProperties().setProperty(blobName, encodedBlob);
  console.log(`${blobName} cached in properties!`);
}


function getBlobFromProperties_(blobKey) {
  const encodedBlob = PropertiesService.getScriptProperties().getProperty(blobKey);
  if (encodedBlob) {
    return Utilities.newBlob(Utilities.base64Decode(encodedBlob), 'image/png', blobKey);
  }
  throw new Error(`Blob ${blobKey} not found. Please add to script properties.`);
}


function generateHtmlFromDraft() {
  const subjectLine = 'Welcome to Our Running Club';
  const fileName = 'welcome-email-html-feb-28';

  const emailTemplate = getGmailTemplateFromDrafts(subjectLine);
  const msgObj = fillInTemplateFromObject_(emailTemplate.message, {});
  DriveApp.createFile(fileName, msgObj.html);
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

