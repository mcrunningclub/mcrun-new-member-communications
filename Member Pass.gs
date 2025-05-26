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
 * Creates a digital pass file for a member using a Google Slides template, fills in member info, generates a QR code, and returns a download link.
 *
 * @param {Object} passInfo  The member information object (must include firstName, lastName, memberId, etc).
 * @returns {string}  The download link for the generated pass PNG.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function createPassFile_(passInfo) {
  const TEMPLATE_ID = '14NG31db-g-bFX1OUHeRByTKN6S2QuMAkDuANOAtwF6o';   // Is not confidential
  const FOLDER_ID = '1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn';    // Is not confidential either

  // Get the template presentation
  const template = DriveApp.getFileById(TEMPLATE_ID);
  const passFolder = DriveApp.getFolderById(FOLDER_ID);

  // Use information to create custom file name
  const memberName = `${passInfo.firstName}-${passInfo.lastName}`;

  // Add formatted name to memberInfo
  const today = new Date();
  passInfo['name'] = `${passInfo.firstName} ${(passInfo.lastName).charAt(0)}.`;
  passInfo['generatedDate'] = Utilities.formatDate(today, TIMEZONE, 'MMM-dd-yyyy');
  passInfo['cYear'] = Utilities.formatDate(today, TIMEZONE, 'yyyy');

  // Make a copy to edit
  const fileDate = Utilities.formatDate(today, TIMEZONE, 'MMdd\'-\'HHmmss');
  const copyRef = template.makeCopy(`${memberName}-McRun-Pass-${fileDate}`, passFolder);
  const copyID = copyRef.getId();
  const copyFilePtr = SlidesApp.openById(copyID);

  // Replace placeholders with member data
  for (const [key, value] of Object.entries(passInfo)) {
    let placeHolder = `{{${key}}}`;
    copyFilePtr.replaceAllText(placeHolder, value);
  }

  // Open the presentation and get the first slide
  const slide = copyFilePtr.getSlides()[0];

  // Create QR code
  const qrCodeUrl = generateQrUrl_(passInfo.memberId);
  const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

  // Find shape with placeholder alt text and replace with qr code
  const qrPlaceholder = '{{qrCodeImage}}';
  const images = slide.getImages();

  for (let image of images) {
    if (image.getDescription() === qrPlaceholder) {
      // Replace the placeholder image with the QR code image
      image.replace(qrCodeBlob);
      Logger.log(`QR Code placeholder replaced with the generated QR Code ${passInfo.memberId}.`);
      break;
    }
  }

  // Save anc close copy template
  copyFilePtr.saveAndClose();

  // Set permissions to general to allow downloading
  copyRef.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

  // Create download link for member
  return `https://docs.google.com/presentation/d/${copyID}/export/png`;   // Download link for user
}


/**
 * Creates a new digital pass for a member at the given row, updates the sheet, and returns the pass URL.
 *
 * @param {number} [row=LITERAL_SHEET.getLastRow()]  The row number for the member.
 * @returns {string}  The URL of the new digital pass.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date
 * @update 
 */

function createNewPass(row = LITERAL_SHEET.getLastRow()) {
  const thisSheet = GET_LITERAL_SHEET_();
  const colSize = thisSheet.getLastColumn() - 1;    // ERROR_STATUS not needed

  const headerKeys = thisSheet.getSheetValues(1, 1, 1, colSize)[0];
  const newMemberValues = thisSheet.getRange(row, 1, 1, colSize).getDisplayValues()[0];
  
  // Package member information using key-values
  const memberInformation = headerKeys.reduce(
    (obj, key, i) => (obj[toCamelCase(key)]= newMemberValues[i], obj), {}
  );

  console.log(memberInformation);

  // Try to send email and record status
  const passUrl = createPassFile_(memberInformation);
  
  thisSheet.getRange(row, COL_MAP.DIGITAL_PASS_URL).setValue(passUrl);
  return passUrl;

  function toCamelCase(str) {
    return str
      .toLowerCase()
      .replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
  }
}


/**
 * Tests the runtime of generating a member pass and logs the result.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function testRuntime() {
  const memberEmail = 'emerson.darling@mail.mcgill.ca';
  const startTime = new Date().getTime();

  // Runtime if creating and sharing as png: 10351 ms
  // If only creating Slides file and sharing download link: 5906 ms
  const url = generateMemberPassByMaster(memberEmail);
  console.log(url);
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}


/**
 * Generates a QR code URL for a given member ID using QuickChart API.
 *
 * @param {string} memberID  The member's unique ID.
 * @returns {string}  The URL to the generated QR code image.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function generateQrUrl_(memberID) {
  const baseUrl = 'https://quickchart.io/qr?';
  const params = `text=${encodeURIComponent(memberID)}&margin=1&size=200`

  return baseUrl + params;
}


/**
 * Fetches an image from a URL and returns it as a PNG blob if successful.
 *
 * @param {string} url  The URL of the image to fetch.
 * @returns {Blob|undefined}  The image blob if successful, otherwise undefined.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function getImage_(url) {
  var response = UrlFetchApp.fetch(url).getResponseCode();
  if (response === 200) {
    var img = UrlFetchApp.fetch(url).getAs('image/png');
  }
  return img;
}


/**
 * Loads the bytes of a Drive file by ID and returns them as a base64-encoded string.
 *
 * @param {string} id  The Drive file ID.
 * @returns {string}  The base64-encoded bytes of the file.
 * 
 * @author  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function loadImageBytes_(id) {
  var bytes = DriveApp.getFileById(id).getBlob().getBytes();
  return Utilities.base64Encode(bytes);
}


/**
 * Tests the QR code generator by creating a QR code PNG in the pass folder.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */

function testQRGenerator_() {
  const FOLDER_ID = '1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn';
  const passFolder = DriveApp.getFolderById(FOLDER_ID);

  const memberId = '1zIQfQzTj1h5FNSTttXn';

  const qrCodeUrl = generateQrUrl_(memberId);
  const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

  passFolder.createFile(qrCodeBlob).setName(`QRCode-test.png`);
}
