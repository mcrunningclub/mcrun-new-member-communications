/**
 * Gets time zone of the script
 * @return {string}  Time zone
 */
function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

/**
 * Gets email address of the current user
 * @return {string}  The user's email's address, or a blank string if address can't be accessed
 */
function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
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

/**
 * Find row in Literals sheet that has the given email
 * 
 * @param {string} targetEmail  Email address to find
 * @return {integer}  Row in Literals sheet (1-indexed), or 0 if email not found
 */
function findRowByEmail_(targetEmail) {
  const sheet = GET_LITERAL_SHEET_();
  const allEmail = sheet.getRange(1, LITERALS.EMAIL, sheet.getLastRow()).getValues();
  return allEmail.findIndex(row => row[0] === targetEmail) + 1;  // 0 to 1-index
}

/**
 * Replaces snake case string with camel case string
 * 
 * @param {string} str  String in snake case, e.g. hello_world
 * @return {string}  Input converted to camel case, e.g. helloWorld
 */
function toCamelCase_(str) {
  return str
    .toLowerCase()
    .replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
}

/**
 * Stores email header, linktree logo, strava logo, and run map images in script properties
 */
function cacheBlobToStore() {
  //cacheBlobToProperties_('1ctHsQstsoHVyCH7XcbkUNjPEka9zV9L6', 'emailHeaderBlob');
  //cacheBlobToProperties_('1Im1c4-20Sx1xLlGgWKkTxXU9OXTKct8I', 'linktreeLogoBlob');
  //cacheBlobToProperties_('1rg72NxBtCAzQsKhCRx_Fb0azzoD8ztZ-', 'stravaLogoBlob');
  cacheBlobToProperties_('1v8bSVxgM9rr5u1vjKB7qLEuaSu5xjgf2', 'runMapBlob');
}

/**
 * Stores given Google Drive file in script properties under given name
 * 
 * Gets blob from file, outputs encoded string in console. Need to manually add
 * to script properties.
 * 
 * @param {string} fileId  ID of file to store
 * @param {string} blobName  Name of property to store encoded file under
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
 * Fetches and decodes image blob from script properties using given name
 * 
 * Throws error if not found.
 * 
 * @param {string} blobName  Name of property that stores the image
 * @return {Image}  Blob decoded as png
 */
function getBlobFromProperties_(blobName) {
  const encodedBlob = PropertiesService.getScriptProperties().getProperty(blobName);
  if (encodedBlob) {
    return Utilities.newBlob(Utilities.base64Decode(encodedBlob), 'image/png', blobName);
  }
  throw new Error(`Blob ${blobName} not found. Please add to script properties.`);
}
