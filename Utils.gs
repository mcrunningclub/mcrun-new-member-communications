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
function toCamelCase(str) {
  return str
    .toLowerCase()
    .replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
}