function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}


function getDraftBySubject_(subject = DRAFT_SUBJECT_LINE) {
  return GmailApp
  .getDrafts()
  .filter(
    subjectFilter_(subject)
  )[0];
}

function getDraftById_(id = DRAFT_ID) {
  return GmailApp.getDraft(id);
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


function toCamelCase(str) {
  return str
    .toLowerCase()
    .replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
}