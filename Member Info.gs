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
 * Workflow for when new member registers
 * 
 * Add new member's information to Literals sheet,
 * create a pass, save pass information, and send welcome email
 * 
 * @param {Object} memberObj  Object containing member informations
 */
function createNewMemberCommunications(memberObj) {
  const thisSheet = GET_LITERAL_SHEET_();
  console.log('Starting execution now...');

  try {
    // Append member info
    const newRow = createNewMemberLiteral_(memberObj);
    console.log('Successfully imported values to row ' + newRow);

    // Create member pass 
    const passUrl = createPassFile_(memberObj);   // Get download url for member pass
    console.log('Successfully created digital pass with url:\n' + passUrl);

    // Save url of digital pass to sheet and `memberObj`
    thisSheet.getRange(newRow, LITERALS.DIGITAL_PASS_URL).setValue(passUrl);
    memberObj['passUrl'] = passUrl;
    console.log('Successfully saved url to row ' + newRow);

    // Send welcome email and log result
    const returnMessage = sendWelcomeEmail_(memberObj);
    console.log(returnMessage);
    logEmailStatus_(returnMessage, newRow);
  }
  catch(e) {
    logEmailStatus_(e.message, newRow);
    throw e;
  }
}

/**
 * Appends a member object as a new row in the sheet, mapping fields to correct columns.
 * 
 * @param {Object} memberData - The member object (e.g., { email: '...', firstName : '...' })
 * @param {SpreadsheetApp.Sheet} literalsSheet - The Literals sheet object
 * @return {number} The row index of the newly appended row
 */
function createNewMemberLiteral_(memberObj) {
  const literalsSheet = GET_LITERAL_SHEET_();

  // Convert member object to its entries
  const memberInfo = Object.entries(memberObj)

  // Make array for information to add to new row
  const rowValues = Array(Object.keys(IMPORT_MAP).length);

  for (let [key, value] of memberInfo) {
    if (key in IMPORT_MAP) {
      // Get column and corresponding index in array (0-indexed) for each value
      let colInSheet = IMPORT_MAP[key];
      let indexInSheet = colInSheet - 1;
      rowValues[indexInSheet] = value;
    }
  }

  // Append imported values and return new row index
  const newRow = literalsSheet.getLastRow() + 1;
  const colSize = rowValues.length;
  literalsSheet.getRange(newRow, 1, 1, colSize).setValues([rowValues]);
  
  return newRow;
}

/**
 * Appends a new log to the Payment Logs sheet
 * 
 * @param {Object} statusObj  Object containing payment information
 *                            Should include timestamp, email, feeStatus
 */
function logPaymentStatus_(statusObj) {
  const sheet = GET_PAYMENT_LOG_SHEET_();
  const newRowValues = [];

  // Map values from the status object to the correct indexes using PAYMENT_LOG_MAP
  Object.entries(PAYMENT_LOG_MAP).forEach(([key, index]) => {
    newRowValues[index - 1] = statusObj[key];    // Turn 1-index to 0-index
  });

  sheet.appendRow(newRowValues);
}

/**
 * Updates the Email Status column in the literals sheet with a new message
 * 
 * Includes date and time of the message
 * 
 * @param {string} message  Message to log
 * @param {number} row  Row to log email status for
 */
function logEmailStatus_(message, row) {
  const literalsSheet = GET_LITERAL_SHEET_();

  // Update the status of email for new member
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss]');
  const statusRange = literalsSheet.getRange(row, LITERALS.EMAIL_LOG);

  // Append status to previous value
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  statusRange.setValue(`${previousValue}${currentTime}: ${message}`);
}

