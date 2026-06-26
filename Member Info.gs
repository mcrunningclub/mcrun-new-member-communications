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

function createNewMemberCommunications(memberObj) {
  const thisSheet = GET_LITERAL_SHEET_();
  console.log('Starting execution now...');

  try {
    // Append member info
    const newRow = appendNewValues_(memberObj, thisSheet);
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
    logMessage_(returnMessage, thisSheet, newRow);
  }
  catch(e) {
    logMessage_(e.message, thisSheet, newRow);
    throw e;
  }
}


function appendNewValues_(memberObj, thisSheet = GET_LITERAL_SHEET_()) {
  const importMap = IMPORT_MAP;
  const entries = Object.entries(memberObj)
  const valuesToAppend = Array(entries.length);

  for (let [key, value] of entries) {
    if (key in importMap) {
      let indexInSheet = importMap[key] - 1;   // Set 1-index to 0-index
      valuesToAppend[indexInSheet] = value;
    }
  }

  // Append imported values and return new row index
  const newRow = thisSheet.getLastRow() + 1;
  const colSize = entries.length;
  thisSheet.getRange(newRow, 1, 1, colSize).setValues([valuesToAppend]);
  
  return newRow;
}

function logPaymentStatus_(status) {
  const sheet = GET_PAYMENT_LOG_SHEET_();
  const updatedRow = [];

  // Map values from the status object to the correct indexes using PAYMENT_LOG_MAP
  Object.entries(PAYMENT_LOG_MAP).forEach(([key, index]) => {
    updatedRow[index - 1] = status[key];    // Turn 1-index to 0-index
  });

  sheet.appendRow(updatedRow);
}


function findRowByEmail_(targetEmail) {
  const sheet = GET_LITERAL_SHEET_();
  const allEmail = sheet.getRange(1, LITERALS.EMAIL, sheet.getLastRow()).getValues();
  return allEmail.findIndex(row => row[0] === targetEmail) + 1;  // 0 to 1-index
}


function logMessage_(message, thisSheet =  GET_LITERAL_SHEET_(), thisRow = thisSheet.getLastRow()) {
  // Update the status of email for new member
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss]');
  const statusRange = thisSheet.getRange(thisRow, LITERALS.EMAIL_LOG);

  // Append status to previous value
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  statusRange.setValue(`${previousValue}${currentTime}: ${message}`);
}

