/******************************************************************************
 * This tutorial is based on the work of Martin Hawksey twitter.com/mhawksey  *
 * But has been simplified and cleaned up to make it more beginner friendly   *
 * All credit still goes to Martin and any issues/complaints/questions to me. *
 ******************************************************************************/

// If you want to store your email server-side (hidden), uncomment the next line
// const TO_ADDRESS = "example@email.net";

// Spit out all the keys/values from the form in HTML for email
// Uses an array of keys if provided or the object to determine field order
function formatMailBody(obj, order) {
  let result = "";
  if (!order) {
    order = Object.keys(obj);
  }

  // Loop over all keys in the ordered form data
  for (let idx in order) {
    let key = order[idx];
    result += `<h4 style='text-transform: capitalize; margin-bottom: 0'>${key}</h4><div>${sanitizeInput(obj[key])}</div>`;
    // For every key, concatenate an `<h4 />`/`<div />` pairing of the key name and its value, 
    // and append it to the `result` string created at the start.
  }
  return result; // Once the looping is done, `result` will be one long string to put in the email body
}

// Sanitize content from the user - trust no one 
// Ref: https://developers.google.com/apps-script/reference/html/html-output#appendUntrusted(String)
function sanitizeInput(rawInput) {
  const placeholder = HtmlService.createHtmlOutput(" ");
  placeholder.appendUntrusted(rawInput);

  return placeholder.getContent();
}

function doPost(e) {
  try {
    Logger.log(e); // The Google Script version of console.log see: Class Logger
    if (e.parameters.itsatrap || e.parameters.submit) {
      throw new Error("It's a trap!");
    }

    record_data(e);

    // Shorter name for form data
    const mailData = e.parameters;

    // Names and order of form elements (if set)
    const orderParameter = e.parameters.formDataNameOrder;
    let dataOrder;
    if (orderParameter) {
      dataOrder = JSON.parse(orderParameter);
    }

    // Determine recipient of the email
    // If you have your email uncommented above, it uses that `TO_ADDRESS`
    // Otherwise, it defaults to the email provided by the form's data attribute
    const sendEmailTo = (typeof TO_ADDRESS !== "undefined") ? TO_ADDRESS : mailData.formGoogleSendEmail;

    // Send email if to address is set
    if (sendEmailTo) {
      MailApp.sendEmail({
        to: String(sendEmailTo),
        subject: "Contact form submitted",
        // replyTo: String(mailData.email), // This is optional and reliant on your form actually collecting a field named `email`
        htmlBody: formatMailBody(mailData, dataOrder)
      });
    }

    return ContentService    // Return JSON success results
      .createTextOutput(
        JSON.stringify({
          "result": "success",
          "data": JSON.stringify(e.parameters)
        }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) { // If error return this
    Logger.log(error);
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Record_data inserts the data received from the HTML form submission
 * e is the data received from the POST
 */
function record_data(e) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000); // Hold off up to 30 sec to avoid concurrent writing

  try {
    Logger.log(JSON.stringify(e)); // Log the POST data in case we need to debug it

    // Select the 'responses' sheet by default
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = e.parameters.formGoogleSheetName || "responses";
    const sheet = doc.getSheetByName(sheetName);

    const oldHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newHeader = oldHeader.slice();
    const fieldsFromForm = getDataColumns(e.parameters);
    const row = [new Date()]; // First element in the row should always be a timestamp

    // Loop through the header columns
    for (let i = 1; i < oldHeader.length; i++) { // Start at 1 to avoid Timestamp column
      const field = oldHeader[i];
      const output = getFieldFromData(field, e.parameters);
      row.push(output);

      // Mark as stored by removing from form fields
      const formIndex = fieldsFromForm.indexOf(field);
      if (formIndex > -1) {
        fieldsFromForm.splice(formIndex, 1);
      }
    }

    // Set any new fields in our form
    for (let i = 0; i < fieldsFromForm.length; i++) {
      const field = fieldsFromForm[i];
      const output = getFieldFromData(field, e.parameters);
      row.push(output);
      newHeader.push(field);
    }

    // More efficient to set values as [][] array than individually
    const nextRow = sheet.getLastRow() + 1; // Get next row
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

    // Update header row with any new data
    if (newHeader.length > oldHeader.length) {
      sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
    }
  } catch (error) {
    Logger.log(error);
  } finally {
    lock.releaseLock();
    return;
  }
}

function getDataColumns(data) {
  return Object.keys(data).filter(function (column) {
    return !(column === 'formDataNameOrder' || column === 'formGoogleSheetName' || column === 'formGoogleSendEmail');
  });
}

function getFieldFromData(field, data) {
  const values = data[field] || '';
  const output = values.join ? values.join(', ') : values;
  return output;
}
