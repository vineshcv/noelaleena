/**
 * Paste this into Extensions > Apps Script on the RSVP Google Sheet.
 * Deploy > New deployment > Web app.
 * Execute as: Me
 * Who has access: Anyone
 * Copy the web app URL (ends with /exec) into RSVP_SHEET_URL in index.html.
 */
function doGet() {
  return ContentService
    .createTextOutput('Noel & Aleena RSVP sheet is connected.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getRsvpSheet_();
    const fields = readRsvpFields_(e);
    const allergyDetails = String(fields.allergyDetails || '').trim();
    const hasAllergies = allergyDetails ? 'yes' : (fields.hasAllergies || '');
    sheet.appendRow([
      new Date(),
      fields.joiningFor || '',
      fields.guestName || '',
      fields.mobileNumber || '',
      fields.guestCount || '',
      fields.foodPreference || '',
      hasAllergies,
      allergyDetails
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function readRsvpFields_(e) {
  const fields = {};
  const params = (e && e.parameter) ? e.parameter : {};
  Object.keys(params).forEach(function (key) {
    fields[key] = params[key];
  });
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function (key) {
        if (body[key] !== undefined && body[key] !== null) {
          fields[key] = String(body[key]);
        }
      });
    } catch (err) {
      // URL-encoded posts are already in e.parameter.
    }
  }
  return fields;
}

function getRsvpSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('RSVP');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('RSVP');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Submitted At',
      'Joining On Behalf Of',
      'Guest Name',
      'Mobile Number',
      'Number of Guests',
      'Food Preference',
      'Food Allergies',
      'Allergy Details'
    ]);
  } else {
    ensureMobileNumberColumn_(sheet);
  }
  return sheet;
}

/** Inserts Mobile Number after Guest Name if the sheet was created before that field existed. */
function ensureMobileNumberColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Mobile Number') !== -1) {
    return;
  }
  const guestNameCol = headers.indexOf('Guest Name') + 1;
  if (guestNameCol > 0) {
    sheet.insertColumnAfter(guestNameCol);
    sheet.getRange(1, guestNameCol + 1).setValue('Mobile Number');
  } else {
    sheet.getRange(1, lastCol + 1).setValue('Mobile Number');
  }
}
