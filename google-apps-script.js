/**
 * Topic Eyewear — Waitlist Collector
 *
 * DEPLOY (required — 401 errors mean this step was skipped or wrong):
 *   1. Extensions → Apps Script → paste this file → Save
 *   2. Run → doGet  (authorize when prompted)
 *   3. Deploy → New deployment → Web app
 *        Execute as:  Me
 *        Who has access:  Anyone   ← must be "Anyone", NOT "Anyone with Google account"
 *   4. Deploy → copy the NEW /exec URL into index.html
 *
 * If you still get 401: Deploy → Manage deployments → Edit → set Anyone → Deploy
 * then create a NEW deployment and use that new URL (old URLs keep old permissions).
 */

var SHEET_NAME = 'Sheet1';

function doPost(e) {
  try {
    var data = parsePayload(e);
    var email = (data.email || '').trim();
    var phone = (data.phone || '').trim();
    var timestamp = data.timestamp || new Date().toISOString();

    if (!email) {
      return respond({ result: 'error', message: 'No email provided' });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return respond({ result: 'error', message: 'Sheet not found: ' + SHEET_NAME });
    }

    sheet.appendRow([timestamp, email, phone, 'Waitlist Page']);
    return respond({ result: 'success' });

  } catch (err) {
    return respond({ result: 'error', message: err.toString() });
  }
}

function doGet(e) {
  // Allow testing in browser: /exec?email=test@example.com&phone=123
  if (e && e.parameter && e.parameter.email) {
    return doPost({ parameter: e.parameter, postData: null });
  }
  return respond({ result: 'ok', message: 'Topic Eyewear waitlist endpoint is live.' });
}

function parsePayload(e) {
  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // fall through to parameters
    }
  }
  return {
    email:     e.parameter.email,
    phone:     e.parameter.phone,
    timestamp: e.parameter.timestamp
  };
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
