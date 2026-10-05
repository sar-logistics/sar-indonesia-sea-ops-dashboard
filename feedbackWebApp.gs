// ── SAR Indonesia Sea Ops — Feedback Web App ───────────────────────────────
// Separate script, deployed as a Web App (not run manually like the other
// scripts). Receives feedback submissions from the dashboard's "Send
// Feedback" form and logs them into a "Feedback" tab on this sheet,
// auto-creating the tab with headers on first use.
//
// DEPLOYMENT:
//   1. Paste this file into the same Apps Script project as the others.
//   2. Deploy → New deployment → type: Web app.
//      - Execute as: Me
//      - Who has access: Anyone
//   3. Copy the resulting Web App URL.
//   4. Send that URL back — it needs to replace the placeholder
//      FEEDBACK_WEB_APP_URL constant in index.html.

const FEEDBACK_SHEET_ID = '1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY';
const FEEDBACK_TAB_NAME = 'Feedback';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(FEEDBACK_SHEET_ID);
    let sheet = ss.getSheetByName(FEEDBACK_TAB_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(FEEDBACK_TAB_NAME);
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Branch', 'Feedback Type', 'Description']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.timestamp     || new Date().toISOString(),
      data.name          || '',
      data.email         || '',
      data.branch        || '',
      data.feedbackType  || '',
      data.description   || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Feedback endpoint is live. Submissions are sent via POST.');
}
