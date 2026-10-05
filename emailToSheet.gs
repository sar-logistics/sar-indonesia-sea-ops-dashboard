// ── SAR Indonesia Sea Ops — Email → Sheet Sync ─────────────────────────────
// Separate script, run on a daily TIME-DRIVEN TRIGGER (not manually, not on
// open/edit). Source emails arrive from noreply@sargroup.net around
// 2:02–2:04 AM every day, so set the trigger for e.g. 3:00–4:00 AM:
//   Triggers (clock icon) → + Add Trigger
//     Function: syncEmailToSheet
//     Event source: Time-driven
//     Type: Day timer, between 3am and 4am
//
// ONE-TIME SETUP REQUIRED before this works:
//   In the Apps Script editor, click Services (+ icon in the left sidebar) →
//   add "Drive API" (this is Google's Advanced Drive Service, needed to
//   convert the incoming .xlsx attachment into a readable Google Sheet —
//   Apps Script cannot parse raw .xlsx binary data on its own).
//
// WHAT IT DOES:
//   For each of Export/Import, finds the latest matching email from
//   noreply@sargroup.net with "IND CUS" in the subject and the right
//   "- EXP"/"- IMP" suffix, downloads its .xlsx attachment, converts it to
//   a temporary Google Sheet to read the data, logs any column-header
//   mismatches against the existing tab (so you can see exactly what
//   doesn't line up), then REPLACES that tab's data rows with the fresh
//   dump (each email is a full current snapshot, not incremental — matches
//   what you described). The temporary converted file is deleted after.

const SYNC_SHEET_ID = '1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY';
const SYNC_SOURCES = [
  { label: 'Export', subjectContains: 'IND CUS', directionSuffix: 'EXP', targetTab: 'Shipment Profile Export' },
  { label: 'Import', subjectContains: 'IND CUS', directionSuffix: 'IMP', targetTab: 'Shipment Profile Import' },
];

function syncEmailToSheet() {
  Logger.log('=== Email → Sheet Sync Starting (Indonesia) ===');
  SYNC_SOURCES.forEach(src => {
    try {
      processSyncSource(src);
    } catch (e) {
      Logger.log(src.label + ': ERROR - ' + e.message);
    }
  });
  Logger.log('=== Email → Sheet Sync Done ===');
}

function processSyncSource(src) {
  Logger.log('--- ' + src.label + ' ---');

  const query = 'from:noreply@sargroup.net subject:"' + src.subjectContains + '" has:attachment newer_than:2d';
  const threads = GmailApp.search(query, 0, 10);
  if (!threads.length) { Logger.log(src.label + ': no matching emails found for query: ' + query); return; }

  // Gmail's search can't filter on "subject ends with", so walk matches
  // newest-first and pick the first one whose subject actually ends with
  // the right direction suffix and has an .xlsx attachment
  let targetMessage = null, targetAttachment = null;
  for (const thread of threads) {
    const messages = thread.getMessages();
    for (let i = messages.length - 1; i >= 0; i--) {
      const msg = messages[i];
      const subject = msg.getSubject().trim().toUpperCase();
      if (subject.indexOf(src.subjectContains) === -1) continue;
      if (!subject.endsWith('- ' + src.directionSuffix) && !subject.endsWith(src.directionSuffix)) continue;
      const atts = msg.getAttachments();
      const xlsx = atts.filter(a => /\.xlsx$/i.test(a.getName()))[0];
      if (xlsx) { targetMessage = msg; targetAttachment = xlsx; break; }
    }
    if (targetMessage) break;
  }

  if (!targetAttachment) {
    Logger.log(src.label + ': found matching emails but none had a usable .xlsx attachment');
    return;
  }

  Logger.log(src.label + ': using email "' + targetMessage.getSubject() + '" (' + targetMessage.getDate() + '), attachment "' + targetAttachment.getName() + '"');

  // Convert the .xlsx blob to a temporary Google Sheet so its data can be read.
  // Tries the modern Drive API v3 syntax first, falls back to v2 — whichever
  // version of the Drive API service you added will work.
  const tempFileId = convertXlsxToTempSheet(targetAttachment, 'TEMP_SYNC_' + src.label + '_' + new Date().getTime());

  try {
    const tempSS = SpreadsheetApp.openById(tempFileId);

    // Confirmed directly against real data: the real table always lives on
    // the "Shipment Profile" tab, headers on row 15, starting at column C
    // (columns A/B are blank leader columns in the report template).
    // Hardcoded per your confirmation rather than searched for, with one
    // safety check so this doesn't silently write garbage if the report
    // template ever shifts again.
    const DATA_TAB_NAME  = 'Shipment Profile';
    const HEADER_ROW     = 15;
    const START_COLUMN   = 3; // C

    const dataTab = tempSS.getSheetByName(DATA_TAB_NAME);
    if (!dataTab) {
      Logger.log(src.label + ': ERROR - tab "' + DATA_TAB_NAME + '" not found. Tabs present: ' + tempSS.getSheets().map(s => s.getName()).join(', '));
      return;
    }

    const fullData = dataTab.getDataRange().getValues().map(row => row.slice(START_COLUMN - 1));
    if (fullData.length < HEADER_ROW) {
      Logger.log(src.label + ': ERROR - "' + DATA_TAB_NAME + '" only has ' + fullData.length + ' rows, expected headers on row ' + HEADER_ROW);
      return;
    }

    const headerCheck = String(fullData[HEADER_ROW - 1][0]).trim();
    if (headerCheck !== 'Shipment ID') {
      Logger.log(src.label + ': ⚠ WARNING - expected "Shipment ID" at row ' + HEADER_ROW + ', column ' + START_COLUMN + ' but found "' + headerCheck + '" instead. The report format may have shifted - check manually before trusting this run.');
    }

    const newHeaders = fullData[HEADER_ROW - 1].map(h => String(h).trim());
    const newRows = fullData.slice(HEADER_ROW).filter(row => row.some(cell => String(cell).trim() !== ''));
    Logger.log(src.label + ': attachment has ' + newHeaders.length + ' columns, ' + newRows.length + ' data rows (tab "' + DATA_TAB_NAME + '", header row ' + HEADER_ROW + ', from column ' + START_COLUMN + ')');

    const ss = SpreadsheetApp.openById(SYNC_SHEET_ID);
    const targetSheet = ss.getSheetByName(src.targetTab);
    if (!targetSheet) { Logger.log(src.label + ': ERROR - target tab "' + src.targetTab + '" not found'); return; }

    const lastCol = targetSheet.getLastColumn();
    const existingHeaders = lastCol > 0
      ? targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()).filter(Boolean)
      : [];

    // Columns the sheet adds on its own (addLobCol.gs etc.) that never arrive
    // in the raw email - purely informational now, logged so you can see at
    // a glance whether anything unexpected changed
    const knownDerivedColumns = ['Lob', 'Derived Margin'];
    const derivedFound = existingHeaders.filter(h => knownDerivedColumns.indexOf(h) !== -1);
    if (derivedFound.length) Logger.log(src.label + ': (expected) derived columns beyond column ' + newHeaders.length + ', not touched by sync: ' + derivedFound.join(', '));

    // Full replace, HEADER ROW INCLUDED: clears row 1 through the last data
    // row, scoped to exactly the email's own column width, then writes the
    // email's own header row fresh followed by its data rows. This means the
    // sheet's header text always matches the source exactly - no separate
    // append/compare step needed - while anything beyond the email's own
    // column count (Lob, Derived Margin, or anything else added later) is
    // left completely untouched, since the clear/write never reaches past
    // column newHeaders.length.
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 0) {
      targetSheet.getRange(1, 1, lastRow, newHeaders.length).clearContent();
    }
    targetSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    if (newRows.length) {
      targetSheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
    }
    Logger.log(src.label + ': wrote header row + ' + newRows.length + ' data rows into "' + src.targetTab + '" (columns 1-' + newHeaders.length + ' only - derived columns left alone)');

  } finally {
    // Clean up the temporary converted file regardless of outcome
    try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) {}
  }
}

// Converts an .xlsx Blob to a temporary Google Sheet, returning its file ID.
// Requires the Drive API Advanced Service to be enabled (Services → Drive API).
function convertXlsxToTempSheet(attachment, tempName) {
  const blob = attachment.copyBlob();
  const resource = { name: tempName, title: tempName, mimeType: MimeType.GOOGLE_SHEETS };
  try {
    // Drive API v3 syntax
    const file = Drive.Files.create(resource, blob, { convert: true });
    return file.id;
  } catch (e) {
    // Drive API v2 syntax (older service version)
    const file = Drive.Files.insert(resource, blob, { convert: true });
    return file.id;
  }
}
