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

    // The source file has multiple tabs (Cover Note, Shipment Profile, Filter,
    // Sort, Optional Templates, etc.) and the real data table doesn't always
    // start on row 1 - there's a report title/filter-summary block stacked
    // above it first. Rather than hardcode a tab name or row number (which
    // shifted between Export/Import and could shift again), search every tab
    // for the row containing "Shipment ID" - the one header we know for
    // certain exists, since it's already the sheet's own first column - and
    // use that as the anchor for both which tab and which row is the real
    // header row.
    const found = findShipmentDataTable(tempSS);
    if (!found) {
      Logger.log(src.label + ': ERROR - could not find a "Shipment ID" header row in any tab of this attachment. Tabs present: ' + tempSS.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    Logger.log(src.label + ': found real data table on tab "' + found.sheetName + '", header row ' + found.headerRowNum);

    const data = found.data;
    const newHeaders = data[found.headerRowNum - 1].map(h => String(h).trim());
    const newRows = data.slice(found.headerRowNum).filter(row => row.some(cell => String(cell).trim() !== ''));
    Logger.log(src.label + ': attachment has ' + newHeaders.length + ' columns, ' + newRows.length + ' data rows');

    const ss = SpreadsheetApp.openById(SYNC_SHEET_ID);
    const targetSheet = ss.getSheetByName(src.targetTab);
    if (!targetSheet) { Logger.log(src.label + ': ERROR - target tab "' + src.targetTab + '" not found'); return; }

    const lastCol = targetSheet.getLastColumn();
    const existingHeaders = lastCol > 0
      ? targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()).filter(Boolean)
      : [];

    // Columns the sheet adds on its own (addLobCol.gs etc.) that never arrive
    // in the raw email - these are EXPECTED to be "missing from email" every
    // single run, so call them out separately instead of flagging them as a
    // generic warning each time
    const knownDerivedColumns = ['Lob', 'Derived Margin'];
    const missingFromEmail = existingHeaders.filter(h => newHeaders.indexOf(h) === -1 && knownDerivedColumns.indexOf(h) === -1);
    const derivedFound     = existingHeaders.filter(h => knownDerivedColumns.indexOf(h) !== -1);
    const extraInEmail     = newHeaders.filter(h => existingHeaders.indexOf(h) === -1);
    if (derivedFound.length)     Logger.log(src.label + ': (expected) derived columns not touched by sync: ' + derivedFound.join(', '));
    if (missingFromEmail.length) Logger.log(src.label + ': ⚠ columns in the SHEET but missing from this email: ' + missingFromEmail.join(', '));
    if (extraInEmail.length)     Logger.log(src.label + ': ⚠ columns in this email but not yet in the sheet: ' + extraInEmail.join(', '));
    if (!missingFromEmail.length && !extraInEmail.length) Logger.log(src.label + ': ✓ columns match exactly (aside from expected derived columns)');

    // Full replace, scoped to exactly the email's own column width - this
    // deliberately leaves any sheet-only derived columns (Lob, Derived
    // Margin, or anything added later) completely untouched, since clearing
    // the full row width would blank them out until addLobCol() re-runs
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      targetSheet.getRange(2, 1, lastRow - 1, newHeaders.length).clearContent();
    }
    if (newRows.length) {
      targetSheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
    }
    Logger.log(src.label + ': wrote ' + newRows.length + ' rows into "' + src.targetTab + '" (columns 1-' + newHeaders.length + ' only - derived columns left alone)');

  } finally {
    // Clean up the temporary converted file regardless of outcome
    try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) {}
  }
}

// Scans every tab of the converted spreadsheet for the row containing
// "Shipment ID" (exact match after trim) in any cell - that row is the real
// header row, and whichever tab it's on is the real data table. Returns
// { sheetName, data, headerRowNum } (headerRowNum is 1-indexed, matching
// actual sheet row numbers) or null if no tab has a matching row.
function findShipmentDataTable(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  const SCAN_DEPTH = 50; // report title/filter-summary blocks are never this tall
  for (const sheet of sheets) {
    const maxRow = Math.min(sheet.getLastRow(), SCAN_DEPTH);
    if (maxRow < 1) continue;
    const sample = sheet.getRange(1, 1, maxRow, sheet.getLastColumn()).getValues();
    for (let r = 0; r < sample.length; r++) {
      const hasShipmentId = sample[r].some(cell => String(cell).trim() === 'Shipment ID');
      if (hasShipmentId) {
        // Found it - now pull the FULL data for this tab (not just the scan sample)
        const fullData = sheet.getDataRange().getValues();
        return { sheetName: sheet.getName(), data: fullData, headerRowNum: r + 1 };
      }
    }
  }
  return null;
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
