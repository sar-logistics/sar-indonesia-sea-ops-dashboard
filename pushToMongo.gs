// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// Paste this in Google Apps Script (script.google.com) attached to the sheet
// Set triggers: Mon & Thu at 9:00 AM (matches refresh frequency in req doc)

const BATCH_URL    = 'https://YOUR-VERCEL-URL.vercel.app/api/mongo-batch';
const BATCH_SECRET = 'YOUR_BATCH_SECRET_HERE'; // same as Vercel env BATCH_SECRET
const SHEET_ID     = '1H610k9P24-3iUXlrl-RhqrH-7Iy5zYoJ4IjsuKapg84';
const IDR_TO_USD   = 1 / 16500;

// Column indices (0-based) from "Shipment Details" tab
// Based on actual headers read from the sheet
const COL = {
  SHIPMENT_ID:        0,   // Shipment ID
  TRANS:              1,   // Trans
  MODE:               7,   // Mode (FCL/LCL/AIR)
  JOB_SALES_REP:     55,  // Job Sales Rep
  JOB_OPERATOR:      56,  // Job Operator
  JOB_STATUS:        57,  // Job Status
  JOB_OPENED:        58,  // Job Opened
  REV_RECOGNIZED:    59,  // Recognized Revenue
  COST_RECOGNIZED:   62,  // Recognized Cost
  JOB_PROFIT:        64,  // Job Profit
  ETD_FIRST_LOAD:    69,  // ETD First Load
  ETA_LAST_DISC:     70,  // ETA Last Discharge
  ETD_LOAD:          76,  // ETD Load
  ETA_DISC:          77,  // ETA Discharge
  DIRECTION:        105,  // Direction (Export/Import)
  JOB_BRANCH:        48,  // Job Branch
  JOB_DEPT:          49,  // Job Dept
  LOCAL_CLIENT_NAME: 51,  // Local Client Name
  CONSOL_ATD:       101,  // Consol ATD
  CONSOL_ATA:       102,  // Consol ATA
  REV_RECOGNITION:  103,  // Job Revenue Recognition Date
  UNRECOG_REV:      104,  // Unrecognized Revenue
  MBL_MASTER:        65,  // Master (MBL number / released)
  VESSEL:            66,  // Vessel
  ORIGIN:            12,  // Origin
  DEST:              14,  // Destination
  CARRIER_NAME:      84,  // Carrier Name
  TEU:               85,  // TEU
};

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val.toISOString();
  const d = new Date(val);
  return isNaN(d) ? null : d.toISOString();
}

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(String(val).replace(/[,\s]/g,''));
  return isNaN(n) ? 0 : n;
}

function pushDirection(direction) {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const sh   = ss.getSheetByName('Shipment Details');
  if (!sh) { Logger.log('Sheet "Shipment Details" not found'); return; }

  const data = sh.getDataRange().getValues();
  const rows = data.slice(1); // skip header

  const records = [];
  rows.forEach(row => {
    const dir = String(row[COL.DIRECTION] || '').trim();
    if (dir !== direction) return;

    const shipmentId = String(row[COL.SHIPMENT_ID] || '').trim();
    if (!shipmentId) return;

    const etdDate    = parseDate(row[COL.ETD_FIRST_LOAD] || row[COL.ETD_LOAD]);
    const etaDate    = parseDate(row[COL.ETA_LAST_DISC]  || row[COL.ETA_DISC]);
    const sobDate    = etdDate; // ETD = SOB for sea
    const jobStatus  = String(row[COL.JOB_STATUS] || '').trim();
    const revRec     = parseNum(row[COL.REV_RECOGNIZED]);
    const costRec    = parseNum(row[COL.COST_RECOGNIZED]);
    const jobProfit  = parseNum(row[COL.JOB_PROFIT]);
    const unrecogRev = parseNum(row[COL.UNRECOG_REV]);

    // MBL/HBL release: if Master field is filled and ATD exists, consider released
    const mblNum          = String(row[COL.MBL_MASTER] || '').trim();
    const consolAtd       = parseDate(row[COL.CONSOL_ATD]);
    const revRecognDate   = parseDate(row[COL.REV_RECOGNITION]);

    records.push({
      shipmentId,
      direction,
      mode:             String(row[COL.MODE] || '').trim(),
      jobStatus,
      jobBranch:        String(row[COL.JOB_BRANCH] || '').trim(),
      jobDept:          String(row[COL.JOB_DEPT] || '').trim(),
      salesRep:         String(row[COL.JOB_SALES_REP] || '').trim(),
      operator:         String(row[COL.JOB_OPERATOR] || '').trim(),
      localClientName:  String(row[COL.LOCAL_CLIENT_NAME] || '').trim(),
      origin:           String(row[COL.ORIGIN] || '').trim(),
      destination:      String(row[COL.DEST] || '').trim(),
      carrierName:      String(row[COL.CARRIER_NAME] || '').trim(),
      etdDate,
      etaDate,
      consolAtd,
      consolAta:        parseDate(row[COL.CONSOL_ATA]),
      jobOpenedDate:    parseDate(row[COL.JOB_OPENED]),
      recognizedRevenue: revRec,
      recognizedCost:    costRec,
      jobProfit,
      unrecognizedRevenue: unrecogRev,
      totalRevenue:     revRec + unrecogRev,
      // MBL released = master BL number exists AND vessel departed
      mblReleasedDate:  (mblNum && consolAtd) ? consolAtd : null,
      // HBL released = revenue recognition date (proxy for doc completion)
      hblReleasedDate:  revRecognDate,
      // Invoice date = revenue recognition date proxy
      invoiceDate:      revRecognDate,
      teu:              parseNum(row[COL.TEU]),
      vessel:           String(row[COL.VESSEL] || '').trim(),
    });
  });

  Logger.log(`${direction}: ${records.length} records to push`);
  if (!records.length) return;

  // Push in chunks of 500
  const CHUNK = 500;
  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    const payload = JSON.stringify({ action: 'push', direction, records: chunk });
    const resp = UrlFetchApp.fetch(BATCH_URL, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'x-batch-secret': BATCH_SECRET },
      payload,
      muteHttpExceptions: true,
    });
    Logger.log(`Chunk ${Math.floor(i/CHUNK)+1}: ${resp.getResponseCode()} ${resp.getContentText().slice(0,100)}`);
  }
}

// Push both Export and Import
function pushAll() {
  Logger.log('=== SAR ID Ops Push Starting ===');
  pushDirection('Export');
  pushDirection('Import');
  Logger.log('=== Done ===');
}

// Individual functions for manual triggers
function pushExport() { pushDirection('Export'); }
function pushImport() { pushDirection('Import'); }
