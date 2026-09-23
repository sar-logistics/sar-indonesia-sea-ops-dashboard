// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// 1. Go to script.google.com → New project
// 2. Paste this entire file
// 3. Update BATCH_URL and BATCH_SECRET
// 4. Run pushAll() once to test
// 5. Set trigger: pushAll → Mon & Thu 9:00 AM

const BATCH_URL    = 'https://YOUR-VERCEL-URL.vercel.app/api/mongo-batch';
const BATCH_SECRET = 'YOUR_BATCH_SECRET_HERE';
const SHEET_ID     = '1H610k9P24-3iUXlrl-RhqrH-7Iy5zYoJ4IjsuKapg84';

// ── EXACT column indices from "Shipment Details" tab (0-based) ─────────────
const C = {
  SHIPMENT_ID:     0,   // Shipment ID
  TRANS:           1,   // Trans
  MODE:            7,   // Mode
  ORIGIN:         12,   // Origin
  DEST:           14,   // Destination
  ORIGIN_ETD:     25,   // Origin ETD
  JOB_BRANCH:     53,   // Job Branch
  JOB_DEPT:       54,   // Job Dept
  LOCAL_CLIENT:   56,   // Local Client Name
  SALES_REP:      57,   // Job Sales Rep
  OPERATOR:       58,   // Job Operator
  JOB_STATUS:     59,   // Job Status
  JOB_OPENED:     60,   // Job Opened
  REV_REC:        61,   // Recognized Revenue
  COST_REC:       64,   // Recognized Cost
  JOB_PROFIT:     67,   // Job Profit
  ETD_FIRST_LOAD: 71,   // ETD First Load  ← SOB date
  ETA_LAST_DISC:  72,   // ETA Last Discharge
  MBL_NUM:        73,   // Master (MBL number)
  VESSEL:         74,   // Vessel
  ETD_LOAD:       78,   // ETD Load
  ETA_DISC:       79,   // ETA Discharge
  CARRIER_NAME:   87,   // Carrier Name
  TEU:            88,   // TEU
  UNRECOG_REV:   100,   // Unrecognized Revenue
};

// Direction is determined from Job Dept:
// FES = Sea Export, FIS = Sea Import
function getDirection(dept) {
  const d = String(dept || '').toUpperCase();
  if (d === 'FES') return 'Export';
  if (d === 'FIS') return 'Import';
  return null; // skip WFS (warehouse) etc
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val.toISOString();
  const d = new Date(val);
  return isNaN(d) ? null : d.toISOString();
}

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(String(val).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function pushAll() {
  Logger.log('=== SAR ID Ops Push Starting ===');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Shipment Details');
  if (!sh) { Logger.log('ERROR: Sheet "Shipment Details" not found'); return; }

  const data = sh.getDataRange().getValues();
  const rows = data.slice(1); // skip header row

  // Separate into Export and Import
  const byDir = { Export: [], Import: [] };

  rows.forEach(row => {
    const shipmentId = String(row[C.SHIPMENT_ID] || '').trim();
    if (!shipmentId) return;

    const dept = String(row[C.JOB_DEPT] || '').trim();
    const dir  = getDirection(dept);
    if (!dir) return; // skip non sea-freight

    const etdDate  = parseDate(row[C.ETD_FIRST_LOAD]) || parseDate(row[C.ETD_LOAD]) || parseDate(row[C.ORIGIN_ETD]);
    const revRec   = parseNum(row[C.REV_REC]);
    const costRec  = parseNum(row[C.COST_REC]);
    const profit   = parseNum(row[C.JOB_PROFIT]);
    const unrecRev = parseNum(row[C.UNRECOG_REV]);

    byDir[dir].push({
      shipmentId,
      direction: dir,
      trans:           String(row[C.TRANS]        || '').trim(),
      mode:            String(row[C.MODE]          || '').trim(),
      origin:          String(row[C.ORIGIN]        || '').trim(),
      destination:     String(row[C.DEST]          || '').trim(),
      jobBranch:       String(row[C.JOB_BRANCH]    || '').trim(),
      jobDept:         dept,
      localClientName: String(row[C.LOCAL_CLIENT]  || '').trim(),
      salesRep:        String(row[C.SALES_REP]     || '').trim(),
      operator:        String(row[C.OPERATOR]      || '').trim(),
      jobStatus:       String(row[C.JOB_STATUS]    || '').trim(),
      vessel:          String(row[C.VESSEL]        || '').trim(),
      carrierName:     String(row[C.CARRIER_NAME]  || '').trim(),
      teu:             parseNum(row[C.TEU]),
      jobOpenedDate:   parseDate(row[C.JOB_OPENED]),
      etdDate,                                        // SOB date
      etaDate:         parseDate(row[C.ETA_LAST_DISC]) || parseDate(row[C.ETA_DISC]),
      mblNumber:       String(row[C.MBL_NUM]       || '').trim(),
      recognizedRevenue: revRec,
      recognizedCost:    Math.abs(costRec),           // stored as negative, flip
      jobProfit:         profit,
      unrecognizedRevenue: unrecRev,
      totalRevenue:    revRec + unrecRev,
      // Proxy flags — real HBL/MBL/invoice dates not in sheet, use logic:
      // MBL released = MBL number exists
      mblReleasedDate: String(row[C.MBL_NUM] || '').trim() ? etdDate : null,
      // HBL released = revenue is recognized (revenue > 0)
      hblReleasedDate: revRec > 0 ? parseDate(row[C.JOB_OPENED]) : null,
      // Invoice raised = revenue > 0
      invoiceDate:     revRec > 0 ? parseDate(row[C.JOB_OPENED]) : null,
    });
  });

  Logger.log('Export records: ' + byDir.Export.length);
  Logger.log('Import records: ' + byDir.Import.length);

  // Push each direction
  ['Export', 'Import'].forEach(dir => {
    const records = byDir[dir];
    if (!records.length) { Logger.log('No ' + dir + ' records, skipping'); return; }

    // Chunk into 500
    const CHUNK = 500;
    let pushed = 0;
    for (let i = 0; i < records.length; i += CHUNK) {
      const chunk = records.slice(i, i + CHUNK);
      try {
        const resp = UrlFetchApp.fetch(BATCH_URL, {
          method: 'POST',
          contentType: 'application/json',
          headers: { 'x-batch-secret': BATCH_SECRET },
          payload: JSON.stringify({ action: 'push', direction: dir, records: chunk }),
          muteHttpExceptions: true,
        });
        const code = resp.getResponseCode();
        const body = resp.getContentText().slice(0, 200);
        Logger.log(dir + ' chunk ' + (Math.floor(i/CHUNK)+1) + ': HTTP ' + code + ' — ' + body);
        if (code === 200) pushed += chunk.length;
        Utilities.sleep(300); // avoid rate limit
      } catch(e) {
        Logger.log('ERROR pushing ' + dir + ': ' + e.message);
      }
    }
    Logger.log(dir + ': pushed ' + pushed + '/' + records.length + ' records');
  });

  Logger.log('=== SAR ID Ops Push Done ===');
}
