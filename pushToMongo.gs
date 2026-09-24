// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// Sheet: Shipment Profile Report (Cargowise export)
// Sheet ID: 1x1WEhIxCPJamtDnKNyGvF6R3H88cwuDDK2ud1OemV2c
// Row 1 = headers, data starts row 2
//
// HOW TO USE:
// 1. In Apps Script, create a new file called "pushOps"
// 2. Paste this entire script
// 3. Run pushAll() once manually to test
// 4. Set trigger: pushAll → Mon & Thu 9:00 AM

const OPS_BATCH_URL    = 'https://sar-indonesia-sea-ops-dashboard.vercel.app/api/mongo-batch';
const OPS_BATCH_SECRET = 'Harsh@2644';
const OPS_SHEET_ID     = '1x1WEhIxCPJamtDnKNyGvF6R3H88cwuDDK2ud1OemV2c';

// ── EXACT column indices (0-based) from row 1 header ──────────────────────
const C = {
  SHIPMENT_ID:     2,   // Shipment ID
  MODE:            9,   // Mode (FCL/LCL)
  ORIGIN:         14,   // Origin
  ORIGIN_CTRY:    15,   // Origin Ctry
  DEST:           16,   // Destination
  DEST_CTRY:      17,   // Destination Country
  ORIGIN_ETD:     28,   // Origin ETD
  JOB_BRANCH:     56,   // Job Branch
  JOB_DEPT:       57,   // Job Dept (FES=Export, FIS=Import)
  LOCAL_CLIENT:   59,   // Local Client Name
  SALES_REP:      60,   // Job Sales Rep
  OPERATOR:       61,   // Job Operator
  JOB_STATUS:     62,   // Job Status (WRK/CLS/CMP)
  JOB_OPENED:     63,   // Job Opened
  REV_REC:        64,   // Recognized Revenue
  COST_REC:       67,   // Recognized Cost
  JOB_PROFIT:     70,   // Job Profit
  ETD_FIRST_LOAD: 74,   // ETD First Load ← SOB date
  ETA_LAST_DISC:  75,   // ETA Last Discharge
  MBL_NUMBER:     76,   // Master (MBL number)
  VESSEL:         77,   // Vessel
  ETD_LOAD:       81,   // ETD Load
  ETA_DISC:       82,   // ETA Discharge
  CARRIER_NAME:   90,   // Carrier Name
  TEU:            91,   // TEU
  UNRECOG_REV:   103,   // Unrecognized Revenue
  DIRECTION:     121,   // Direction (Export/Import) ← direct field!
  CONSOL_ATD:    118,   // Consol ATD
  CONSOL_ATA:    119,   // Consol ATA
  HBL_RELEASED:  137,   // HBL Released Date ← direct field!
  INVOICE_DATE:  138,   // Invoice Date ← direct field!
  SHIPPED_OB:    140,   // Shipped On Board
  MBL_RELEASED:  142,   // MBL Released ← direct field!
  MARGIN_PCT:    143,   // Margin %
};

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val.toISOString();
  // Handle "DEP 19-Sep-22" or "IMM" style strings
  const s = String(val).trim();
  if (!s || s === 'IMM' || s === 'N/A') return null;
  // Try extracting date from strings like "DEP 19-Sep-22"
  const match = s.match(/(\d{1,2})[\-\/]([A-Za-z]+)[\-\/](\d{2,4})/);
  if (match) {
    const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const day = parseInt(match[1]);
    const mon = months[match[2].toLowerCase().slice(0,3)];
    const yr  = match[3].length === 2 ? 2000 + parseInt(match[3]) : parseInt(match[3]);
    const d = new Date(yr, mon, day);
    return isNaN(d) ? null : d.toISOString();
  }
  const d = new Date(s);
  return isNaN(d) ? null : d.toISOString();
}

function parseNum(val) {
  if (!val && val !== 0) return 0;
  const n = parseFloat(String(val).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function pushAll() {
  Logger.log('=== SAR ID Ops Push Starting ===');

  const ss = SpreadsheetApp.openById(OPS_SHEET_ID);
  const sh = ss.getSheets()[0]; // first sheet
  if (!sh) { Logger.log('ERROR: Sheet not found'); return; }

  const data = sh.getDataRange().getValues();
  const rows = data.slice(1); // skip header row 1

  const byDir = { Export: [], Import: [] };

  rows.forEach(row => {
    const shipmentId = String(row[C.SHIPMENT_ID] || '').trim();
    if (!shipmentId || shipmentId.startsWith('#')) return; // skip config/empty rows

    // Use Direction column directly — "Export" or "Import"
    const dir = String(row[C.DIRECTION] || '').trim();
    if (dir !== 'Export' && dir !== 'Import') return;

    // Also verify via Job Dept as backup
    const dept = String(row[C.JOB_DEPT] || '').trim();

    const etdDate  = parseDate(row[C.ETD_FIRST_LOAD]) || parseDate(row[C.ETD_LOAD]) || parseDate(row[C.ORIGIN_ETD]);
    const revRec   = parseNum(row[C.REV_REC]);
    const costRec  = Math.abs(parseNum(row[C.COST_REC])); // stored negative
    const profit   = parseNum(row[C.JOB_PROFIT]);
    const unrecRev = parseNum(row[C.UNRECOG_REV]);

    byDir[dir].push({
      shipmentId,
      direction: dir,
      mode:            String(row[C.MODE]          || '').trim(),
      origin:          String(row[C.ORIGIN]        || '').trim(),
      originCtry:      String(row[C.ORIGIN_CTRY]   || '').trim(),
      destination:     String(row[C.DEST]          || '').trim(),
      destCtry:        String(row[C.DEST_CTRY]     || '').trim(),
      jobBranch:       String(row[C.JOB_BRANCH]    || '').trim(),
      jobDept:         dept,
      localClientName: String(row[C.LOCAL_CLIENT]  || '').trim(),
      salesRep:        String(row[C.SALES_REP]     || '').trim(),
      operator:        String(row[C.OPERATOR]      || '').trim(),
      jobStatus:       String(row[C.JOB_STATUS]    || '').trim(),
      vessel:          String(row[C.VESSEL]        || '').trim(),
      carrierName:     String(row[C.CARRIER_NAME]  || '').trim(),
      mblNumber:       String(row[C.MBL_NUMBER]    || '').trim(),
      teu:             parseNum(row[C.TEU]),
      jobOpenedDate:   parseDate(row[C.JOB_OPENED]),
      etdDate,                                          // SOB date
      etaDate:         parseDate(row[C.ETA_LAST_DISC]) || parseDate(row[C.ETA_DISC]),
      consolAtd:       parseDate(row[C.CONSOL_ATD]),
      consolAta:       parseDate(row[C.CONSOL_ATA]),
      shippedOnBoard:  parseDate(row[C.SHIPPED_OB]),
      recognizedRevenue:   revRec,
      recognizedCost:      costRec,
      jobProfit:           profit,
      unrecognizedRevenue: unrecRev,
      totalRevenue:        revRec + unrecRev,
      marginPct:           parseNum(row[C.MARGIN_PCT]),
      // Direct milestone fields from sheet
      hblReleasedDate: parseDate(row[C.HBL_RELEASED]),
      mblReleasedDate: parseDate(row[C.MBL_RELEASED]),
      invoiceDate:     parseDate(row[C.INVOICE_DATE]),
    });
  });

  Logger.log('Export records: ' + byDir.Export.length);
  Logger.log('Import records: ' + byDir.Import.length);

  ['Export', 'Import'].forEach(dir => {
    const records = byDir[dir];
    if (!records.length) { Logger.log('No ' + dir + ' records, skipping'); return; }

    const CHUNK = 500;
    let pushed = 0;
    for (let i = 0; i < records.length; i += CHUNK) {
      const chunk = records.slice(i, i + CHUNK);
      try {
        const resp = UrlFetchApp.fetch(OPS_BATCH_URL, {
          method: 'POST',
          contentType: 'application/json',
          headers: { 'x-batch-secret': OPS_BATCH_SECRET },
          payload: JSON.stringify({ action: 'push', direction: dir, records: chunk }),
          muteHttpExceptions: true,
        });
        const code = resp.getResponseCode();
        const body = resp.getContentText().slice(0, 200);
        Logger.log(dir + ' chunk ' + (Math.floor(i/CHUNK)+1) + ': HTTP ' + code + ' — ' + body);
        if (code === 200) pushed += chunk.length;
        Utilities.sleep(300);
      } catch(e) {
        Logger.log('ERROR ' + dir + ': ' + e.message);
      }
    }
    Logger.log(dir + ': pushed ' + pushed + '/' + records.length);
  });

  Logger.log('=== SAR ID Ops Push Done ===');
}
