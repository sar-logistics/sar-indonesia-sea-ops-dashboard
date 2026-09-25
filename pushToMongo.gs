// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// File: SHIPMENT PROFILE REPORT IND CUS .XLSX
// Sheet ID: 1x1WEhIxCPJamtDnKNyGvF6R3H88cwuDDK2ud1OemV2c
// Tab "Shipment Profile Export" → Export records
// Tab "Shipment Profile Import" → Import records
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
const TAB_EXPORT       = 'Shipment Profile Export';
const TAB_IMPORT       = 'Shipment Profile Import';

// ── EXACT column indices (0-based, A=0) verified from sheet on 25-Sep-2026 ──
const C = {
  SHIPMENT_ID:     0,   // A  — Shipment ID
  WIP:             3,   // D  — WIP (SUMIF from WIP,ACCURAL tab)
  ACCURAL:         4,   // E  — Accrual (SUMIF from WIP,ACCURAL tab)
  MODE:            9,   // J  — Mode (FCL/LCL/AIR)
  ORIGIN:         14,   // O  — Origin
  ORIGIN_CTRY:    15,   // P  — Origin Ctry
  DEST:           16,   // Q  — Destination
  DEST_CTRY:      17,   // R  — Destination Country
  ORIGIN_ETD:     27,   // AB — Origin ETD
  JOB_BRANCH:     55,   // BD — Job Branch
  JOB_DEPT:       56,   // BE — Job Dept
  LOCAL_CLIENT:   58,   // BG — Local Client Name
  SALES_REP:      59,   // BH — Job Sales Rep
  OPERATOR:       60,   // BI — Job Operator
  JOB_STATUS:     61,   // BJ — Job Status (WRK/CLS/CMP)
  JOB_OPENED:     62,   // BK — Job Opened
  REV_REC:        63,   // BL — Recognized Revenue
  REV_WIP:        64,   // BM — Recognized WIP
  COST_REC:       67,   // BP — Recognized Cost
  JOB_PROFIT:     69,   // BR — Job Profit
  ETD_FIRST_LOAD: 73,   // BV — ETD First Load ← SOB date
  ETA_LAST_DISC:  74,   // BW — ETA Last Discharge
  MBL_NUMBER:     75,   // BX — Master (MBL number)
  VESSEL:         76,   // BY — Vessel
  ETD_LOAD:       79,   // CC — ETD Load
  ETA_DISC:       80,   // CD — ETA Discharge
  CARRIER_NAME:   88,   // CL — Carrier Name
  TEU:            89,   // CM — TEU
  CONSOL_ATD:    113,   // DF — Consol ATD
  CONSOL_ATA:    114,   // DG — Consol ATA
  DIRECTION:     112,   // DI — Direction (Export/Import)
  HBL_RELEASED:  129,   // DZ — HBL Released Date
  INVOICE_DATE:  130,   // EA — Invoice Date
  SHIPPED_OB:    132,   // EC — Shipped On Board
  MBL_RELEASED:  134,   // EE — MBL Released
  MARGIN_PCT:    135,   // EF — Margin %
};

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val.toISOString();
  const s = String(val).trim();
  if (!s || s === 'IMM' || s === 'N/A' || s === '#DIV/0!') return null;
  const match = s.match(/(\d{1,2})[\-\/]([A-Za-z]+)[\-\/](\d{2,4})/);
  if (match) {
    const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const day = parseInt(match[1]);
    const mon = months[match[2].toLowerCase().slice(0,3)];
    const yr  = match[3].length === 2 ? 2000 + parseInt(match[3]) : parseInt(match[3]);
    const d   = new Date(yr, mon, day);
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

function processTab(sheet, direction) {
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // skip header
  const records = [];

  rows.forEach(row => {
    const shipmentId = String(row[C.SHIPMENT_ID] || '').trim();
    if (!shipmentId || shipmentId.startsWith('#') || shipmentId === 'Shipment ID') return;

    const etdDate = parseDate(row[C.ETD_FIRST_LOAD]) || parseDate(row[C.ETD_LOAD]) || parseDate(row[C.ORIGIN_ETD]);
    const revRec  = parseNum(row[C.REV_REC]);
    const costRec = Math.abs(parseNum(row[C.COST_REC]));
    const profit  = parseNum(row[C.JOB_PROFIT]);

    records.push({
      shipmentId,
      direction,
      mode:                String(row[C.MODE]          || '').trim(),
      origin:              String(row[C.ORIGIN]        || '').trim(),
      originCtry:          String(row[C.ORIGIN_CTRY]   || '').trim(),
      destination:         String(row[C.DEST]          || '').trim(),
      destCtry:            String(row[C.DEST_CTRY]     || '').trim(),
      jobBranch:           String(row[C.JOB_BRANCH]    || '').trim(),
      jobDept:             String(row[C.JOB_DEPT]      || '').trim(),
      localClientName:     String(row[C.LOCAL_CLIENT]  || '').trim(),
      salesRep:            String(row[C.SALES_REP]     || '').trim(),
      operator:            String(row[C.OPERATOR]      || '').trim(),
      jobStatus:           String(row[C.JOB_STATUS]    || '').trim(),
      vessel:              String(row[C.VESSEL]        || '').trim(),
      carrierName:         String(row[C.CARRIER_NAME]  || '').trim(),
      mblNumber:           String(row[C.MBL_NUMBER]    || '').trim(),
      teu:                 parseNum(row[C.TEU]),
      jobOpenedDate:       parseDate(row[C.JOB_OPENED]),
      etdDate,
      etaDate:             parseDate(row[C.ETA_LAST_DISC]) || parseDate(row[C.ETA_DISC]),
      consolAtd:           parseDate(row[C.CONSOL_ATD]),
      consolAta:           parseDate(row[C.CONSOL_ATA]),
      shippedOnBoard:      parseDate(row[C.SHIPPED_OB]),
      recognizedRevenue:   revRec,
      recognizedCost:      costRec,
      jobProfit:           profit,
      wip:                 parseNum(row[C.WIP]),
      accural:             parseNum(row[C.ACCURAL]),
      marginPct:           parseNum(row[C.MARGIN_PCT]),
      hblReleasedDate:     parseDate(row[C.HBL_RELEASED]),
      mblReleasedDate:     parseDate(row[C.MBL_RELEASED]),
      invoiceDate:         parseDate(row[C.INVOICE_DATE]),
    });
  });

  return records;
}

function pushRecords(records, direction) {
  if (!records.length) { Logger.log('No ' + direction + ' records, skipping'); return; }
  const CHUNK = 500;
  let pushed = 0;
  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(OPS_BATCH_URL, {
        method: 'POST',
        contentType: 'application/json',
        headers: { 'x-batch-secret': OPS_BATCH_SECRET },
        payload: JSON.stringify({ action: 'push', direction, records: chunk }),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      Logger.log(direction + ' chunk ' + (Math.floor(i/CHUNK)+1) + ': HTTP ' + code + ' — ' + resp.getContentText().slice(0,100));
      if (code === 200) pushed += chunk.length;
      Utilities.sleep(300);
    } catch(e) {
      Logger.log('ERROR ' + direction + ': ' + e.message);
    }
  }
  Logger.log(direction + ': pushed ' + pushed + '/' + records.length);
}

function pushAll() {
  Logger.log('=== SAR ID Ops Push Starting ===');
  const ss = SpreadsheetApp.openById(OPS_SHEET_ID);

  // Export tab
  const expSheet = ss.getSheetByName(TAB_EXPORT);
  if (!expSheet) { Logger.log('ERROR: Tab "' + TAB_EXPORT + '" not found'); }
  else {
    const expRecords = processTab(expSheet, 'Export');
    Logger.log('Export records: ' + expRecords.length);
    pushRecords(expRecords, 'Export');
  }

  // Import tab
  const impSheet = ss.getSheetByName(TAB_IMPORT);
  if (!impSheet) { Logger.log('ERROR: Tab "' + TAB_IMPORT + '" not found'); }
  else {
    const impRecords = processTab(impSheet, 'Import');
    Logger.log('Import records: ' + impRecords.length);
    pushRecords(impRecords, 'Import');
  }

  Logger.log('=== SAR ID Ops Push Done ===');
}
