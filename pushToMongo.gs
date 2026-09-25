// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// Sheet ID: 1x1WEhIxCPJamtDnKNyGvF6R3H88cwuDDK2ud1OemV2c
// Verified column mapping: 25-Sep-2026
// Run pushAll() manually or set trigger Mon & Thu 9:00 AM

const OPS_BATCH_URL    = 'https://sar-indonesia-sea-ops-dashboard.vercel.app/api/mongo-batch';
const OPS_BATCH_SECRET = 'Harsh@2644';
const OPS_SHEET_ID     = '1x1WEhIxCPJamtDnKNyGvF6R3H88cwuDDK2ud1OemV2c';
const TAB_EXPORT       = 'Shipment Profile Export';
const TAB_IMPORT       = 'Shipment Profile Import';

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val.toISOString();
  const s = String(val).trim();
  if (!s || s === 'IMM' || s === 'N/A' || s === '#DIV/0!') return null;
  const match = s.match(/(\d{1,2})[\-\/]([A-Za-z]+)[\-\/](\d{2,4})/);
  if (match) {
    const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const d = new Date(
      match[3].length === 2 ? 2000 + parseInt(match[3]) : parseInt(match[3]),
      months[match[2].toLowerCase().slice(0,3)],
      parseInt(match[1])
    );
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

function str(val) { return String(val || '').trim(); }

function processTab(sheet, direction) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => str(h));
  const rows = data.slice(1);
  const records = [];

  rows.forEach(row => {
    const shipmentId = str(row[0]); // A - Shipment ID
    if (!shipmentId || shipmentId.startsWith('#') || shipmentId === 'Shipment ID') return;

    // Build record with ALL columns as key:value using header names
    const rec = { shipmentId, direction };

    // Map every column by header name → camelCase key
    headers.forEach((h, i) => {
      if (!h || i === 0) return; // skip empty headers and Shipment ID (already set)
      const val = row[i];
      if (val === '' || val === null || val === undefined) return;

      // Determine if date, number or string
      const key = headerToKey(h, i);
      if (!key) return;

      if (isDateCol(i)) {
        const d = parseDate(val);
        if (d) rec[key] = d;
      } else if (isNumCol(i)) {
        rec[key] = parseNum(val);
      } else {
        const s = str(val);
        if (s) rec[key] = s;
      }
    });

    records.push(rec);
  });

  return records;
}

// Map column index to camelCase key name
function headerToKey(header, idx) {
  // Key columns we care about — named explicitly
  const explicit = {
    1:  'trans',
    2:  'bookingReceived',
    3:  'wip',
    4:  'accural',
    5:  'bookingIssued',
    6:  'customsInfo',
    7:  'carrierConfirmed',
    8:  'contractNo',
    9:  'mode',
    10: 'stuffingLocation',
    11: 'customerService',
    12: 'documentationContact',
    13: 'documentationContact2',
    14: 'origin',
    15: 'originCtry',
    16: 'destination',
    17: 'destCtry',
    18: 'consignorCode',
    19: 'consignorName',
    20: 'consigneeCode',
    21: 'consigneeName',
    22: 'houseRef',
    23: 'incoterm',
    24: 'additionalIncoterm',
    25: 'ppdCcx',
    26: 'goodsDescription',
    27: 'originEtd',
    28: 'etdMonth',
    29: 'destinationEta',
    30: 'etaMonth',
    31: 'weight',
    33: 'volume',
    35: 'loadingMeters',
    36: 'chargeable',
    38: 'innerPkgs',
    40: 'outerPkgs',
    42: 'added',
    43: 'controllingOffice1',
    44: 'controllingOffice2',
    45: 'controllingOffice3',
    46: 'controllingOffice4',
    47: 'transportJob',
    48: 'brokerageJob',
    49: 'isMasterLeader',
    50: 'masterLeaderRef',
    51: 'importBrokerage1',
    52: 'importBrokerage2',
    53: 'exportBrokerage1',
    54: 'exportBrokerage2',
    55: 'jobBranch',
    56: 'jobDept',
    57: 'localClientCode',
    58: 'localClientName',
    59: 'salesRep',
    60: 'operator',
    61: 'jobStatus',
    62: 'jobOpenedDate',
    63: 'recognizedRevenue',
    64: 'recognizedWip',
    65: 'totalRecognizedRevenue',
    66: 'recognizedCost',
    67: 'recognizedCost2',
    68: 'totalRecognizedCost',
    69: 'jobProfit',
    70: 'consolId',
    71: 'firstLoad',
    72: 'lastDischarge',
    73: 'etdDate',         // ETD First Load ← SOB
    74: 'etaDate',
    75: 'mblNumber',
    76: 'vessel',
    77: 'flightVoyage',
    78: 'loadPort',
    79: 'dischargePort',
    80: 'etdLoad',
    81: 'etaDischarge',
    82: 'sendingAgent1',
    83: 'sendingAgent2',
    84: 'receivingAgent1',
    85: 'receivingAgent2',
    86: 'coLoadedWith',
    87: 'coLoaderName',
    88: 'carrierCode',
    89: 'carrierName',
    90: 'teu',
    91: 'containerCount',
    92: 'other',
    93: 'cnt20F',
    94: 'cnt20R',
    95: 'cnt20H',
    96: 'cnt40F',
    97: 'cnt40R',
    98: 'cnt40H',
    99: 'cnt45F',
    100:'cntGen',
    101:'serviceLevel',
    102:'shippersRef',
    103:'consignorCity',
    104:'consignorState',
    105:'consignorPostal',
    106:'consigneeCity',
    107:'consigneeState',
    108:'consigneePostal',
    109:'consolAtd',
    110:'consolAta',
    111:'jobRevenueCode',
    112:'direction',
    113:'localClientCity',
    114:'localClientCountry',
    115:'overseasAgent1',
    116:'overseasAgent2',
    117:'jobOverseasAgent1',
    118:'jobOverseasAgent2',
    119:'carrBookingRef',
    120:'containerNo',
    121:'consoleType',
    122:'sector',
    123:'networkName',
    124:'releaseType',
    125:'registeredDate',
    126:'todayExchangeRate',
    127:'jobProfitLocal',
    128:'consolPayment',
    129:'hblReleasedDate',
    130:'invoiceDate',
    131:'preAlertDate',
    132:'shippedOnBoard',
    133:'firstCmpDate',
    134:'mblReleasedDate',
    135:'marginPct',
    136:'invoicedDate',
    137:'vendorPaymentDate',
    138:'vendorPaymentStatus',
  };
  return explicit[idx] || null;
}

// Date column indices
function isDateCol(i) {
  return [2,3,4,5,7,27,29,62,73,74,80,81,109,110,125,129,130,131,132,133,134,136,137].indexOf(i) >= 0;
}

// Numeric column indices
function isNumCol(i) {
  return [31,33,35,36,38,40,63,64,65,66,67,68,69,90,91,93,94,95,96,97,98,99,100,126,127,135].indexOf(i) >= 0;
}

function pushRecords(records, direction) {
  if (!records.length) { Logger.log('No ' + direction + ' records, skipping'); return; }
  const CHUNK = 200; // smaller chunks for larger payloads
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
      Logger.log(direction + ' chunk ' + (Math.floor(i/CHUNK)+1) + ': HTTP ' + code + ' — ' + resp.getContentText().slice(0,120));
      if (code === 200) pushed += chunk.length;
      Utilities.sleep(500);
    } catch(e) {
      Logger.log('ERROR ' + direction + ': ' + e.message);
    }
  }
  Logger.log(direction + ': pushed ' + pushed + '/' + records.length);
}

function pushAll() {
  Logger.log('=== SAR ID Ops Push Starting ===');
  const ss = SpreadsheetApp.openById(OPS_SHEET_ID);

  const expSheet = ss.getSheetByName(TAB_EXPORT);
  if (!expSheet) { Logger.log('ERROR: Tab "' + TAB_EXPORT + '" not found'); }
  else {
    const expRecords = processTab(expSheet, 'Export');
    Logger.log('Export records found: ' + expRecords.length);
    pushRecords(expRecords, 'Export');
  }

  const impSheet = ss.getSheetByName(TAB_IMPORT);
  if (!impSheet) { Logger.log('ERROR: Tab "' + TAB_IMPORT + '" not found'); }
  else {
    const impRecords = processTab(impSheet, 'Import');
    Logger.log('Import records found: ' + impRecords.length);
    pushRecords(impRecords, 'Import');
  }

  Logger.log('=== SAR ID Ops Push Done ===');
}

// ── WIPE ALL — clears both Export and Import from MongoDB ──────────────────
function wipeAll() {
  Logger.log('=== Wiping all records ===');
  ['Export', 'Import'].forEach(dir => {
    try {
      const resp = UrlFetchApp.fetch(OPS_BATCH_URL, {
        method: 'POST',
        contentType: 'application/json',
        headers: { 'x-batch-secret': OPS_BATCH_SECRET },
        payload: JSON.stringify({ action: 'wipe', direction: dir }),
        muteHttpExceptions: true,
      });
      Logger.log('Wipe ' + dir + ': HTTP ' + resp.getResponseCode() + ' — ' + resp.getContentText().slice(0,100));
    } catch(e) {
      Logger.log('Wipe ERROR ' + dir + ': ' + e.message);
    }
  });
  Logger.log('=== Wipe done ===');
}

// ── WIPE THEN PUSH — safest full refresh ──────────────────────────────────
function wipeAndPushAll() {
  wipeAll();
  Utilities.sleep(1000);
  pushAll();
}
