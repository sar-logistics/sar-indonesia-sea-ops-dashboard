// ── SAR Indonesia Sea Ops — Apps Script Data Push ──────────────────────────
// Sheet ID: 1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY
// Run pushAll() manually or set trigger

const OPS_BATCH_URL    = 'https://sar-ind-sea-ops-dashboard.vercel.app/api/mongo-batch';
const OPS_BATCH_SECRET = 'Harsh@2644';
const OPS_SHEET_ID     = '1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY';
const TAB_EXPORT       = 'Shipment Profile Export';
const TAB_IMPORT       = 'Shipment Profile Import';

// Header name → MongoDB field key mapping
const HEADER_MAP = {
  'Shipment ID':                        'shipmentId',
  'Trans':                              'trans',
  'Booking Received':                   'bookingReceived',
  'WIP':                                'wip',
  'Accrual':                            'accural',
  'Booking Issued':                     'bookingIssued',
  'Customs Info':                       'customsInfo',
  'Carrier Confirmed':                  'carrierConfirmed',
  'Contract No':                        'contractNo',
  'Mode':                               'mode',
  'Stuffing Location':                  'stuffingLocation',
  'Customer Service':                   'customerService',
  'Documentation Contact':              'documentationContact',
  'Documentation Contact 2':            'documentationContact2',
  'Origin':                             'origin',
  'Origin Ctry':                        'originCtry',
  'Destination':                        'destination',
  'Dest Ctry':                          'destCtry',
  'Consignor Code':                     'consignorCode',
  'Consignor Name':                     'consignorName',
  'Consignee Code':                     'consigneeCode',
  'Consignee Name':                     'consigneeName',
  'House Ref':                          'houseRef',
  'Incoterm':                           'incoterm',
  'Additional Incoterm':                'additionalIncoterm',
  'PPD/CCX':                            'ppdCcx',
  'Goods Description':                  'goodsDescription',
  'Origin ETD':                         'originEtd',
  'ETD Month':                          'etdMonth',
  'Destination ETA':                    'destinationEta',
  'ETA Month':                          'etaMonth',
  'Weight':                             'weight',
  'Volume':                             'volume',
  'Loading Meters':                     'loadingMeters',
  'Chargeable':                         'chargeable',
  'Inner':                         'innerPkgs',
  'Outer':                         'outerPkgs',
  'Added':                              'added',
  'Controlling Office 1':               'controllingOffice1',
  'Controlling Office 2':               'controllingOffice2',
  'Controlling Office 3':               'controllingOffice3',
  'Controlling Office 4':               'controllingOffice4',
  'Transport Job':                      'transportJob',
  'Brokerage Job':                      'brokerageJob',
  'Is Master Leader':                   'isMasterLeader',
  'Master Leader Ref':                  'masterLeaderRef',
  'Import Brokerage 1':                 'importBrokerage1',
  'Import Brokerage 2':                 'importBrokerage2',
  'Export Brokerage 1':                 'exportBrokerage1',
  'Export Brokerage 2':                 'exportBrokerage2',
  'Job Branch':                         'jobBranch',
  'Job Dept':                           'jobDept',
  'Local Client Code':                  'localClientCode',
  'Local Client Name':                  'localClientName',
  'Job Sales Rep':                          'salesRep',
  'Job Operator':                           'operator',
  'Job Status':                         'jobStatus',
  'Job Opened':                    'jobOpenedDate',
  'Recognized Revenue':                 'recognizedRevenue',
  'Recognized WIP':                     'recognizedWip',
  'Total Recognized Revenue':           'totalRecognizedRevenue',
  'Recognized Cost':                    'recognizedCost',
  'Recognized Cost 2':                  'recognizedCost2',
  'Total Recognized Cost':              'totalRecognizedCost',
  'Job Profit':                         'jobProfit',
  'Consol ID':                          'consolId',
  'First Load':                         'firstLoad',
  'Last Discharge':                     'lastDischarge',
  'ETD First Load':                           'etdDate',
  'ETA Last Di...':                           'etaDate',
  'MBL Number':                         'mblNumber',
  'Vessel':                             'vessel',
  'Flight/Voyage':                      'flightVoyage',
  'Load Port':                          'loadPort',
  'Discharge Port':                     'dischargePort',
  'ETD Load':                           'etdLoad',
  'ETA Discharge':                      'etaDischarge',
  'Sending Agent 1':                    'sendingAgent1',
  'Sending Agent 2':                    'sendingAgent2',
  'Receiving Agent 1':                  'receivingAgent1',
  'Receiving Agent 2':                  'receivingAgent2',
  'Co-loaded With':                     'coLoadedWith',
  'Co-loader Name':                     'coLoaderName',
  'Carrier Code':                       'carrierCode',
  'Carrier Name':                       'carrierName',
  'TEU':                                'teu',
  'Container Count':                    'containerCount',
  'Other':                              'other',
  '20F':                                'cnt20F',
  '20R':                                'cnt20R',
  '20H':                                'cnt20H',
  '40F':                                'cnt40F',
  '40R':                                'cnt40R',
  '40H':                                'cnt40H',
  '45F':                                'cnt45F',
  'Gen':                                'cntGen',
  'Service Level':                      'serviceLevel',
  'Shippers Ref':                       'shippersRef',
  'Consignor City':                     'consignorCity',
  'Consignor State':                    'consignorState',
  'Consignor Postal':                   'consignorPostal',
  'Consignee City':                     'consigneeCity',
  'Consignee State':                    'consigneeState',
  'Consignee Postal':                   'consigneePostal',
  'Consol ATD':                         'consolAtd',
  'Consol ATA':                         'consolAta',
  'Job Revenue Code':                   'jobRevenueCode',
  'Direction':                          'directionCol',
  'Local Client City':                  'localClientCity',
  'Local Client Country':               'localClientCountry',
  'Overseas Agent 1':                   'overseasAgent1',
  'Overseas Agent 2':                   'overseasAgent2',
  'Job Overseas Agent 1':               'jobOverseasAgent1',
  'Job Overseas Agent 2':               'jobOverseasAgent2',
  'Carr Booking Ref':                   'carrBookingRef',
  'Container No':                       'containerNo',
  'Console Type':                       'consoleType',
  'Sector':                             'sector',
  'Network Name':                       'networkName',
  'Release Type':                       'releaseType',
  'Registered Date':                    'registeredDate',
  'Today Exchange Rate':                'todayExchangeRate',
  'Job Profit Local':                   'jobProfitLocal',
  'Consol Payment':                     'consolPayment',
  'HBL Released Date':                  'hblReleasedDate',
  'Invoice Date':                       'invoiceDate',
  'Pre-Alert Date':                     'preAlertDate',
  'Shipped On Board':                   'shippedOnBoard',
  'First CMP Date':                     'firstCmpDate',
  'MBL Released Date':                  'mblReleasedDate',
  'Margin %':                           'marginPct',
  'Invoiced Date':                      'invoicedDate',
  'Vendor Payment Date':                'vendorPaymentDate',
  'Vendor Payment Status':              'vendorPaymentStatus',
  'Lob':                                'lob',
  'Derived Margin':                     'derivedMargin',
};

// Fields that are dates (by MongoDB key name)
const DATE_KEYS = new Set(['bookingReceived','bookingIssued','carrierConfirmed','originEtd','destinationEta',
  'jobOpenedDate','etdDate','etaDate','etdLoad','etaDischarge','consolAtd','consolAta',
  'registeredDate','hblReleasedDate','invoiceDate','preAlertDate','shippedOnBoard',
  'firstCmpDate','mblReleasedDate','invoicedDate','vendorPaymentDate']);

// Fields that are numbers (by MongoDB key name)
const NUM_KEYS = new Set(['wip','accural','weight','volume','loadingMeters','chargeable','innerPkgs','outerPkgs',
  'recognizedRevenue','recognizedWip','totalRecognizedRevenue','recognizedCost','recognizedCost2',
  'totalRecognizedCost','jobProfit','teu','containerCount','cnt20F','cnt20R','cnt20H',
  'cnt40F','cnt40R','cnt40H','cnt45F','cntGen','todayExchangeRate','jobProfitLocal','marginPct','derivedMargin']);

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
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => str(h));
  const rows    = data.slice(1);
  const records = [];

  // Build index → key map from headers
  const colMap = {};
  headers.forEach((h, i) => {
    const key = HEADER_MAP[h];
    if (key) colMap[i] = key;
  });

  Logger.log(direction + ': found ' + Object.keys(colMap).length + ' mapped columns out of ' + headers.length);

  rows.forEach(row => {
    const shipmentId = str(row[0]);
    if (!shipmentId || shipmentId.startsWith('#') || shipmentId === 'Shipment ID') return;

    const rec = { shipmentId, direction };

    Object.entries(colMap).forEach(([i, key]) => {
      if (key === 'shipmentId') return; // already set
      const val = row[i];
      if (val === '' || val === null || val === undefined) return;

      if (DATE_KEYS.has(key)) {
        const d = parseDate(val);
        if (d) rec[key] = d;
      } else if (NUM_KEYS.has(key)) {
        rec[key] = parseNum(val);
      } else {
        const s = str(val);
        if (s) rec[key] = s;
      }
    });

    // Derived: lob from trans + direction (fallback if not in sheet)
    if (!rec.lob) {
      const t = (rec.trans || '').toUpperCase();
      if (t === 'SEA') rec.lob = direction === 'Export' ? 'FES' : 'FIS';
      else if (t === 'AIR') rec.lob = direction === 'Export' ? 'FEA' : 'FIA';
    }

    records.push(rec);
  });

  return records;
}

function pushRecords(records, direction) {
  if (!records.length) { Logger.log('No ' + direction + ' records, skipping'); return; }
  const CHUNK = 200;
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

function wipeAndPushAll() {
  wipeAll();
  Utilities.sleep(1000);
  pushAll();
  Utilities.sleep(1000);
  pushWip();
}

function addLobColumn() {
  const ss = SpreadsheetApp.openById(OPS_SHEET_ID);
  _addLob(ss.getSheetByName(TAB_EXPORT), 'Export');
  _addLob(ss.getSheetByName(TAB_IMPORT), 'Import');
  Logger.log('=== addLobColumn Done ===');
}

function _addLob(sheet, direction) {
  if (!sheet) { Logger.log('Tab not found: ' + direction); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Check if Lob already exists
  if (headers.indexOf('Lob') > -1) { Logger.log(direction + ': Lob column already exists'); return; }
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1).setValue('Lob');
  const data = sheet.getDataRange().getValues();
  let filled = 0;
  for (let i = 1; i < data.length; i++) {
    const trans = String(data[i][1] || '').trim().toUpperCase();
    let lob = '';
    if (trans === 'SEA') lob = direction === 'Export' ? 'FES' : 'FIS';
    else if (trans === 'AIR') lob = direction === 'Export' ? 'FEA' : 'FIA';
    if (lob) { sheet.getRange(i + 1, lastCol + 1).setValue(lob); filled++; }
  }
  Logger.log(direction + ': Lob filled ' + filled + ' rows');
}

// ── WIP/Accrual Push ──────────────────────────────────────────────────────
const WIP_BATCH_URL = 'https://sar-ind-sea-ops-dashboard.vercel.app/api/wip';

function _fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }
  return String(val).trim();
}

function pushWip() {
  Logger.log('=== WIP Push Starting ===');
  const ss  = SpreadsheetApp.openById(OPS_SHEET_ID);
  const sheet = ss.getSheetByName('WIP, ACCURAL');
  if (!sheet) { Logger.log('ERROR: WIP, ACCURAL tab not found'); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows    = data.slice(1);

  // Map headers to keys
  const h = (name) => headers.indexOf(name);
  const JOB_NO    = h('Job #');
  const LOCAL_REF = h('Job Local Ref');
  const TRANS     = h('Trans');
  const CONT      = h('Cont');
  const OPENED    = h('Opened');
  const JOB_BR    = h('Job Branch');
  const JOB_DEPT  = h('Job Dept');
  const JOB_STAT  = h('Job Stat');
  const CLIENT    = h('Local Client');
  const TYPE      = h('Type');
  const CHARGE    = h('Charge');
  const INVOICE   = h('Invoice Number');
  const RECOG     = h('Recognized');
  const REVENUE   = h('Revenue');
  const WIP       = h('WIP');
  const COST      = h('Cost');
  const ACCRUAL   = h('Accrual');
  const PROFIT    = h('Job Profit');
  const ETD       = h('ETD');
  const ETA       = h('ETA');
  const LOB       = h('LOB');
  const ETDETA    = h('ETD/ETA');

  const records = [];
  rows.forEach(row => {
    const jobNo = String(row[JOB_NO] || '').trim();
    if (!jobNo || jobNo === 'Job #') return;

    const wip     = parseFloat(row[WIP])     || 0;
    const revenue = parseFloat(row[REVENUE]) || 0;
    const cost    = parseFloat(row[COST])    || 0;
    const accrual = parseFloat(row[ACCRUAL]) || 0;
    const profit  = parseFloat(row[PROFIT])  || 0;

    records.push({
      shipmentId:    jobNo,
      localRef:      String(row[LOCAL_REF] || '').trim(),
      trans:         String(row[TRANS]     || '').trim(),
      cont:          String(row[CONT]      || '').trim(),
      jobBranch:     String(row[JOB_BR]    || '').trim(),
      jobDept:       String(row[JOB_DEPT]  || '').trim(),
      jobStatus:     String(row[JOB_STAT]  || '').trim(),
      localClient:   String(row[CLIENT]    || '').trim(),
      chargeType:    String(row[TYPE]      || '').trim(),
      chargeCode:    String(row[CHARGE]    || '').trim(),
      invoiceNumber: String(row[INVOICE]   || '').trim(),
      recognized:    _fmtDate(row[RECOG]),
      revenue,
      wip,
      cost,
      accrual,
      jobProfit:     profit,
      etd:           _fmtDate(row[ETD]),
      eta:           _fmtDate(row[ETA]),
      lob:           String(row[LOB]       || '').trim(),
      etdEta:        _fmtDate(row[ETDETA]),
      operator:      String(row[h('Job Operator')] || '').trim(),
    });
  });

  Logger.log('WIP records: ' + records.length);

  // Wipe then push in chunks
  UrlFetchApp.fetch(WIP_BATCH_URL, {
    method: 'POST', contentType: 'application/json',
    headers: { 'x-batch-secret': OPS_BATCH_SECRET },
    payload: JSON.stringify({ action: 'wipe' }),
    muteHttpExceptions: true,
  });
  Logger.log('WIP wiped');

  const CHUNK = 500;
  let pushed = 0;
  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    const resp = UrlFetchApp.fetch(WIP_BATCH_URL, {
      method: 'POST', contentType: 'application/json',
      headers: { 'x-batch-secret': OPS_BATCH_SECRET },
      payload: JSON.stringify({ action: 'push', records: chunk }),
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    Logger.log('WIP chunk ' + (Math.floor(i/CHUNK)+1) + ': HTTP ' + code + ' — ' + resp.getContentText().slice(0,80));
    if (code === 200) pushed += chunk.length;
    Utilities.sleep(300);
  }

  Logger.log('=== WIP Push Done: ' + pushed + '/' + records.length + ' ===');
}



function masterPush() {
  addLobCol();
  fillWipCols();
  wipeAndPushAll();
}
