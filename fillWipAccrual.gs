// ── SAR Indonesia Sea Ops — Fill WIP/Accrual Sheet ─────────────────────────
// Separate script (not part of pushToMongo.gs). Run fillWipCols() before
// every data push — it fills the WIP sheet's ETD, ETA, LOB, ETD/ETA, and
// Job Operator columns by cross-referencing the Export/Import tabs.
//
// ETD (AE) = lookup Shipment ID in Export tab → Origin ETD
// ETA (AF) = lookup Shipment ID in Import tab → Destination ETA
// LOB (AG) = FES/FIS/FEA/FIA based on Trans + which tab the job is in
// ETD/ETA (AH) = ETD if Export LOB, ETA if Import LOB
// Job Operator (AI) = lookup Shipment ID in Export/Import tab → Job Operator

function fillWipCols() {
  const ss = SpreadsheetApp.openById('1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY');

  const wipSheet = ss.getSheetByName('WIP, ACCURAL');
  const expSheet = ss.getSheetByName('Shipment Profile Export');
  const impSheet = ss.getSheetByName('Shipment Profile Import');

  if (!wipSheet || !expSheet || !impSheet) {
    Logger.log('ERROR: Sheet not found'); return;
  }

  const expMap = buildMap(expSheet);
  const impMap = buildMap(impSheet);

  Logger.log('Export map: ' + Object.keys(expMap).length);
  Logger.log('Import map: ' + Object.keys(impMap).length);

  const wipData = wipSheet.getDataRange().getValues();
  const numRows = wipData.length - 1;
  Logger.log('WIP rows: ' + numRows);

  const JOB_COL    = 0;  // A = Job #
  const TRANS_COL  = 2;  // C = Trans
  const ETD_COL    = 30; // AE
  const ETA_COL    = 31; // AF
  const LOB_COL    = 32; // AG
  const ETDETA_COL = 33; // AH
  const JOBOP_COL  = 34; // AI = Job Operator

  const etdVals = [], etaVals = [], lobVals = [], etdetaVals = [], opVals = [];

  for (let i = 1; i < wipData.length; i++) {
    const row   = wipData[i];
    const jobNo = String(row[JOB_COL]   || '').trim();
    const trans = String(row[TRANS_COL] || '').trim().toUpperCase();

    const rec   = expMap[jobNo] || impMap[jobNo] || {};
    const inExp = !!expMap[jobNo];
    const inImp = !!impMap[jobNo];

    const etd      = rec.originEtd      || '';
    const eta      = rec.destinationEta || '';
    const operator = rec.operator       || '';

    let lob = '';
    if      (trans === 'SEA' && inExp) lob = 'FES';
    else if (trans === 'SEA' && inImp) lob = 'FIS';
    else if (trans === 'AIR' && inExp) lob = 'FEA';
    else if (trans === 'AIR' && inImp) lob = 'FIA';

    const etdeta = (lob === 'FES' || lob === 'FEA') ? etd :
                   (lob === 'FIS' || lob === 'FIA') ? eta : '';

    etdVals.push([etd]);
    etaVals.push([eta]);
    lobVals.push([lob]);
    etdetaVals.push([etdeta]);
    opVals.push([operator]);
  }

  if (numRows > 0) {
    wipSheet.getRange(2, ETD_COL+1,    numRows, 1).setValues(etdVals);
    wipSheet.getRange(2, ETA_COL+1,    numRows, 1).setValues(etaVals);
    wipSheet.getRange(2, LOB_COL+1,    numRows, 1).setValues(lobVals);
    wipSheet.getRange(2, ETDETA_COL+1, numRows, 1).setValues(etdetaVals);
    wipSheet.getRange(2, JOBOP_COL+1,  numRows, 1).setValues(opVals);
  }

  Logger.log('=== fillWipCols Done: ' + numRows + ' rows processed ===');
}

function buildMap(sheet) {
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const map     = {};

  const shipIdx = 0; // Shipment ID always col A
  const etdIdx  = headers.findIndex(h => h === 'Origin ETD');
  const etaIdx  = headers.findIndex(h => h === 'Destination ETA' || h.startsWith('Destination E'));
  const opIdx   = headers.findIndex(h => h === 'Job Operator');

  Logger.log(sheet.getName() + ': ETD col=' + etdIdx + ' ETA col=' + etaIdx + ' Operator col=' + opIdx);

  for (let i = 1; i < data.length; i++) {
    const shipId = String(data[i][shipIdx] || '').trim();
    if (!shipId || shipId.startsWith('#') || shipId === 'Shipment ID') continue;
    map[shipId] = {
      originEtd:      etdIdx > -1 ? formatDate(data[i][etdIdx]) : '',
      destinationEta: etaIdx > -1 ? formatDate(data[i][etaIdx]) : '',
      operator:       opIdx  > -1 ? String(data[i][opIdx] || '').trim() : '',
    };
  }
  return map;
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }
  return String(val).trim();
}
