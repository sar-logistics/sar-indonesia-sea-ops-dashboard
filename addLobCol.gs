// ── Add Lob + Derived Margin Columns ──────────────────────────────────────
// Run addLobCol() — safe to re-run, checks if columns already exist

function addLobCol() {
  const ss = SpreadsheetApp.openById('1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY');

  ['Shipment Profile Export', 'Shipment Profile Import'].forEach(tabName => {
    const direction = tabName.includes('Export') ? 'Export' : 'Import';
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('Tab not found: ' + tabName); return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // LOB column
    let lobCol = headers.indexOf('Lob') + 1;
    if (lobCol === 0) {
      lobCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lobCol).setValue('Lob');
      Logger.log(direction + ': Lob created at col ' + lobCol);
    } else {
      Logger.log(direction + ': Lob exists at col ' + lobCol);
    }

    // Derived Margin column
    let dmCol = headers.indexOf('Derived Margin') + 1;
    if (dmCol === 0) {
      dmCol = lobCol + 1;
      sheet.getRange(1, dmCol).setValue('Derived Margin');
      Logger.log(direction + ': Derived Margin created at col ' + dmCol);
    } else {
      Logger.log(direction + ': Derived Margin exists at col ' + dmCol);
    }

    const transIdx    = headers.indexOf('Trans');
    const jobProfIdx  = headers.indexOf('Job Profit');
    const totIncomeIdx = headers.indexOf('Total Income (Recognized+Unrecognized REV+WIP)');

    const data = sheet.getDataRange().getValues();
    let lobFilled = 0, marginFilled = 0;
    let seaCount = 0, airCount = 0, roadCount = 0, blankCount = 0;

    for (let i = 1; i < data.length; i++) {
      const trans = String(data[i][transIdx] || '').trim().toUpperCase();

      // Always recalculate LOB from current Trans and overwrite, every run
      let lob = '';
      if (trans === 'SEA') { lob = direction === 'Export' ? 'FES' : 'FIS'; seaCount++; }
      else if (trans === 'AIR') { lob = direction === 'Export' ? 'FEA' : 'FIA'; airCount++; }
      else if (trans === 'ROA') { lob = direction === 'Export' ? 'FER' : 'FIR'; roadCount++; }
      else { blankCount++; }
      sheet.getRange(i + 1, lobCol).setValue(lob);
      lobFilled++;

      // Fill Derived Margin always
      const jobProfit = parseFloat(data[i][jobProfIdx]) || 0;
      const income    = parseFloat(data[i][totIncomeIdx]) || 0;
      const margin    = income !== 0 ? (jobProfit / income) * 100 : 0;
      sheet.getRange(i + 1, dmCol).setValue(parseFloat(margin.toFixed(2)));
      marginFilled++;
    }

    Logger.log(direction + ': Lob=' + lobFilled + ' rows, DerivedMargin=' + marginFilled + ' rows');
    Logger.log(direction + ' breakdown -> Sea: ' + seaCount + ' | Air: ' + airCount + ' | Road: ' + roadCount + ' | Blank/Unmatched Trans: ' + blankCount);
    if (blankCount > 0) {
      Logger.log('⚠ ' + direction + ': ' + blankCount + ' row(s) had a Trans value other than SEA/AIR/ROAD (or blank) — Lob left empty for those.');
    }
  });

  Logger.log('=== addLobCol Done ===');
}
