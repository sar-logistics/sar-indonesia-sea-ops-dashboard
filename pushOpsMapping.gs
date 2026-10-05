// ── SAR Indonesia Sea Ops — Push Ops Mapping (Users / Reporting Hierarchy) ─
// Separate script (not part of pushToMongo.gs). Run pushOpsMapping() any
// time after editing the 'Ops Mapping' tab to resync users in MongoDB.
//
// Reads the 'Ops Mapping' tab and upserts each row into MongoDB's users
// collection, matching exactly what the Admin Panel's Add/Edit User form
// writes (same fields, same shape) so these users work identically to ones
// added manually through the UI.
//
// Sheet columns expected: Job Owner | Persona | Job Owner Email | Reporting To | Reporting To - Email
// 'Reporting To - Email' (not the name column) is stored as reportingManager,
// since the dashboard looks up managers by email, not by name.

const OPS_MAP_SHEET_ID  = '1eTuXf5ngxTkJrKPVfXrZiqKw6XUUPam1P4ig7cmV3BY';
const USERS_API_URL     = 'https://sar-ind-sea-ops-dashboard.vercel.app/api/mongo';

// Persona (sheet) -> role (dashboard). Add more mappings here as new
// personas are introduced; anything not listed is pushed through as-is.
const PERSONA_ROLE_MAP = {
  'ops head': 'Ops Head',
  'docs':     'Docs',
  'cs':       'CS',
};

function pushOpsMapping() {
  Logger.log('=== Ops Mapping Push Starting ===');
  const ss = SpreadsheetApp.openById(OPS_MAP_SHEET_ID);
  const sheet = ss.getSheetByName('Ops Mapping');
  if (!sheet) { Logger.log('ERROR: Ops Mapping tab not found'); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows    = data.slice(1);

  const h = (name) => headers.indexOf(name);
  const NAME            = h('Job Owner');
  const PERSONA         = h('Persona');
  const EMAIL           = h('Job Owner Email');
  const REPORTING_EMAIL = h('Reporting To - Email');

  if (NAME === -1 || EMAIL === -1) {
    Logger.log('ERROR: Required columns (Job Owner, Job Owner Email) not found');
    return;
  }

  let pushed = 0, failed = 0;
  rows.forEach(row => {
    const name = String(row[NAME] || '').trim();
    const email = String(row[EMAIL] || '').trim().toLowerCase();
    if (!name || !email) return;

    const personaRaw = String(row[PERSONA] || '').trim();
    const role = PERSONA_ROLE_MAP[personaRaw.toLowerCase()] || personaRaw;
    const reportingManager = String(row[REPORTING_EMAIL] || '').trim().toLowerCase();

    try {
      const resp = UrlFetchApp.fetch(USERS_API_URL, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'updateOne',
          collection: 'users',
          filter: { email },
          update: {
            $set: { name, email, role, reportingManager, active: true, updatedAt: new Date().toISOString() },
            $setOnInsert: { zones: [], branches: [], phone: '', createdAt: new Date().toISOString() }
          }
        }),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      Logger.log(name + ' (' + email + ', role=' + role + '): HTTP ' + code);
      if (code === 200) pushed++; else failed++;
    } catch(e) {
      Logger.log('ERROR pushing ' + name + ': ' + e.message);
      failed++;
    }
  });

  Logger.log('=== Ops Mapping Push Done: ' + pushed + ' pushed, ' + failed + ' failed ===');
}
