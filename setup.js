/**
 * Postage Forecast Portal — setup.gs
 *
 * Creates the 30 tabs and seeds reference data. Every function here is
 * IDEMPOTENT: safe to run again, skips what already exists, never overwrites
 * data you have entered.
 *
 * Run order (or just run setupAll):
 *   1. setupCreateAllTabs()
 *   2. setupSeedReferenceData()
 *   3. setupBuildCalendar()
 *   4. setupVerify()
 */

// ─────────────────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────────

function setupAll() {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log('=== Postage Forecast Portal — setup ===');
  setupCreateAllTabs();
  setupSeedReferenceData();
  setupBuildCalendar();
  const r = setupVerify();
  Logger.log('=== Done in ' + (Date.now() - t0) + 'ms ===');
  return r;
}


// ─────────────────────────────────────────────────────────────────────────────
// STEP 1 — CREATE TABS
// ─────────────────────────────────────────────────────────────────────────────

function setupCreateAllTabs() {
  requireMaintenance_();
  if (spreadsheetId_() === 'PASTE_YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Set SPREADSHEET_ID_FALLBACK at the top of utils.gs, or a SPREADSHEET_ID ' +
                    'Script Property, before running setup. Run showEnvironment() to check.');
  }

  const ss = _getSs_();
  let created = 0, updated = 0, skipped = 0;

  for (const key in TABLES) {
    const t = TABLES[key];
    let sh = ss.getSheetByName(t.sheet);

    if (!sh) {
      sh = ss.insertSheet(t.sheet);
      writeHeaderRow_(sh, t);
      Logger.log('Created tab: ' + t.sheet + '  (' + t.headers.length + ' columns)');
      created++;
    } else {
      const existing = sh.getLastColumn()
        ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeStr)
        : [];
      const missing = t.headers.filter(h => existing.indexOf(h) < 0);
      if (missing.length) {
        // Append only new columns by name — never reorder, never destroy data.
        let col = existing.length + 1;
        const needed = col + missing.length - 1;
        if (sh.getMaxColumns() < needed) sh.insertColumnsAfter(sh.getMaxColumns(), needed - sh.getMaxColumns());
        missing.forEach(h => {
          sh.getRange(1, col).setValue(h).setFontWeight('bold');
          const f = formatForColumn_(h);
          if (f) sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat(f);
          col++;
        });
        Logger.log('Updated tab: ' + t.sheet + '  (added ' + missing.join(', ') + ')');
        updated++;
      } else {
        skipped++;
      }
    }
  }

  // Remove the default "Sheet1" if it is still empty and we created real tabs.
  const dflt = ss.getSheetByName('Sheet1');
  if (dflt && ss.getSheets().length > 1 && dflt.getLastRow() === 0 && dflt.getLastColumn() === 0) {
    ss.deleteSheet(dflt);
    Logger.log('Removed empty default Sheet1');
  }

  Logger.log('Tabs: ' + created + ' created, ' + updated + ' updated, ' + skipped + ' already correct');
  return { created: created, updated: updated, skipped: skipped };
}

function writeHeaderRow_(sh, t) {
  const n = t.headers.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  sh.getRange(1, 1, 1, n).setValues([t.headers])
    .setFontWeight('bold')
    .setBackground('#F0F0FF')      // brand tint
    .setFontColor('#0B0B0F');
  sh.setFrozenRows(1);

  for (let c = 0; c < n; c++) {
    const f = formatForColumn_(t.headers[c]);
    if (f) sh.getRange(2, c + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat(f);
  }
  try { sh.setColumnWidths(1, n, t.width || 130); } catch (e) { /* narrow sheets */ }

  // Trim unused columns so the tab is not 26 wide when it needs 11.
  const extra = sh.getMaxColumns() - n;
  if (extra > 0) sh.deleteColumns(n + 1, extra);
}


// ─────────────────────────────────────────────────────────────────────────────
// STEP 2 — SEED REFERENCE DATA
//
// All values below were extracted from Postage_RFQ3_New.xlsx, not typed by hand.
// ─────────────────────────────────────────────────────────────────────────────

const SEED_REFERENCE = [
  ['BRAND', 'MEDEXPRESS',  'MedExpress',  10],
  ['BRAND', 'DERMATICA',   'Dermatica',   20],
  ['BRAND', 'ZIPHEALTH',   'ZipHealth',   30],
  ['BRAND', 'ROCKETRX',    'RocketRx',    40],
  ['BRAND', 'LEVITY',      'Levity',      50],

  ['GEO', 'GB', 'United Kingdom', 10],
  ['GEO', 'US', 'United States',  20],
  ['GEO', 'DE', 'Germany',        30],
  ['GEO', 'CA', 'Canada',         40],

  ['TREATMENT_TYPE', 'WL',      'Weight Loss', 10],
  ['TREATMENT_TYPE', 'CORE_RX', 'Core Rx',     20],

  ['WL_SPLIT', 'MOUNJARO',   'Mounjaro',       10],
  ['WL_SPLIT', 'WEGOVY',     'WeGovy',         20],
  ['WL_SPLIT', 'BRANDED',    'Branded',        30],
  ['WL_SPLIT', 'COMPOUNDED', 'Compounded',     40],
  ['WL_SPLIT', '*',          'Not applicable', 90],

  ['LETTER_PARCEL', 'LETTER', 'Letter', 10],
  ['LETTER_PARCEL', 'PARCEL', 'Parcel', 20],

  ['TEMP_REGIME', 'CC',      'Cold chain', 10],
  ['TEMP_REGIME', 'AMBIENT', 'Ambient',    20],

  ['CURRENCY', 'GBP', 'Pound sterling', 10],
  ['CURRENCY', 'USD', 'US dollar',      20],
  ['CURRENCY', 'EUR', 'Euro',           30],
  ['CURRENCY', 'CAD', 'Canadian dollar',40]
];

const SEED_CARRIERS = [
  ['ROYALMAIL', 'Royal Mail', 'GBP'],
  ['DPD',       'DPD',        'GBP'],
  ['DHL',       'DHL',        'EUR'],
  ['FEDEX',     'FedEx',      'USD'],
  ['USPS',      'USPS',       'USD'],
  ['OTHER',     'Other',      'GBP'],
  ['TBC',       'TBC',        'GBP']
];

const SEED_METHODS = [
  ['DHLEXPRESS','DHL','DHLExpress'],
  ['DHLSTANDARD','DHL','DHLStandard'],
  ['DPDNEXTDAY','DPD','DPDNextDay'],
  ['FEDEXPRIORITYOVERNIGHT','FEDEX','FedExPriorityOvernight'],
  ['FEDEXSTANDARDOVERNIGHT','FEDEX','FedExStandardOvernight'],
  ['FEDEXTWODAY','FEDEX','FedExTwoDay'],
  ['GOPHR','OTHER','Gophr'],
  ['SHUTLIT','OTHER','ShutlIt'],
  ['UKMAIL','OTHER','UKMail'],
  ['USAMAIL','OTHER','USAMail'],
  ['FIRSTCLASSNOSIG','ROYALMAIL','FirstClassNoSig'],
  ['INTERNATIONAL','ROYALMAIL','International'],
  ['RM24','ROYALMAIL','RM24'],
  ['RM48','ROYALMAIL','RM48'],
  ['ROYALMAIL9AM','ROYALMAIL','RoyalMail9am'],
  ['ROYALMAILFIRSTCLASS','ROYALMAIL','RoyalMailFirstClass'],
  ['ROYALMAILINTERNATIONAL','ROYALMAIL','RoyalMailInternational'],
  ['ROYALMAILINTERNATIONALSIGNED','ROYALMAIL','RoyalMailInternationalSigned'],
  ['ROYALMAILNEXTDAY','ROYALMAIL','RoyalMailNextDay'],
  ['ROYALMAILSATURDAYGUARANTEED1PM','ROYALMAIL','RoyalMailSaturdayGuaranteed1pm'],
  ['ROYALMAILSATURDAYGUARANTEED9AM','ROYALMAIL','RoyalMailSaturdayGuaranteed9am'],
  ['ROYALMAILTRACKED24','ROYALMAIL','RoyalMailTracked24'],
  ['ROYALMAILTRACKED24LETTERBOX','ROYALMAIL','RoyalMailTracked24Letterbox'],
  ['ROYALMAILTRACKED24NOSIGNATURE','ROYALMAIL','RoyalMailTracked24NoSignature'],
  ['ROYALMAILTRACKED48','ROYALMAIL','RoyalMailTracked48'],
  ['ROYALMAILTRACKED48LETTERBOX','ROYALMAIL','RoyalMailTracked48Letterbox'],
  ['SPECIALDELIVERY1PM','ROYALMAIL','SpecialDelivery1PM'],
  ['TRACKED24','ROYALMAIL','Tracked24'],
  ['PHARMACYCOLLECT','TBC','PharmacyCollect'],
  ['STANDARD','TBC','Standard'],
  ['UPSEXPRESS','USPS','UPSExpress'],
  ['USPSFIRST','USPS','USPSFirst'],
  ['USPSGROUNDADVANTAGE','USPS','USPSGroundAdvantage'],
  ['USPSPRIORITY','USPS','USPSPriority'],
  ['USPSPRIORITYEXPRESS','USPS','USPSPriorityExpress']
];

// Value_Type, Applies_To, Apply_Order, Proration — matches the workbook exactly.
const SEED_SURCHARGES = [
  ['FUEL',        'Fuel surcharge', 'PCT', 'BASE', 1, 'DAY_WEIGHTED'],
  ['GREEN',       'Green levy',     'AMT', 'BASE', 1, 'DAY_WEIGHTED'],
  ['PEAK',        'Peak surcharge', 'AMT', 'BASE', 2, 'DAY_WEIGHTED'],
  ['SHIPSTATION', 'ShipStation fee','AMT', 'BASE', 3, 'DAY_WEIGHTED']
];

//        Role,      All,   Write, Rates, Mixes, Struct, Calc,  Publish, Users, Audit
const SEED_ROLES = [
  ['ADMIN',    true,  true,  true,  true,  true,  true,  true,  true,  true ],
  ['MODELLER', false, true,  true,  true,  true,  true,  true,  false, true ],
  ['ANALYST',  false, true,  false, true,  false, true,  false, false, false],
  ['VIEWER',   false, false, false, false, false, false, false, false, false]
];

/**
 * Config seed rows.
 *
 * A function rather than a `const` like every seed above it, because this is the
 * one seed that reads a value declared in another file: ENGINE_VERSION, in
 * utils.gs.
 *
 * Apps Script evaluates script files in the project's file order, in one shared
 * scope, and a top-level `const`/`let` is in the temporal dead zone until its own
 * file has been evaluated. So a top-level array literal here throws
 * `ReferenceError: ENGINE_VERSION is not defined` whenever setup.gs is ordered
 * before utils.gs — which is what a plain `clasp push` produces, because it
 * pushes alphabetically and "setup" sorts before "utils". The error fires at file
 * load, before any function runs, so it takes down every entry point including
 * doGet, and no amount of reverting file *contents* fixes a file *order* problem.
 *
 * Building the rows inside a function defers the read until every file is loaded,
 * which makes file order irrelevant. `.clasp.json` also pins utils.js first as
 * defence in depth. See FINDINGS.md M9 — and please don't turn this back into a
 * const.
 */
function seedConfigRows_() {
  return [
    ['HORIZON_START',             '2026-01-01', 'First month of the forecast horizon'],
    ['HORIZON_MONTHS',            '36',         'Number of months in the horizon'],
    ['OPEN_ENDED_DATE',           '9999-12-31', 'Sentinel Valid_To meaning no end date'],
    ['AUTO_CALC_ON_WRITE',        'FALSE',      'Recalculate automatically after every write'],
    ['VALIDATION_BLOCKS_PUBLISH', 'TRUE',       'Block publishing when validation returns ERROR'],
    ['ENGINE_VERSION',            ENGINE_VERSION, 'Version stamped onto each calculation run'],
    ['REPORTING_CURRENCY',        'GBP',        'Label only — no FX conversion is applied'],
    ['BOOTSTRAP_OWNER_EMAIL',     '',           'Escape hatch: this email always has ADMIN'],
    ['SCOPE_DEFAULT_ALLOW',       'TRUE',       'What a user with no Scope_Mapping rows can see. TRUE = everything'],
    ['METABASE_URL',              '',           'e.g. https://metabase.heliosx.co — leave blank until tested'],
    ['METABASE_CARD_ID',          '',           'The saved question ID that returns actual spend and orders'],
    ['ACTUALS_VARIANCE_WARN_PCT', '10',         'Flag a forecast-vs-actual gap larger than this percentage'],
    ['ACTUALS_WL_TREATMENTS',     'WeightLossGlp1', 'Treatment values in the import that count as weight loss. Comma separated.']
  ];
}


function setupSeedReferenceData() {
  requireMaintenance_();
  let n = 0;
  n += seedTable_('DIM_REFERENCE', 'Code', SEED_REFERENCE, (r, id) => {
    const row = blankRow_('DIM_REFERENCE'), C = COL.DIM_REFERENCE;
    row[C.Ref_ID] = id;      row[C.List_Name]  = r[0];
    row[C.Code]   = r[1];    row[C.Label]      = r[2];
    row[C.Sort_Order] = r[3]; row[C.Active]    = true;
    return row;
  }, r => r[0] + '|' + r[1], d => safeStr(d[COL.DIM_REFERENCE.List_Name]) + '|' + safeStr(d[COL.DIM_REFERENCE.Code]));

  n += seedTable_('DIM_CARRIER', 'Carrier_Code', SEED_CARRIERS, (r) => {
    const row = blankRow_('DIM_CARRIER'), C = COL.DIM_CARRIER;
    row[C.Carrier_Code] = r[0]; row[C.Carrier_Name] = r[1];
    row[C.Default_Currency] = r[2]; row[C.Active] = true;
    return row;
  }, r => r[0], d => safeStr(d[COL.DIM_CARRIER.Carrier_Code]));

  n += seedTable_('DIM_METHOD', 'Method_Code', SEED_METHODS, (r) => {
    const row = blankRow_('DIM_METHOD'), C = COL.DIM_METHOD;
    row[C.Method_Code] = r[0]; row[C.Carrier_Code] = r[1];
    row[C.Method_Name] = r[2]; row[C.Is_Tracked] = false; row[C.Active] = true;
    return row;
  }, r => r[0], d => safeStr(d[COL.DIM_METHOD.Method_Code]));

  n += seedTable_('DIM_SURCHARGE', 'Surcharge_Code', SEED_SURCHARGES, (r) => {
    const row = blankRow_('DIM_SURCHARGE'), C = COL.DIM_SURCHARGE;
    row[C.Surcharge_Code] = r[0]; row[C.Surcharge_Name] = r[1];
    row[C.Value_Type] = r[2];     row[C.Applies_To] = r[3];
    row[C.Apply_Order] = r[4];    row[C.Proration] = r[5];
    row[C.Active] = true;
    return row;
  }, r => r[0], d => safeStr(d[COL.DIM_SURCHARGE.Surcharge_Code]));

  n += seedTable_('PORTAL_ROLES', 'Role', SEED_ROLES, (r) => {
    const row = blankRow_('PORTAL_ROLES'), C = COL.PORTAL_ROLES;
    row[C.Role] = r[0];               row[C.All_Access] = r[1];
    row[C.Write_Access] = r[2];       row[C.Can_Edit_Rates] = r[3];
    row[C.Can_Edit_Mixes] = r[4];     row[C.Can_Edit_Structure] = r[5];
    row[C.Can_Run_Calc] = r[6];       row[C.Can_Publish_Output] = r[7];
    row[C.Can_Manage_Users] = r[8];   row[C.Can_View_Audit] = r[9];
    return row;
  }, r => r[0], d => safeStr(d[COL.PORTAL_ROLES.Role]));

  n += seedTable_('CONFIG', 'Key', seedConfigRows_(), (r) => {
    const row = blankRow_('CONFIG'), C = COL.CONFIG;
    row[C.Key] = r[0]; row[C.Value] = r[1]; row[C.Description] = r[2];
    row[C.Updated_TS] = new Date(); row[C.Updated_By] = 'setup';
    return row;
  }, r => r[0], d => safeStr(d[COL.CONFIG.Key]));

  n += seedTable_('SCENARIOS', 'Scenario_ID', [['BASE', 'Base case — the live forecast']], (r) => {
    const row = blankRow_('SCENARIOS'), C = COL.SCENARIOS;
    row[C.Scenario_ID] = 1;  row[C.Scenario_Name] = r[0];
    row[C.Description] = r[1]; row[C.Is_Default] = true;
    row[C.Locked] = false;   row[C.Active] = true;
    row[C.Created_TS] = new Date(); row[C.Created_By] = 'setup';
    return row;
  }, r => r[0], d => safeStr(d[COL.SCENARIOS.Scenario_Name]));

  Logger.log('Reference data: ' + n + ' rows inserted (existing rows untouched)');
  return n;
}

/**
 * Insert seed rows that are not already present, matched on a natural key.
 * Never updates or deletes, so running twice is a no-op.
 */
function seedTable_(tableKey, idHeader, seeds, buildRow, seedKeyFn, rowKeyFn) {
  const t  = TABLES[tableKey];
  const sh = getSheet_(t.sheet);
  const data = sh.getDataRange().getValues();

  const present = {};
  for (let i = 1; i < data.length; i++) {
    const k = rowKeyFn(data[i]);
    if (k) present[normKey(k)] = true;
  }

  let nextId = 0;
  const idCol = t.headers.indexOf(idHeader);
  if (idCol >= 0 && /_ID$/.test(idHeader)) {
    for (let i = 1; i < data.length; i++) {
      const v = parseInt(data[i][idCol], 10);
      if (!isNaN(v) && v > nextId) nextId = v;
    }
  }

  const toAdd = [];
  seeds.forEach(s => {
    if (present[normKey(seedKeyFn(s))]) return;
    toAdd.push(buildRow(s, ++nextId));
  });

  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, t.headers.length).setValues(toAdd);
    Logger.log('  ' + t.sheet + ': +' + toAdd.length + ' rows');
  }
  invalidateSheetCache_(t.sheet);
  return toAdd.length;
}


// ─────────────────────────────────────────────────────────────────────────────
// STEP 3 — BUILD THE CALENDAR
//
// Generated from Config, so extending the horizon is a config change plus a
// re-run of this function. Existing rows are preserved; only new months append.
// ─────────────────────────────────────────────────────────────────────────────

function setupBuildCalendar() {
  requireMaintenance_();
  const start  = normaliseDate(configDate('HORIZON_START', '2026-01-01'));
  const months = configInt('HORIZON_MONTHS', 36);
  if (!start) throw new Error('HORIZON_START in Config is not a valid date.');

  const t  = TABLES.DIM_CALENDAR;
  const sh = getSheet_(t.sheet);
  const C  = COL.DIM_CALENDAR;
  const data = sh.getDataRange().getValues();

  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const k = dateKey(data[i][C.Month_Start]);
    if (k) existing[k] = true;
  }

  const rows = [];
  for (let i = 0; i < months; i++) {
    const ms = new Date(start.getFullYear(), start.getMonth() + i, 1);
    if (existing[dateKey(ms)]) continue;
    const me = new Date(ms.getFullYear(), ms.getMonth() + 1, 0);
    const row = blankRow_('DIM_CALENDAR');
    row[C.Date_ID]       = i + 1;
    row[C.Month_Start]   = ms;
    row[C.Month_End]     = me;
    row[C.Days_In_Month] = me.getDate();
    row[C.Year]          = ms.getFullYear();
    row[C.Month_No]      = ms.getMonth() + 1;
    row[C.Month_Label]   = fmtMonthLabel(ms);
    row[C.Quarter]       = ms.getFullYear() + '-Q' + (Math.floor(ms.getMonth() / 3) + 1);
    row[C.In_Horizon]    = true;
    rows.push(row);
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, t.headers.length).setValues(rows);
  }
  invalidateSheetCache_(t.sheet);
  Logger.log('Calendar: ' + rows.length + ' months added, ' +
             Object.keys(existing).length + ' already present ' +
             '(' + fmtDate(start) + ' + ' + months + ' months)');
  return rows.length;
}


// ─────────────────────────────────────────────────────────────────────────────
// STEP 4 — VERIFY
// ─────────────────────────────────────────────────────────────────────────────

function setupVerify() {
  requireMaintenance_();
  const ss = _getSs_();
  const problems = [];
  let ok = 0;

  for (const key in TABLES) {
    const t  = TABLES[key];
    const sh = ss.getSheetByName(t.sheet);
    if (!sh) { problems.push('MISSING TAB: ' + t.sheet); continue; }

    const hdr = sh.getLastColumn()
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeStr)
      : [];
    const missing = t.headers.filter(h => hdr.indexOf(h) < 0);
    if (missing.length) problems.push(t.sheet + ': missing columns ' + missing.join(', '));
    else ok++;
  }

  const counts = {
    Dim_Reference: countRows_(SHEET.DIM_REFERENCE),
    Dim_Carrier:   countRows_(SHEET.DIM_CARRIER),
    Dim_Method:    countRows_(SHEET.DIM_METHOD),
    Dim_Surcharge: countRows_(SHEET.DIM_SURCHARGE),
    Dim_Calendar:  countRows_(SHEET.DIM_CALENDAR),
    Portal_Roles:  countRows_(SHEET.PORTAL_ROLES),
    Scenarios:     countRows_(SHEET.SCENARIOS),
    Config:        countRows_(SHEET.CONFIG)
  };

  const expected = { Dim_Reference: 24, Dim_Carrier: 7, Dim_Method: 35, Dim_Surcharge: 4,
                     Dim_Calendar: configInt('HORIZON_MONTHS', 36), Portal_Roles: 4,
                     Scenarios: 1, Config: 13 };

  Logger.log('');
  Logger.log('--- VERIFICATION ---');
  Logger.log('Tabs correct: ' + ok + ' of ' + Object.keys(TABLES).length);
  for (const k in counts) {
    const flag = (counts[k] === expected[k]) ? 'OK  ' : 'CHECK';
    Logger.log('  ' + flag + ' ' + k + ': ' + counts[k] + ' rows (expected ' + expected[k] + ')');
    if (counts[k] !== expected[k]) problems.push(k + ': ' + counts[k] + ' rows, expected ' + expected[k]);
  }

  // Referential check on the seed data itself.
  const carriers = {};
  const cData = getAllData_(SHEET.DIM_CARRIER);
  for (let i = 1; i < cData.length; i++) carriers[safeStr(cData[i][COL.DIM_CARRIER.Carrier_Code])] = true;
  const mData = getAllData_(SHEET.DIM_METHOD);
  for (let i = 1; i < mData.length; i++) {
    const cc = safeStr(mData[i][COL.DIM_METHOD.Carrier_Code]);
    if (cc && !carriers[cc]) problems.push('Dim_Method row ' + (i + 1) + ': unknown Carrier_Code ' + cc);
  }

  if (problems.length) {
    Logger.log('');
    Logger.log('PROBLEMS (' + problems.length + '):');
    problems.forEach(p => Logger.log('  - ' + p));
  } else {
    Logger.log('');
    Logger.log('ALL CHECKS PASSED — ready for step 2 (migration).');
  }
  return { ok: ok, problems: problems, counts: counts };
}

function countRows_(sheetName) {
  const sh = _getSs_().getSheetByName(sheetName);
  return sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
}


/**
 * Rewrites Month_Label and Quarter as plain text.
 *
 * Needed because Sheets parses "Jan-26" as a date the moment it is written into
 * an unformatted cell, turning the label into a serial number. Sets the column
 * format to text first, then rewrites the values. Safe to run any time.
 */
/**
 * Rewrite the Config Value column as plain text.
 *
 * Sheets parses 2026-01-01 into a date the moment it lands in an unformatted
 * cell, so HORIZON_START came back as a serial number. Setting the column format
 * first, then rewriting the values, fixes both the existing rows and any future
 * edit made directly in the sheet.
 */
function repairConfigValues() {
  requireMaintenance_();
  const t = TABLES.CONFIG, C = COL.CONFIG, sh = getSheet_(t.sheet);
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('Config is empty.'); return 0; }

  sh.getRange(2, C.Value + 1, last - 1, 1).setNumberFormat('@');

  const rows = sh.getRange(2, 1, last - 1, t.headers.length).getValues();
  const out = [], fixed = [];
  for (let i = 0; i < rows.length; i++) {
    const key = safeStr(rows[i][C.Key]);
    let v = rows[i][C.Value];
    if (isDate_(v)) {
      const s = fmtDate(v);
      fixed.push(key + ': date -> "' + s + '"');
      v = s;
    } else if (typeof v === 'number' && v > SERIAL_EPOCH_OFFSET && v < 2958466 &&
               /DATE|START/.test(key.toUpperCase())) {
      const s = fmtDate(v);
      fixed.push(key + ': ' + rows[i][C.Value] + ' -> "' + s + '"');
      v = s;
    } else {
      v = safeStr(v);
    }
    out.push([v]);
  }
  sh.getRange(2, C.Value + 1, out.length, 1).setValues(out);
  invalidateSheetCache_(t.sheet);
  _configCache_ = null;

  Logger.log('=== CONFIG VALUES ===');
  Logger.log('  ' + out.length + ' value(s) rewritten as text');
  if (fixed.length) { Logger.log('  corrected:'); fixed.forEach(f => Logger.log('    ' + f)); }
  else Logger.log('  nothing needed correcting');
  Logger.log('');
  Logger.log('  HORIZON_START now reads: ' + configDate('HORIZON_START', '2026-01-01'));
  return fixed.length;
}


function repairCalendarLabels() {
  requireMaintenance_();
  const sh = getSheet_(SHEET.DIM_CALENDAR), C = COL.DIM_CALENDAR;
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('Dim_Calendar is empty — run setupBuildCalendar first.'); return 0; }

  sh.getRange(2, C.Month_Label + 1, last - 1, 1).setNumberFormat('@');
  sh.getRange(2, C.Quarter + 1,     last - 1, 1).setNumberFormat('@');

  const starts = sh.getRange(2, C.Month_Start + 1, last - 1, 1).getValues();
  const labels = [], quarters = [];
  for (let i = 0; i < starts.length; i++) {
    const d = normaliseDate(starts[i][0]);
    labels.push([d ? fmtMonthLabel(d) : '']);
    quarters.push([d ? d.getFullYear() + '-Q' + (Math.floor(d.getMonth() / 3) + 1) : '']);
  }
  sh.getRange(2, C.Month_Label + 1, labels.length, 1).setValues(labels);
  sh.getRange(2, C.Quarter + 1,     quarters.length, 1).setValues(quarters);
  invalidateSheetCache_(SHEET.DIM_CALENDAR);

  Logger.log('Repaired ' + labels.length + ' calendar labels.');
  Logger.log('  first: ' + labels[0][0] + '   last: ' + labels[labels.length - 1][0]);
  return labels.length;
}


/**
 * Resolve the validation warnings that have a single obvious correct answer.
 *
 * Two of the three do:
 *
 *   METHOD_CARRIER  a route filed under the wrong carrier. Dim_Method already
 *                   says which carrier owns each method, so the route is moved
 *                   to match. Your original workbook had DHLExpress filed under
 *                   USPS on two routes.
 *
 *   CC_COVERAGE     a segment with no cold-chain rows is already treated as
 *                   100% ambient, but implicitly. Writing an explicit 0% row
 *                   makes it a stated assumption rather than a silence.
 *
 * TBC_CARRIER is left alone. Which courier those eight routes should actually
 * use is a commercial decision, and guessing would be worse than the warning.
 */
function fixValidationWarnings() {
  requireMaintenance_();
  Logger.log('=== RESOLVING VALIDATION WARNINGS ===');
  const email = getActiveEmail_() || 'system';
  let changed = 0;

  // ---- METHOD_CARRIER --------------------------------------------------
  Logger.log('');
  Logger.log('--- routes filed under the wrong carrier ---');
  const owner = {};
  const dm = getAllData_(SHEET.DIM_METHOD), DM = COL.DIM_METHOD;
  for (let i = 1; i < dm.length; i++) {
    owner[safeStr(dm[i][DM.Method_Code])] = safeStr(dm[i][DM.Carrier_Code]);
  }

  const t = TABLES.MODELLING_IDS, M = COL.MODELLING_IDS, width = t.headers.length;
  const sh = getSheet_(t.sheet);
  const last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, width).getValues();
    const changes = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const mc = safeStr(r[M.Method_Code]), cc = safeStr(r[M.Carrier_Code]);
      if (!mc || !cc || !owner[mc] || owner[mc] === cc) continue;

      const before = r.slice();
      r[M.Carrier_Code] = owner[mc];
      // the code embeds the carrier, so it has to move too
      const code = safeStr(r[M.Modelling_Code]);
      if (code) r[M.Modelling_Code] = code.replace('__' + cc + '__', '__' + owner[mc] + '__');
      r[M.Notes] = (safeStr(r[M.Notes]) ? safeStr(r[M.Notes]) + '; ' : '') +
                   'carrier corrected from ' + cc + ' to ' + owner[mc];
      r[M.Updated_TS] = new Date();
      r[M.Updated_By] = email;

      Logger.log('  Modelling ID ' + safeInt(r[M.Modelling_ID]) + ': ' + cc +
                 ' -> ' + owner[mc] + '  (' + mc + ')');
      changes.push({ id: safeInt(before[M.Modelling_ID]), before: before,
                     after: r.slice(), type: 'UPDATE' });
    }
    if (changes.length) {
      sh.getRange(2, 1, data.length, width).setValues(data);
      invalidateSheetCache_(t.sheet);
      recordChangesBatch_('MODELLING_IDS', changes);
      changed += changes.length;
    } else {
      Logger.log('  none found');
    }
  }

  // ---- CC_COVERAGE -----------------------------------------------------
  Logger.log('');
  Logger.log('--- segments with no cold-chain row ---');
  const start  = monthStart(configDate('HORIZON_START', '2026-01-01'));
  const months = configInt('HORIZON_MONTHS', 36);
  const horizonFrom = start;
  const horizonTo   = monthEnd(new Date(start.getFullYear(), start.getMonth() + months - 1, 1));

  const covered = {};
  const cc = getAllData_(SHEET.MIX_COLDCHAIN), CC = COL.MIX_COLDCHAIN;
  for (let i = 1; i < cc.length; i++) {
    if (!safeBool(cc[i][CC.Active])) continue;
    covered[safeInt(cc[i][CC.High_Level_ID])] = true;
  }

  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  const ccTable = TABLES.MIX_COLDCHAIN, ccSheet = getSheet_(ccTable.sheet);
  let nextCc = getNextId_(ccTable.sheet, CC.CC_Mix_ID);
  const ccRows = [];
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id || !safeBool(hl[i][H.Active]) || covered[id]) continue;
    const row = blankRow_('MIX_COLDCHAIN');
    row[CC.CC_Mix_ID]      = nextCc++;
    row[CC.High_Level_ID]  = id;
    row[CC.Valid_From]     = horizonFrom;
    row[CC.Valid_To]       = horizonTo;
    row[CC.CC_Mix_Pct]     = 0;
    row[CC.Scenario_ID]    = 1;
    row[CC.Active]         = true;
    row[CC.Notes]          = 'Explicit 0% — this segment is entirely ambient';
    row[CC.Created_TS]     = new Date();
    row[CC.Created_By]     = email;
    row[CC.Updated_TS]     = new Date();
    row[CC.Updated_By]     = email;
    ccRows.push(row);
    Logger.log('  High Level ID ' + id + ': added an explicit 0% cold-chain row');
  }
  if (ccRows.length) {
    ccSheet.getRange(ccSheet.getLastRow() + 1, 1, ccRows.length,
                     ccTable.headers.length).setValues(ccRows);
    invalidateSheetCache_(ccTable.sheet);
    recordChangesBatch_('MIX_COLDCHAIN', ccRows.map(r =>
      ({ id: r[CC.CC_Mix_ID], before: null, after: r, type: 'CREATE' })));
    changed += ccRows.length;
  } else {
    Logger.log('  none found');
  }

  // ---- TBC_CARRIER: reported, not changed ------------------------------
  Logger.log('');
  Logger.log('--- routes still on the placeholder carrier TBC ---');
  const md2 = getAllData_(SHEET.MODELLING_IDS);
  const tbc = [];
  for (let i = 1; i < md2.length; i++) {
    if (safeStr(md2[i][M.Carrier_Code]) === 'TBC' && safeBool(md2[i][M.Active])) {
      tbc.push(safeInt(md2[i][M.Modelling_ID]) + ' (High Level ID ' +
               safeInt(md2[i][M.High_Level_ID]) + ', ' + safeStr(md2[i][M.Method_Code]) + ')');
    }
  }
  if (tbc.length) {
    tbc.forEach(x => Logger.log('  ' + x));
    Logger.log('');
    Logger.log('  Not changed. Which courier these should use is your decision — set it');
    Logger.log('  under Admin > High Level & Modelling IDs once you know.');
  } else {
    Logger.log('  none');
  }

  Logger.log('');
  Logger.log('  ' + changed + ' row(s) changed.');
  Logger.log('');
  Logger.log('  Next: run runValidation() to refresh the findings, then publishOutput().');
  logAudit_('UPDATE', 'FIX_WARNINGS', '', '', '', String(changed) + ' rows',
            'fixValidationWarnings', true);
  return { changed: changed, tbcRemaining: tbc.length };
}


/**
 * Remove everything the diagnostic functions wrote.
 *
 * The tests in Steps 5 to 7 deliberately wrote into 2035 and used ID 999999, so
 * they could never touch a forecast number. That kept them safe but left clutter,
 * and every one of those writes counts towards "changes unpublished".
 *
 * This deletes only rows matching those markers. Real data is identified by being
 * inside the forecast horizon, and is never touched.
 */
function cleanupTestData() {
  requireMaintenance_();
  Logger.log('=== REMOVING TEST DATA ===');

  const horizonEnd = new Date(2030, 0, 1);   // anything dated later is test data
  let removed = 0;

  function purge(tableKey, dateCol, idCol, testId) {
    const t = TABLES[tableKey], sh = getSheet_(t.sheet);
    const last = sh.getLastRow();
    if (last < 2) return 0;
    const data = sh.getRange(2, 1, last - 1, t.headers.length).getValues();
    const keep = [];
    let dropped = 0;

    for (let i = 0; i < data.length; i++) {
      const d = dateCol !== null ? normaliseDate(data[i][dateCol]) : null;
      const isFuture = d && d.getFullYear() >= horizonEnd.getFullYear();
      const isTestId = testId !== null && safeInt(data[i][idCol]) === testId;
      if (isFuture || isTestId) { dropped++; continue; }
      keep.push(data[i]);
    }
    if (!dropped) { Logger.log('  ' + t.sheet + ': nothing to remove'); return 0; }

    sh.getRange(2, 1, data.length, t.headers.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, t.headers.length).setValues(keep);
    invalidateSheetCache_(t.sheet);
    Logger.log('  ' + t.sheet + ': removed ' + dropped + ', kept ' + keep.length);
    return dropped;
  }

  removed += purge('MIX_METHOD',       COL.MIX_METHOD.Valid_From, COL.MIX_METHOD.Modelling_ID, null);
  removed += purge('RATE_BASE',        COL.RATE_BASE.Valid_From,  COL.RATE_BASE.Rate_ID, 999999);
  removed += purge('MIX_METHOD_AMENDS', COL.MIX_METHOD_AMENDS.Valid_From, COL.MIX_METHOD_AMENDS.Mix_ID, null);
  removed += purge('RATE_BASE_AMENDS',  COL.RATE_BASE_AMENDS.Valid_From,  COL.RATE_BASE_AMENDS.Rate_ID, 999999);

  Logger.log('');
  Logger.log('  ' + removed + ' test row(s) removed in total.');
  Logger.log('');
  Logger.log('  Audit_Log is left alone on purpose — it is the record of what happened,');
  Logger.log('  including the tests. Deleting history to make a counter look tidy is the');
  Logger.log('  wrong instinct.');
  Logger.log('');
  Logger.log('  Next: run publishOutput() so the forecast reflects the current data.');
  logAudit_('DELETE', 'TEST_DATA', '', '', '', String(removed) + ' rows',
            'cleanupTestData', true);
  return { removed: removed };
}


// ─────────────────────────────────────────────────────────────────────────────
// UTILITY — reset (destructive, asks for explicit confirmation)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Clears every data row from every tab but keeps headers. Use only while
 * rebuilding from scratch during setup. Change the constant to run it.
 */
function setupDANGERResetAllData() {
  requireMaintenance_();
  const CONFIRM = false;   // <- set to true to actually run
  if (!CONFIRM) {
    Logger.log('Refused: set CONFIRM = true inside setupDANGERResetAllData to proceed.');
    return;
  }
  const ss = _getSs_();
  for (const key in TABLES) {
    const sh = ss.getSheetByName(TABLES[key].sheet);
    if (sh && sh.getLastRow() > 1) {
      clearDataRows_(TABLES[key].sheet);
      Logger.log('Cleared ' + TABLES[key].sheet);
    }
  }
}