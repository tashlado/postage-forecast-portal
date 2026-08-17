/**
 * Postage Forecast Portal — utils.gs
 *
 * Single source of truth for sheet layout, plus the caching, coercion and
 * date helpers every other file depends on.
 *
 * Nothing in here reads user input or writes data. It is pure plumbing.
 */

// ─────────────────────────────────────────────────────────────────────────────
// 1. CONFIGURATION
//
// Which spreadsheet this project reads and writes is resolved at run time, not
// compiled in, so one code base can be pointed at a test copy without editing a
// source file you then have to remember not to commit:
//
//   Script Property `SPREADSHEET_ID`   used when it is set
//   SPREADSHEET_ID_FALLBACK            used when it is not
//
// The fallback is the literal this file has always carried, so a project with no
// property set behaves exactly as it did before the indirection existed.
//
// Never read SPREADSHEET_ID_FALLBACK anywhere but in spreadsheetId_() — that
// function is the only thing that knows whether a property is in play. Run
// showEnvironment() to see which of the two is live before running anything that
// writes.
// ─────────────────────────────────────────────────────────────────────────────

const SPREADSHEET_ID_FALLBACK = '1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE';

const ENGINE_VERSION   = '1.0.0';
const OPEN_ENDED_DATE  = '9999-12-31';   // sentinel meaning "no end date"
const DEFAULT_TZ       = 'Europe/London';
const PERF_LOG_ENABLED = true;

const _envIds_ = {};

/**
 * Resolve one spreadsheet ID from a Script Property, falling back to a literal.
 * Returns { id, from } where `from` is 'property' or 'fallback', memoised for
 * the execution so it costs one PropertiesService call however often it is asked.
 *
 * A read error on the property store is deliberately NOT treated as "no property
 * set". Swallowed, the two are indistinguishable, and guessing wrong here has the
 * worst consequence available in this codebase: a test project silently reading
 * and writing the production forecast. So it refuses rather than guesses.
 */
function envId_(propertyName, fallback) {
  if (_envIds_[propertyName]) return _envIds_[propertyName];

  let raw;
  try {
    raw = PropertiesService.getScriptProperties().getProperty(propertyName);
  } catch (e) {
    throw new Error('Could not read Script Properties, so the target spreadsheet cannot be ' +
                    'confirmed and this run is refusing to guess at it. Google said: ' + e.message);
  }

  const v = safeStr(raw);
  _envIds_[propertyName] = v ? { id: v, from: 'property' } : { id: fallback, from: 'fallback' };
  return _envIds_[propertyName];
}

/** The spreadsheet holding all app data. The only correct way to name it. */
function spreadsheetId_() {
  return envId_('SPREADSHEET_ID', SPREADSHEET_ID_FALLBACK).id;
}


// ─────────────────────────────────────────────────────────────────────────────
// 2. TABLE SCHEMA
//
// One definition drives everything: setup.gs creates tabs from it, and COL is
// derived from it. There is no second place where column order is written down,
// so the sheet and the code cannot drift apart.
// ─────────────────────────────────────────────────────────────────────────────

const AUDIT_COLS  = ['Created_TS', 'Created_By', 'Updated_TS', 'Updated_By'];
const AMEND_COLS  = ['Amend_ID', 'Amend_TS', 'Amend_By', 'Amend_Type'];

const TABLES = {

  // ---- Group A: dimensions & reference -------------------------------------
  DIM_REFERENCE: {
    sheet: 'Dim_Reference', width: 140,
    headers: ['Ref_ID', 'List_Name', 'Code', 'Label', 'Sort_Order', 'Active', 'Notes']
  },
  DIM_CARRIER: {
    sheet: 'Dim_Carrier', width: 150,
    headers: ['Carrier_Code', 'Carrier_Name', 'Default_Currency', 'Active', 'Notes']
  },
  DIM_METHOD: {
    sheet: 'Dim_Method', width: 170,
    headers: ['Method_Code', 'Carrier_Code', 'Method_Name', 'Service_Level', 'Is_Tracked', 'Active', 'Notes']
  },
  DIM_SURCHARGE: {
    sheet: 'Dim_Surcharge', width: 140,
    headers: ['Surcharge_Code', 'Surcharge_Name', 'Value_Type', 'Applies_To', 'Apply_Order', 'Proration', 'Active', 'Notes']
  },
  DIM_CALENDAR: {
    sheet: 'Dim_Calendar', width: 110,
    headers: ['Date_ID', 'Month_Start', 'Month_End', 'Days_In_Month', 'Year', 'Month_No', 'Month_Label', 'Quarter', 'In_Horizon']
  },

  // ---- Group B: structure ---------------------------------------------------
  HIGH_LEVEL_IDS: {
    sheet: 'High_Level_IDs', width: 130, amends: true,
    headers: ['High_Level_ID', 'High_Level_Code', 'Brand', 'Geo', 'Treatment_Type', 'WL_Split',
              'Currency', 'Active', 'Sort_Order', 'Notes'].concat(AUDIT_COLS)
  },
  MODELLING_IDS: {
    sheet: 'Modelling_IDs', width: 150, amends: true,
    headers: ['Modelling_ID', 'High_Level_ID', 'Carrier_Code', 'Method_Code', 'Letter_Parcel',
              'Modelling_Code', 'Active', 'Notes'].concat(AUDIT_COLS)
  },

  // ---- Group C: rates -------------------------------------------------------
  RATE_BASE: {
    sheet: 'Rate_Base', width: 120, amends: true,
    headers: ['Rate_ID', 'Modelling_ID', 'Valid_From', 'Valid_To', 'Base_Rate', 'Currency',
              'Scenario_ID', 'Source_Ref', 'Active', 'Notes'].concat(AUDIT_COLS)
  },
  RATE_SURCHARGE: {
    sheet: 'Rate_Surcharge', width: 120, amends: true,
    headers: ['Surcharge_Rate_ID', 'Modelling_ID', 'Surcharge_Code', 'Valid_From', 'Valid_To',
              'Value', 'Currency', 'Scenario_ID', 'Source_Ref', 'Active', 'Notes'].concat(AUDIT_COLS)
  },

  // ---- Group D: mixes -------------------------------------------------------
  MIX_METHOD: {
    sheet: 'Mix_Method', width: 120, amends: true,
    headers: ['Mix_ID', 'Modelling_ID', 'Temp_Regime', 'Valid_From', 'Valid_To', 'Mix_Pct',
              'Scenario_ID', 'Active', 'Notes'].concat(AUDIT_COLS)
  },
  MIX_LETTERPARCEL: {
    sheet: 'Mix_LetterParcel', width: 130, amends: true,
    headers: ['LP_Mix_ID', 'High_Level_ID', 'Valid_From', 'Valid_To', 'Letter_Mix_Pct',
              'Scenario_ID', 'Active', 'Notes'].concat(AUDIT_COLS)
  },
  MIX_COLDCHAIN: {
    sheet: 'Mix_ColdChain', width: 130, amends: true,
    headers: ['CC_Mix_ID', 'High_Level_ID', 'Valid_From', 'Valid_To', 'CC_Mix_Pct',
              'Scenario_ID', 'Active', 'Notes'].concat(AUDIT_COLS)
  },

  // ---- Group E: outputs -----------------------------------------------------
  OUTPUT: {
    sheet: 'OUTPUT', width: 130,
    headers: ['High_Level_ID', 'Date_ID', 'Month_Start', 'Brand', 'Geo', 'Treatment_Type',
              'WL_Split', 'Currency', 'Forecast_Rate_Per_Order', 'Scenario_ID', 'Calc_Run_ID']
  },
  OUTPUT_DETAIL: {
    sheet: 'OUTPUT_Detail', width: 120,
    headers: ['Modelling_ID', 'High_Level_ID', 'Date_ID', 'Month_Start', 'Brand', 'Geo',
              'Treatment_Type', 'WL_Split', 'Carrier_Code', 'Method_Code', 'Letter_Parcel',
              'Base_Rate', 'Surcharge_Pct_Total', 'Surcharge_Amt_Total', 'Rate_Per_Parcel',
              'CC_Share', 'Method_Mix', 'LP_Mix', 'Rate_Contribution', 'Rate_CTS',
              'Scenario_ID', 'Calc_Run_ID']
  },
  ACTUALS: {
    sheet: 'Actuals', width: 130, amends: true,
    headers: ['Actual_ID', 'High_Level_ID', 'Month_Start', 'Orders', 'Total_Spend',
              'Blended_Rate', 'Currency', 'Source', 'Source_Version', 'Active',
              'Notes'].concat(AUDIT_COLS)
  },
  ACTUALS_IMPORT: {
    sheet: 'Actuals_Import', width: 150,
    headers: ['Country Code', 'Brand', 'Delivery Carrier', 'Delivery Method',
              'Treatment Type', 'Dispatched Date: Month', 'Count',
              'Sum of Cost Of Shipping (£)']
  },
  SNAPSHOTS: {
    sheet: 'Snapshots', width: 140,
    headers: ['Snapshot_ID', 'Taken_TS', 'Taken_By', 'Trigger', 'Scenario_ID',
              'Label', 'Rows_Stored', 'Changed_Rows', 'Calc_Run_ID', 'Notes']
  },
  SNAPSHOT_VALUES: {
    sheet: 'Snapshot_Values', width: 120,
    headers: ['Snapshot_ID', 'High_Level_ID', 'Date_ID', 'Month_Start',
              'Forecast_Rate_Per_Order']
  },
  CALC_RUNS: {
    sheet: 'Calc_Runs', width: 140,
    headers: ['Calc_Run_ID', 'Run_TS', 'Run_By', 'Scenario_ID', 'Trigger', 'Rows_Output',
              'Rows_Output_Detail', 'Duration_Ms', 'Validation_Status', 'Validation_Summary',
              'Engine_Version']
  },

  // ---- Group F: governance --------------------------------------------------
  PERMISSIONS: {
    sheet: 'Permissions', width: 170,
    headers: ['Email', 'Display_Name', 'Role', 'Active', 'Tab_Visibility', 'Last_Login_TS', 'Notes']
  },
  PORTAL_ROLES: {
    sheet: 'Portal_Roles', width: 130,
    headers: ['Role', 'All_Access', 'Write_Access', 'Can_Edit_Rates', 'Can_Edit_Mixes',
              'Can_Edit_Structure', 'Can_Run_Calc', 'Can_Publish_Output', 'Can_Manage_Users',
              'Can_View_Audit']
  },
  SCOPE_MAPPING: {
    sheet: 'Scope_Mapping', width: 150,
    headers: ['Scope_ID', 'Email', 'Scope_Type', 'Scope_Value', 'Can_View', 'Can_Edit', 'Active']
  },
  AUDIT_LOG: {
    sheet: 'Audit_Log', width: 150,
    headers: ['Log_ID', 'TS', 'Email', 'Action', 'Entity', 'Entity_ID', 'Field',
              'Old_Value', 'New_Value', 'Detail', 'Success']
  },
  SCENARIOS: {
    sheet: 'Scenarios', width: 150,
    headers: ['Scenario_ID', 'Scenario_Name', 'Description', 'Parent_Scenario_ID', 'Is_Default',
              'Locked', 'Active', 'Created_TS', 'Created_By']
  },
  CONFIG: {
    sheet: 'Config', width: 200,
    headers: ['Key', 'Value', 'Description', 'Updated_TS', 'Updated_By']
  },
  FX_RATES: {
    sheet: 'FX_Rates', width: 130,
    headers: ['FX_ID', 'Month_Start', 'Currency', 'Rate_To_GBP', 'Source', 'Active']
  },
  VALIDATION_RESULTS: {
    sheet: 'Validation_Results', width: 130,
    headers: ['Result_ID', 'Calc_Run_ID', 'Rule_Code', 'Severity', 'Entity', 'Entity_ID',
              'High_Level_ID', 'Modelling_ID', 'Date_ID', 'Message', 'Resolved']
  }
};

// Amends tables are generated from any table flagged `amends: true`, so a column
// added to a source table automatically appears in its history table too.
(function buildAmendsTables_() {
  const keys = Object.keys(TABLES);
  for (let i = 0; i < keys.length; i++) {
    const t = TABLES[keys[i]];
    if (!t.amends) continue;
    TABLES[keys[i] + '_AMENDS'] = {
      sheet:   t.sheet + '_Amends',
      width:   t.width,
      headers: AMEND_COLS.concat(t.headers),
      isAmends: true
    };
  }
})();

// SHEET.RATE_BASE -> 'Rate_Base'
const SHEET = (function () {
  const o = {};
  for (const k in TABLES) o[k] = TABLES[k].sheet;
  return o;
})();

// COL.RATE_BASE.Base_Rate -> 4   (zero-based column index)
const COL = (function () {
  const o = {};
  for (const k in TABLES) {
    const m = {};
    TABLES[k].headers.forEach((h, i) => { m[h] = i; });
    o[k] = m;
  }
  return o;
})();

// Reverse lookup: sheet name -> table key
const TABLE_BY_SHEET = (function () {
  const o = {};
  for (const k in TABLES) o[TABLES[k].sheet] = k;
  return o;
})();

/** Number format for a column, chosen by naming convention. */
function formatForColumn_(header) {
  // Plain text, or Sheets helpfully parses "Jan-26" into 26 January and the
  // label silently becomes a date. Learned the hard way.
  if (/^(Month_Label|Quarter|High_Level_Code|Modelling_Code|Code)$/.test(header)) return '@';
  if (/_TS$/.test(header))                              return 'yyyy-mm-dd hh:mm:ss';
  if (/^(Valid_From|Valid_To|Month_Start|Month_End)$/.test(header)) return 'yyyy-mm-dd';
  if (/_Pct$|^Method_Mix$|^LP_Mix$|^CC_Share$|^Surcharge_Pct_Total$/.test(header)) return '0.00%';
  if (/^(Base_Rate|Value|Rate_Per_Parcel|Rate_Contribution|Rate_CTS|Surcharge_Amt_Total|Forecast_Rate_Per_Order|Rate_To_GBP|Blended_Rate|Total_Spend)$/.test(header)) return '#,##0.0000';
  if (/_ID$|^Sort_Order$|^Apply_Order$|^Days_In_Month$|^Year$|^Month_No$|^Duration_Ms$|^Rows_/.test(header)) return '0';
  return null;
}


// ─────────────────────────────────────────────────────────────────────────────
// 3. CACHING
//
// Module-level state persists for one execution and resets on the next, which
// is exactly the lifetime we want. openById() and getValues() are ~300-500ms
// each, so a single request can otherwise lose several seconds to repeats.
// ─────────────────────────────────────────────────────────────────────────────

let _ssCache_ = null;
const _sheetCache_     = {};
const _sheetDataCache_ = {};
const _nextIdCache_    = {};

function _getSs_() {
  if (!_ssCache_) _ssCache_ = SpreadsheetApp.openById(spreadsheetId_());
  return _ssCache_;
}

function getSheet_(name) {
  if (_sheetCache_[name]) return _sheetCache_[name];
  const sh = _getSs_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setupCreateAllTabs() first.');
  _sheetCache_[name] = sh;
  return sh;
}

function getAllData_(sheetName) {
  if (_sheetDataCache_[sheetName]) return _sheetDataCache_[sheetName];
  const data = getSheet_(sheetName).getDataRange().getValues();
  _sheetDataCache_[sheetName] = (data && data.length) ? data : [[]];
  return _sheetDataCache_[sheetName];
}

function invalidateSheetCache_(sheetName) {
  delete _sheetDataCache_[sheetName];
}

/**
 * Fetch many tabs in ONE HTTP call.
 *
 * Tries three paths in order, and each falls through cleanly to the next:
 *
 *   1. Advanced Sheets Service  — fastest, needs Services > + > Google Sheets API
 *   2. Sheets REST API          — same speed, needs only the script's OAuth token
 *   3. Per-sheet reads          — automatic, correct, just slower
 *
 * Path 2 exists because the Advanced Sheets Service is not offered in every
 * Workspace: an admin API allowlist can remove it from the Services list. The
 * REST call reaches the identical endpoint using the token the script already
 * holds, so it needs no project-level setup.
 *
 * Run diagnoseBatchGet() to see which path your project is using.
 */
function prewarmSheetCache_(sheetNames) {
  const toFetch = sheetNames.filter(n => !_sheetDataCache_[n]);
  if (!toFetch.length) return 'cached';
  const ranges = toFetch.map(n => "'" + String(n).replace(/'/g, "''") + "'");

  // ---- Path 1: Advanced Sheets Service ------------------------------------
  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
    try {
      const resp = Sheets.Spreadsheets.Values.batchGet(spreadsheetId_(), {
        ranges: ranges,
        valueRenderOption:    'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'SERIAL_NUMBER'
      });
      _storeBatch_(toFetch, resp.valueRanges);
      return 'advanced';
    } catch (err) {
      Logger.log('Advanced Sheets Service failed, trying REST: ' + err.message);
    }
  }

  // ---- Path 2: Sheets REST API via the script's own OAuth token ----------
  try {
    const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId_() +
                '/values:batchGet?' +
                ranges.map(r => 'ranges=' + encodeURIComponent(r)).join('&') +
                '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER';
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      _storeBatch_(toFetch, JSON.parse(resp.getContentText()).valueRanges);
      return 'rest';
    }
    Logger.log('REST batchGet returned HTTP ' + resp.getResponseCode() + ': ' +
               resp.getContentText().slice(0, 300));
  } catch (err) {
    Logger.log('REST batchGet failed: ' + err.message);
  }

  // ---- Path 3: per-sheet reads (getAllData_ does this on demand) ----------
  Logger.log('Using per-sheet reads for ' + toFetch.length + ' tabs. Correct but slower — ' +
             'run diagnoseBatchGet() to see why the batch paths were unavailable.');
  return 'fallback';
}

/** Load a batchGet response into the per-execution data cache. */
function _storeBatch_(sheetNames, valueRanges) {
  const vrs = valueRanges || [];
  for (let i = 0; i < sheetNames.length; i++) {
    const v = (vrs[i] && vrs[i].values) || [];
    _sheetDataCache_[sheetNames[i]] = v.length ? v : [[]];
  }
}

/**
 * Reports which bulk-read path works, and how much faster it is than reading
 * tab by tab. Run this from the editor after setup; it only reads.
 */
function diagnoseBatchGet() {
  requireMaintenance_();
  const probe = [SHEET.DIM_REFERENCE, SHEET.DIM_CARRIER, SHEET.DIM_METHOD,
                 SHEET.DIM_SURCHARGE, SHEET.DIM_CALENDAR, SHEET.PORTAL_ROLES,
                 SHEET.CONFIG, SHEET.SCENARIOS];

  Logger.log('--- bulk read diagnostic ---');
  Logger.log('Advanced Sheets Service present: ' +
             (typeof Sheets !== 'undefined' && !!(Sheets.Spreadsheets && Sheets.Spreadsheets.Values)));

  for (const k in _sheetDataCache_) delete _sheetDataCache_[k];
  let t0 = Date.now();
  const path = prewarmSheetCache_(probe);
  const batchMs = Date.now() - t0;
  Logger.log('Path used: ' + path + '  (' + batchMs + 'ms for ' + probe.length + ' tabs)');

  for (const k in _sheetDataCache_) delete _sheetDataCache_[k];
  t0 = Date.now();
  probe.forEach(n => getAllData_(n));
  const serialMs = Date.now() - t0;
  Logger.log('Per-sheet reads: ' + serialMs + 'ms for ' + probe.length + ' tabs');

  if (path === 'fallback') {
    Logger.log('');
    Logger.log('Neither batch path is available. The app will still work correctly.');
    Logger.log('To enable one, either add Services > + > Google Sheets API, or ask your');
    Logger.log('Workspace admin about API access restrictions (Admin console >');
    Logger.log('Security > API controls). See the notes in STEP_1_GUIDE.md section 1.5.');
  } else {
    Logger.log('');
    Logger.log('Speed-up: ' + (serialMs / Math.max(batchMs, 1)).toFixed(1) + 'x');
  }
  return { path: path, batchMs: batchMs, serialMs: serialMs };
}

/** Every write needs these. Call at the very top of a write, before auth. */
function prewarmForWrite_(extraSheets) {
  const sheets = [SHEET.PERMISSIONS, SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING,
                  SHEET.AUDIT_LOG, SHEET.CONFIG];
  (extraSheets || []).forEach(s => { if (sheets.indexOf(s) < 0) sheets.push(s); });
  prewarmSheetCache_(sheets);
  _getSs_();   // batchGet uses a different path, so prime the handle too
}


// ─────────────────────────────────────────────────────────────────────────────
// 4. COERCION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Is this a Date?
 *
 * Duck-typed rather than `instanceof Date`, which compares against one specific
 * realm's constructor and silently returns false for a Date created elsewhere.
 * Checking for a working getTime() is both safer and no slower.
 */
function isDate_(v) {
  return !!v && typeof v.getTime === 'function' && !isNaN(v.getTime());
}

function safeStr(v)  { return (v === null || v === undefined) ? '' : String(v).trim(); }
function safeInt(v)  { const n = parseInt(v, 10);   return isNaN(n) ? 0 : n; }
function safeNum(v)  { const n = parseFloat(v);     return isNaN(n) ? 0 : n; }
function safeBool(v) {
  if (v === true)  return true;
  if (v === false) return false;
  const s = safeStr(v).toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === 'Y' || s === '1';
}
/**
 * Right-align a value to a fixed width, for readable log tables.
 *
 * Lived in enginetest.gs until actuals.gs, output.gs and validate.gs all started
 * using it — a shared helper in a test file works only until someone deletes the
 * test file.
 */
function pad_(v, n) {
  let s = String(v);
  while (s.length < n) s = ' ' + s;
  return s;
}

/** Case-insensitive comparison key. Fixes the WeGovy / Wegovy problem. */
function normKey(v) { return safeStr(v).toUpperCase(); }


// ─────────────────────────────────────────────────────────────────────────────
// 5. DATES
//
// THE SERIAL-DATE TRAP: batchGet returns dates as serial numbers (days since
// 1899-12-30), not Date objects. new Date(47391) is 1970, not 2029. Since this
// whole model is date-range arithmetic, every date must come through here.
// ─────────────────────────────────────────────────────────────────────────────

const SERIAL_EPOCH_OFFSET = 25569;   // days from 1899-12-30 to 1970-01-01
const MS_PER_DAY = 86400000;

function normaliseDate(val) {
  if (val === null || val === undefined || val === '') return null;
  let d;
  if (isDate_(val)) {
    d = new Date(val.getTime());
  } else if (typeof val === 'number') {
    // Bounds-check so a genuine number (a rate, a percentage) is not read as a date.
    if (val < SERIAL_EPOCH_OFFSET || val > 2958466) return null;   // ~1970 to 9999
    d = new Date(Math.round((val - SERIAL_EPOCH_OFFSET) * MS_PER_DAY));
  } else {
    const s = safeStr(val);
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s);
  }
  if (!d || isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

/** Comparable integer key for a date: 20260401. Avoids timezone drift entirely. */
function dateKey(val) {
  const d = normaliseDate(val);
  if (!d) return 0;
  return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

function monthStart(d)   { const x = normaliseDate(d); return x ? new Date(x.getFullYear(), x.getMonth(), 1) : null; }
function monthEnd(d)     { const x = normaliseDate(d); return x ? new Date(x.getFullYear(), x.getMonth() + 1, 0) : null; }
function daysInMonth(d)  { const x = monthEnd(d); return x ? x.getDate() : 0; }
function addMonths(d, n) { const x = normaliseDate(d); return x ? new Date(x.getFullYear(), x.getMonth() + n, 1) : null; }
function daysBetween(a, b) {
  const x = normaliseDate(a), y = normaliseDate(b);
  if (!x || !y) return 0;
  return Math.round((y.getTime() - x.getTime()) / MS_PER_DAY);
}
const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function pad2_(n) { return n < 10 ? '0' + n : String(n); }

/**
 * Format a date as yyyy-MM-dd, in plain JavaScript.
 *
 * Utilities.formatDate is a SERVICE call — it crosses into Google's runtime and
 * costs roughly 0.5ms. Loading the portal formats about 5,000 dates, which came
 * to 2.6 seconds of the 3 second load. Apps Script already sets the runtime's
 * local timezone to the project timezone, so reading the components directly
 * gives an identical result for nothing.
 */
function fmtDate(d) {
  const x = normaliseDate(d);
  if (!x) return '';
  return x.getFullYear() + '-' + pad2_(x.getMonth() + 1) + '-' + pad2_(x.getDate());
}

function fmtTs(val) {
  if (val === null || val === undefined || val === '') return '';
  let d = isDate_(val) ? val
        : (typeof val === 'number') ? new Date(Math.round((val - SERIAL_EPOCH_OFFSET) * MS_PER_DAY))
        : new Date(safeStr(val));
  if (!d || isNaN(d.getTime())) return '';
  return pad2_(d.getDate()) + ' ' + MONTH_ABBR[d.getMonth()] + ' ' + d.getFullYear() +
         ', ' + pad2_(d.getHours()) + ':' + pad2_(d.getMinutes());
}

/** MMM-yy, e.g. Jan-26. Also plain JavaScript, for the same reason. */
function fmtMonthLabel(d) {
  const x = normaliseDate(d);
  if (!x) return '';
  return MONTH_ABBR[x.getMonth()] + '-' + String(x.getFullYear()).slice(2);
}
function isOpenEnded(d) { return dateKey(d) >= dateKey(OPEN_ENDED_DATE); }


// ─────────────────────────────────────────────────────────────────────────────
// 6. IDS AND LOCKING
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Next surrogate key. Reads only the ID column when the sheet is not cached,
 * and increments in memory for repeat calls within one execution — safe
 * because withLock_() serialises writes.
 */
function getNextId_(sheetName, colIndex) {
  if (_nextIdCache_[sheetName] !== undefined) return ++_nextIdCache_[sheetName];

  let max = 0;
  if (_sheetDataCache_[sheetName]) {
    const data = _sheetDataCache_[sheetName];
    for (let i = 1; i < data.length; i++) {
      const v = parseInt(data[i][colIndex], 10);
      if (!isNaN(v) && v > max) max = v;
    }
  } else {
    const sh = getSheet_(sheetName);
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const ids = sh.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        const v = parseInt(ids[i][0], 10);
        if (!isNaN(v) && v > max) max = v;
      }
    }
  }
  _nextIdCache_[sheetName] = max + 1;
  return _nextIdCache_[sheetName];
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('The system is busy with another change. Please try again.');
  }
  try   { return fn(); }
  finally { lock.releaseLock(); }
}

/**
 * Empty every data row of a tab, keeping the header.
 *
 * deleteRows(2, lastRow - 1) looks like the obvious way to do this and works
 * exactly once: the moment the sheet contains nothing but the header and data,
 * that call is asking Sheets to delete every non-frozen row, which it refuses
 * with "Sorry, it is not possible to delete all non-frozen rows."
 *
 * Clearing the contents instead has no such edge case, and is faster — no
 * structural change to the sheet, so no reflow.
 */
function clearDataRows_(sheetName) {
  const sh = getSheet_(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent();
  invalidateSheetCache_(sheetName);
  return last - 1;
}

/** Build an empty row sized to a table, so column count always matches. */
function blankRow_(tableKey) {
  return new Array(TABLES[tableKey].headers.length).fill('');
}


// ─────────────────────────────────────────────────────────────────────────────
// 7. CONFIG ACCESS
// ─────────────────────────────────────────────────────────────────────────────

let _configCache_ = null;
function getConfig_() {
  if (_configCache_) return _configCache_;
  const data = getAllData_(SHEET.CONFIG);
  const C = COL.CONFIG, o = {};
  for (let i = 1; i < data.length; i++) {
    const k = safeStr(data[i][C.Key]);
    if (k) o[k] = safeStr(data[i][C.Value]);
  }
  _configCache_ = o;
  return o;
}
function configStr(key, dflt)  { const v = getConfig_()[key]; return (v === undefined || v === '') ? dflt : v; }

/**
 * A config value that is meant to be a date, as yyyy-MM-dd.
 *
 * The Config tab stores everything as text, but Sheets parses anything
 * date-shaped the moment it is written into an unformatted cell — so
 * HORIZON_START came back as the serial number 46023 rather than 2026-01-01,
 * and leaked into the interface as a prefilled value nobody could read.
 * Same trap as the calendar labels. Normalising on read means it cannot matter
 * how the cell ended up.
 */
function configDate(key, dflt) {
  const raw = getConfig_()[key];
  const d = normaliseDate((raw === undefined || raw === '') ? dflt : raw);
  return d ? fmtDate(d) : safeStr(dflt);
}
function configInt(key, dflt)  { const v = getConfig_()[key]; return (v === undefined || v === '') ? dflt : safeInt(v); }
function configBool(key, dflt) { const v = getConfig_()[key]; return (v === undefined || v === '') ? dflt : safeBool(v); }


// ─────────────────────────────────────────────────────────────────────────────
// 8. PERF INSTRUMENTATION
// ─────────────────────────────────────────────────────────────────────────────

let _perfTimings = [];
let _perfStart   = 0;

function perfReset() { if (!PERF_LOG_ENABLED) return; _perfTimings = []; _perfStart = Date.now(); }
function perfMark(label) {
  if (!PERF_LOG_ENABLED) return;
  _perfTimings.push({ t: Date.now() - _perfStart, label: label });
}
function perfWrap(label, fn) {
  if (!PERF_LOG_ENABLED) return fn();
  const t0 = Date.now();
  const r  = fn();
  _perfTimings.push({ t: Date.now() - _perfStart, label: label + ' [' + (Date.now() - t0) + 'ms]' });
  return r;
}
function perfReport() {
  if (!PERF_LOG_ENABLED) return null;
  const total = Date.now() - _perfStart;
  Logger.log('--- PERF total ' + total + 'ms ---');
  _perfTimings.forEach(t => Logger.log('  ' + t.t + 'ms  ' + t.label));
  return { totalMs: total, marks: _perfTimings.slice() };
}


/**
 * Maintenance functions are meant to be run by an administrator from the editor,
 * but google.script.run can reach any global that does not end in an underscore.
 * These are named without one so they appear in the editor's function list, so
 * they carry their own check instead.
 *
 * Allowed while Permissions is still empty, otherwise ADMIN only — Step 1 has to
 * work before anyone has been given a role.
 */
function requireMaintenance_() {
  try {
    const data = getAllData_(SHEET.PERMISSIONS);
    if (!data || data.length <= 1) return true;          // still bootstrapping
  } catch (e) {
    return true;                                          // tab does not exist yet
  }
  const perms = getUserPermissions_();
  if (perms.found && perms.active && (perms.allAccess || perms.capabilities.manageUsers)) return true;
  logAudit_('DENIED', 'MAINTENANCE', '', '', '', '', 'not an administrator', false);
  throw new Error('That is an administrator function.');
}


/**
 * Which spreadsheet is this project about to read and write?
 *
 * Run this before anything that writes, and first of all after moving between
 * projects. Resolving the ID at run time is what makes a test copy possible, but
 * it also means the answer is no longer visible by reading the source — so there
 * has to be one function whose whole job is to tell you. Read only.
 */
function showEnvironment() {
  requireMaintenance_();
  Logger.log('=== ENVIRONMENT ===');

  let scriptId = '';
  try { scriptId = ScriptApp.getScriptId(); } catch (e) { scriptId = '(unavailable)'; }
  Logger.log('  script project  : ' + scriptId);
  Logger.log('');

  const target = envId_('SPREADSHEET_ID', SPREADSHEET_ID_FALLBACK);
  Logger.log('--- app data: everything the portal reads and writes ---');
  Logger.log('  spreadsheet ID  : ' + target.id);
  Logger.log('  resolved from   : ' + (target.from === 'property'
    ? 'the SPREADSHEET_ID Script Property'
    : 'the committed fallback in utils.gs — no SPREADSHEET_ID property is set'));

  let name = '';
  try {
    name = SpreadsheetApp.openById(target.id).getName();
    Logger.log('  file name       : ' + name);
    Logger.log('  url             : https://docs.google.com/spreadsheets/d/' + target.id + '/edit');
  } catch (e) {
    Logger.log('  CANNOT OPEN IT  : ' + e.message);
    Logger.log('  Check the ID above is a Google Sheet this account can open.');
  }

  // Reported only while migrate.gs is present — it is a candidate for removal
  // (FINDINGS.md M4), so this must not become the thing that keeps it alive.
  let source = null;
  if (typeof sourceEnvironment_ === 'function') {
    source = sourceEnvironment_();
    Logger.log('');
    Logger.log('--- legacy workbook: read only, migration and parity tests ---');
    Logger.log('  spreadsheet ID  : ' + source.id);
    Logger.log('  resolved from   : ' + (source.from === 'property'
      ? 'the SOURCE_SPREADSHEET_ID Script Property'
      : 'the committed fallback in migrate.gs'));
  }

  Logger.log('');
  Logger.log('Every write from this project lands in the "app data" spreadsheet above.');
  Logger.log('To point this project elsewhere: Project Settings > Script Properties > add');
  Logger.log('SPREADSHEET_ID. Delete the property to return to the committed fallback.');

  return {
    scriptId: scriptId,
    target: { id: target.id, from: target.from, name: name },
    source: source ? { id: source.id, from: source.from } : null
  };
}