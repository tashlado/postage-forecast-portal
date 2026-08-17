/**
 * Postage Forecast Portal — migrate.gs
 *
 * Copies the forecast data out of the original workbook and into the normalised
 * tables built in Step 1.
 *
 * Run in this order:
 *   1. migratePreflight()   READ ONLY. Checks the source, reports problems, writes nothing.
 *   2. migrateAll()         Does the migration.
 *   3. migrateVerify()      Checks what landed.
 *
 * Safe to re-run: migrateAll refuses to start if the target tables already hold
 * data, so it can never double-load. Use migrateDANGERClearMigratedData() to reset.
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURATION — the Google Sheets copy of Postage_RFQ3_New.xlsx
//
// Resolved at run time the same way as the app's own spreadsheet (see utils.gs
// §1): the Script Property `SOURCE_SPREADSHEET_ID` wins if set, otherwise the
// literal below. Read it through sourceSpreadsheetId_(), never directly.
// ─────────────────────────────────────────────────────────────────────────────

const SOURCE_SPREADSHEET_ID_FALLBACK = '18a-Pfaa3hy0OlNcl3jYri-qccLku13JyDhfm0lOq4io';

/** The legacy workbook, and whether it came from a property or the fallback. */
function sourceEnvironment_() {
  return envId_('SOURCE_SPREADSHEET_ID', SOURCE_SPREADSHEET_ID_FALLBACK);
}

/** The legacy workbook's ID. The only correct way to name it. */
function sourceSpreadsheetId_() {
  return sourceEnvironment_().id;
}

/** Tab names in the SOURCE workbook. Change only if yours are named differently. */
const SRC = {
  HIGH_LEVEL: 'High Level IDs',
  MODELLING:  'Modelling IDs',
  BASE:       'Base Rate Card',
  FUEL:       'Fuel Surcharge Rate Card',
  OTHER:      'Other Surcharge Rate Card',
  METHOD_MIX: 'Method Mix Inputs',
  LP_MIX:     'LetterParcel Mix',
  COLD_CHAIN: 'Cold Chain Inputs',
  FX:         'FX Rates'
};

/** Expected row counts, used by preflight to catch a truncated conversion. */
const SRC_EXPECTED = {};
SRC_EXPECTED[SRC.HIGH_LEVEL] = 17;
SRC_EXPECTED[SRC.MODELLING]  = 232;
SRC_EXPECTED[SRC.BASE]       = 264;
SRC_EXPECTED[SRC.FUEL]       = 380;
SRC_EXPECTED[SRC.OTHER]      = 1038;
SRC_EXPECTED[SRC.METHOD_MIX] = 330;
SRC_EXPECTED[SRC.LP_MIX]     = 17;
SRC_EXPECTED[SRC.COLD_CHAIN] = 139;

const MIGRATION_SCENARIO_ID = 1;   // everything lands in the BASE scenario


// ─────────────────────────────────────────────────────────────────────────────
// SOURCE READING
// ─────────────────────────────────────────────────────────────────────────────

const _srcCache_ = {};
let _srcTabNames_ = null;
const _srcTabResolved_ = {};

/** Every tab name in the source workbook, as actually spelled. */
function sourceTabNames_() {
  if (_srcTabNames_) return _srcTabNames_;
  _srcTabNames_ = SpreadsheetApp.openById(sourceSpreadsheetId_())
                    .getSheets().map(s => s.getName());
  return _srcTabNames_;
}

/** Strip everything but letters and digits, so spacing and case stop mattering. */
function tabKey_(name) {
  return safeStr(name).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Find a tab even if it's spelled slightly differently in the Google copy —
 * "Letter Parcel Mix" vs "LetterParcel Mix", stray spaces, different casing.
 * Exact match wins; otherwise fall back to the normalised comparison.
 */
function resolveSourceTab_(wanted) {
  if (_srcTabResolved_[wanted]) return _srcTabResolved_[wanted];
  const names = sourceTabNames_();

  if (names.indexOf(wanted) >= 0) { _srcTabResolved_[wanted] = wanted; return wanted; }

  const want = tabKey_(wanted);
  const hit = names.filter(n => tabKey_(n) === want);
  if (hit.length === 1) {
    Logger.log('  note: tab "' + wanted + '" matched to "' + hit[0] + '" in the source');
    _srcTabResolved_[wanted] = hit[0];
    return hit[0];
  }
  if (hit.length > 1) {
    throw new Error('Ambiguous tab name "' + wanted + '" — the source has several that ' +
                    'normalise the same way: ' + hit.join(', '));
  }
  throw new Error('SOURCE_TAB_NOT_FOUND:' + wanted);
}

/** Read one tab from the SOURCE workbook. Values only, dates as serial numbers. */
function readSourceTab_(tabName) {
  const actual = resolveSourceTab_(tabName);
  if (_srcCache_[actual]) return _srcCache_[actual];
  let rows = null;

  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
    try {
      const resp = Sheets.Spreadsheets.Values.get(
        sourceSpreadsheetId_(), "'" + actual.replace(/'/g, "''") + "'",
        { valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER' });
      rows = resp.values || [];
    } catch (err) {
      throw new Error('Could not read "' + actual + '" from the source workbook. ' +
                      'Google said: ' + err.message);
    }
  } else {
    const sh = SpreadsheetApp.openById(sourceSpreadsheetId_()).getSheetByName(actual);
    if (!sh) throw new Error('Source tab "' + actual + '" not found.');
    rows = sh.getDataRange().getValues();
  }

  _srcCache_[actual] = rows;
  return rows;
}

/**
 * Diagnostic: lists every tab in the source workbook and shows which one the
 * migration will use for each table it needs. Run this if preflight reports a
 * missing tab. Read only.
 */
function listSourceTabs() {
  requireMaintenance_();
  Logger.log('=== SOURCE WORKBOOK TABS ===');
  Logger.log('Workbook: ' + SpreadsheetApp.openById(sourceSpreadsheetId_()).getName());
  Logger.log('');
  const names = sourceTabNames_();
  Logger.log('--- all ' + names.length + ' tabs, exactly as spelled ---');
  names.forEach(n => Logger.log('  "' + n + '"   (' + n.length + ' characters)'));

  Logger.log('');
  Logger.log('--- what the migration needs ---');
  const missing = [];
  for (const key in SRC) {
    const wanted = SRC[key];
    try {
      const actual = resolveSourceTab_(wanted);
      Logger.log('  OK       ' + wanted + (actual === wanted ? '' : '   -> using "' + actual + '"'));
    } catch (e) {
      Logger.log('  MISSING  ' + wanted);
      const want = tabKey_(wanted);
      const close = names.filter(n => tabKey_(n).indexOf(want.slice(0, 8)) >= 0 ||
                                      want.indexOf(tabKey_(n).slice(0, 8)) >= 0);
      if (close.length) Logger.log('           closest matches: ' + close.map(n => '"' + n + '"').join(', '));
      missing.push(wanted);
    }
  }

  Logger.log('');
  if (missing.length) {
    Logger.log(missing.length + ' tab(s) could not be found.');
    Logger.log('Either rename the tab in the source workbook to match, or edit the SRC');
    Logger.log('list at the top of migrate.gs to match what the workbook actually calls it.');
  } else {
    Logger.log('All tabs found. Run migratePreflight() next.');
  }
  return { tabs: names, missing: missing };
}

/** Rows with a non-blank first column, excluding the header. */
function sourceDataRows_(tabName, idCol) {
  const all = readSourceTab_(tabName);
  const out = [];
  for (let i = 1; i < all.length; i++) {
    const r = all[i] || [];
    if (safeStr(r[idCol === undefined ? 0 : idCol]) === '') continue;
    out.push(r);
  }
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// CODE NORMALISATION
// ─────────────────────────────────────────────────────────────────────────────

/** 'Core Rx' -> 'CORE_RX', 'MedExpress' -> 'MEDEXPRESS', '*' stays '*'. */
function toCode_(v) {
  const s = safeStr(v);
  if (s === '' || s === '*') return s;
  return s.toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function letterParcelCode_(v) {
  const s = safeStr(v).toUpperCase();
  if (s.indexOf('LETTER') === 0) return 'LETTER';
  if (s.indexOf('PARCEL') === 0) return 'PARCEL';
  return s;
}

function tempRegimeCode_(v) {
  const s = safeStr(v).toUpperCase();
  if (s === 'CC' || s === 'COLD CHAIN' || s === 'COLDCHAIN') return 'CC';
  if (s.indexOf('AMBIENT') === 0) return 'AMBIENT';
  return s;
}


// ─────────────────────────────────────────────────────────────────────────────
// 1. PREFLIGHT — read only, writes nothing
// ─────────────────────────────────────────────────────────────────────────────

function migratePreflight() {
  requireMaintenance_();
  Logger.log('=== MIGRATION PREFLIGHT (read only — nothing will be written) ===');
  const problems = [], warnings = [];

  // -- source reachable? ----------------------------------------------------
  try {
    SpreadsheetApp.openById(sourceSpreadsheetId_()).getName();
  } catch (e) {
    Logger.log('CANNOT OPEN SOURCE WORKBOOK.');
    Logger.log('  ID used: ' + sourceSpreadsheetId_());
    Logger.log('  ' + e.message);
    Logger.log('Check the ID, and that the file is a Google Sheet (not still an .xlsx).');
    return { ok: false, problems: ['source unreadable'] };
  }
  Logger.log('Source workbook: ' + SpreadsheetApp.openById(sourceSpreadsheetId_()).getName());
  Logger.log('');

  // -- tabs present and the right size? -------------------------------------
  Logger.log('--- source tabs ---');
  for (const key in SRC) {
    const tab = SRC[key];
    let rows;
    try { rows = sourceDataRows_(tab); }
    catch (e) {
      const why = (String(e.message).indexOf('SOURCE_TAB_NOT_FOUND') === 0)
        ? 'Missing tab: ' + tab + '  (run listSourceTabs() to see the real names)'
        : 'Could not read tab ' + tab + ': ' + e.message;
      problems.push(why); Logger.log('  MISSING  ' + tab); continue;
    }

    const exp = SRC_EXPECTED[tab];
    if (exp === undefined) { Logger.log('  OK       ' + tab + ': ' + rows.length + ' rows'); continue; }
    if (rows.length === exp) Logger.log('  OK       ' + tab + ': ' + rows.length + ' rows');
    else {
      Logger.log('  CHECK    ' + tab + ': ' + rows.length + ' rows (expected ' + exp + ')');
      warnings.push(tab + ' has ' + rows.length + ' rows, expected ' + exp);
    }
  }

  // -- did the conversion actually compute the formulas? --------------------
  Logger.log('');
  Logger.log('--- conversion sanity ---');
  const mrows = sourceDataRows_(SRC.MODELLING);
  if (mrows.length) {
    const r = mrows[0];
    const ok = safeStr(r[2]) !== '' && safeStr(r[6]) !== '' && safeStr(r[8]) !== '';
    Logger.log('  Modelling IDs row 1: HLID=' + safeStr(r[1]) + ' brand=' + safeStr(r[2]) +
               ' carrier=' + safeStr(r[6]) + ' method=' + safeStr(r[7]) + ' L/P=' + safeStr(r[8]));
    if (!ok) problems.push('Modelling IDs lookup columns are blank — the xlsx->Sheets ' +
                           'conversion did not evaluate the XLOOKUP formulas.');
  }
  const brows = sourceDataRows_(SRC.BASE);
  if (brows.length) {
    const r = brows[0];
    const from = normaliseDate(r[8]), to = normaliseDate(r[9]);
    Logger.log('  Base Rate Card row 1: from=' + fmtDate(from) + ' to=' + fmtDate(to) +
               ' rate=' + safeNum(r[10]));
    if (!from || !to) problems.push('Base Rate Card dates did not read as dates.');
    if (safeNum(r[10]) <= 0) warnings.push('Base Rate Card row 1 has a zero rate.');
  }

  // -- the two known data issues -------------------------------------------
  Logger.log('');
  Logger.log('--- known data issues (see spec 3.1 and 3.2) ---');
  const scan = scanSurchargeIssues_();
  Logger.log('  RANGE_OVERLAP      : ' + scan.overlaps.length + ' found' +
             (scan.overlaps.length ? '  (Peak periods starting before the previous one ends)' : '   OK'));
  scan.overlaps.slice(0, 5).forEach(o => Logger.log('      Modelling ID ' + o.modellingId + ' ' + o.code +
      ': ' + fmtDate(o.from) + ' starts before ' + fmtDate(o.prevTo) + ' ends'));
  if (scan.overlaps.length > 5) Logger.log('      ...and ' + (scan.overlaps.length - 5) + ' more');

  Logger.log('  TERMINAL_MID_MONTH : ' + scan.truncated.length + ' found' +
             (scan.truncated.length ? '  (final period ends mid-month, understating that month)' : '   OK'));
  scan.truncated.slice(0, 5).forEach(o => Logger.log('      Modelling ID ' + o.modellingId + ' ' + o.code +
      ': ends ' + fmtDate(o.to) + ', should be ' + fmtDate(monthEnd(o.to))));
  if (scan.truncated.length > 5) Logger.log('      ...and ' + (scan.truncated.length - 5) + ' more');

  if (!scan.overlaps.length && !scan.truncated.length) {
    Logger.log('  Both already fixed at source. Parity target will be 612 of 612.');
  } else {
    Logger.log('  migrateAll() will correct these as it loads, and log every change.');
    Logger.log('  Parity target will be 611 of 612 (High Level ID 2, Dec-2028 moves 3.0966 -> 3.1861).');
  }

  // -- is the target empty? -------------------------------------------------
  Logger.log('');
  Logger.log('--- target tables ---');
  const targets = ['HIGH_LEVEL_IDS','MODELLING_IDS','RATE_BASE','RATE_SURCHARGE',
                   'MIX_METHOD','MIX_LETTERPARCEL','MIX_COLDCHAIN','FX_RATES'];
  let notEmpty = 0;
  targets.forEach(k => {
    const n = Math.max(getSheet_(SHEET[k]).getLastRow() - 1, 0);
    if (n > 0) { notEmpty++; Logger.log('  NOT EMPTY  ' + SHEET[k] + ': ' + n + ' rows'); }
    else Logger.log('  empty      ' + SHEET[k]);
  });
  if (notEmpty) problems.push(notEmpty + ' target tables already contain data. Run ' +
                              'migrateDANGERClearMigratedData() first, or skip migration.');

  // -- verdict --------------------------------------------------------------
  Logger.log('');
  if (problems.length) {
    Logger.log('PREFLIGHT FAILED — ' + problems.length + ' problem(s):');
    problems.forEach(p => Logger.log('  - ' + p));
  } else {
    Logger.log('PREFLIGHT PASSED — safe to run migrateAll()');
    if (warnings.length) {
      Logger.log('With ' + warnings.length + ' thing(s) worth a look:');
      warnings.forEach(w => Logger.log('  - ' + w));
    }
  }
  return { ok: !problems.length, problems: problems, warnings: warnings, scan: scan };
}


/**
 * Find the two known issues in the source surcharge cards.
 *   overlaps  — a period starting on or before the previous period's end
 *   truncated — the final period for a key ending mid-month inside the horizon
 */
function scanSurchargeIssues_() {
  const horizonEnd = monthStart(configDate('HORIZON_START', '2026-01-01'));
  const months     = configInt('HORIZON_MONTHS', 36);
  const lastMonth  = new Date(horizonEnd.getFullYear(), horizonEnd.getMonth() + months - 1, 1);

  const groups = {};
  [[SRC.FUEL, 380], [SRC.OTHER, 1038]].forEach(pair => {
    sourceDataRows_(pair[0]).forEach(r => {
      const key = safeInt(r[0]) + '|' + toCode_(r[8]);
      (groups[key] = groups[key] || []).push({
        modellingId: safeInt(r[0]), code: toCode_(r[8]),
        from: normaliseDate(r[9]), to: normaliseDate(r[10])
      });
    });
  });

  const overlaps = [], truncated = [];
  for (const key in groups) {
    const g = groups[key].filter(x => x.from && x.to).sort((a, b) => dateKey(a.from) - dateKey(b.from));
    for (let i = 1; i < g.length; i++) {
      if (dateKey(g[i].from) <= dateKey(g[i - 1].to)) {
        overlaps.push({ modellingId: g[i].modellingId, code: g[i].code,
                        from: g[i].from, prevTo: g[i - 1].to });
      }
    }
    if (g.length) {
      const last = g[g.length - 1];
      const isMonthEnd = dateKey(last.to) === dateKey(monthEnd(last.to));
      const insideHorizon = dateKey(last.to) <= dateKey(monthEnd(lastMonth));
      if (!isMonthEnd && insideHorizon) {
        truncated.push({ modellingId: last.modellingId, code: last.code, to: last.to });
      }
    }
  }
  return { overlaps: overlaps, truncated: truncated };
}


// ─────────────────────────────────────────────────────────────────────────────
// 2. MIGRATE
// ─────────────────────────────────────────────────────────────────────────────

function migrateAll() {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log('=== MIGRATION ===');

  const pre = migratePreflight();
  if (!pre.ok) {
    Logger.log('');
    Logger.log('STOPPED — preflight failed. Nothing was written. Fix the problems above.');
    return pre;
  }

  Logger.log('');
  Logger.log('--- loading ---');
  const stats = {};
  stats.highLevel = migrateHighLevelIds_();
  stats.modelling = migrateModellingIds_();
  stats.rateBase  = migrateRateBase_();
  stats.surcharge = migrateRateSurcharge_();
  stats.mixMethod = migrateMixMethod_();
  stats.mixLP     = migrateMixLetterParcel_();
  stats.mixCC     = migrateMixColdChain_();
  stats.fx        = migrateFxRates_();

  Logger.log('');
  Logger.log('=== MIGRATION DONE in ' + (Date.now() - t0) + 'ms ===');
  for (const k in stats) Logger.log('  ' + k + ': ' + stats[k] + ' rows');
  Logger.log('');
  Logger.log('Next: run migrateVerify()');
  return stats;
}


/** Write rows to a target table in one call, and keep the cache honest. */
function writeRows_(tableKey, rows) {
  if (!rows.length) return 0;
  const t  = TABLES[tableKey];
  const sh = getSheet_(t.sheet);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, t.headers.length).setValues(rows);
  invalidateSheetCache_(t.sheet);
  Logger.log('  ' + t.sheet + ': ' + rows.length + ' rows');
  return rows.length;
}

function stampNew_(row, C) {
  const now = new Date();
  if (C.Created_TS !== undefined) row[C.Created_TS] = now;
  if (C.Created_By !== undefined) row[C.Created_By] = 'migration';
  if (C.Updated_TS !== undefined) row[C.Updated_TS] = now;
  if (C.Updated_By !== undefined) row[C.Updated_By] = 'migration';
  return row;
}


// ---- High Level IDs --------------------------------------------------------
function migrateHighLevelIds_() {
  const C = COL.HIGH_LEVEL_IDS, rows = [];
  const currencyByGeo = { GB: 'GBP', US: 'USD', DE: 'EUR', CA: 'CAD' };

  sourceDataRows_(SRC.HIGH_LEVEL).forEach(r => {
    const brand = toCode_(r[1]), geo = toCode_(r[2]),
          tt = toCode_(r[3]), wl = toCode_(r[4]);
    if (!brand) return;
    const row = blankRow_('HIGH_LEVEL_IDS');
    row[C.High_Level_ID]   = safeInt(r[0]);
    row[C.High_Level_Code] = [brand, geo, tt, wl].filter(x => x && x !== '*').join('_');
    row[C.Brand] = brand; row[C.Geo] = geo;
    row[C.Treatment_Type] = tt; row[C.WL_Split] = wl;
    row[C.Currency] = currencyByGeo[geo] || 'GBP';
    row[C.Active] = true;
    row[C.Sort_Order] = safeInt(r[0]) * 10;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('HIGH_LEVEL_IDS', rows);
}


// ---- Modelling IDs ---------------------------------------------------------
function migrateModellingIds_() {
  const C = COL.MODELLING_IDS, rows = [];

  const hlCode = {};
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) hlCode[safeInt(hl[i][H.High_Level_ID])] = safeStr(hl[i][H.High_Level_Code]);

  sourceDataRows_(SRC.MODELLING).forEach(r => {
    const carrier = toCode_(r[6]), method = toCode_(r[7]), lp = letterParcelCode_(r[8]);
    if (!carrier) return;
    const hlid = safeInt(r[1]);
    const row = blankRow_('MODELLING_IDS');
    row[C.Modelling_ID]   = safeInt(r[0]);
    row[C.High_Level_ID]  = hlid;
    row[C.Carrier_Code]   = carrier;
    row[C.Method_Code]    = method;
    row[C.Letter_Parcel]  = lp;
    row[C.Modelling_Code] = (hlCode[hlid] || ('HL' + hlid)) + '__' + carrier + '__' + method + '__' + lp.charAt(0);
    row[C.Active] = true;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('MODELLING_IDS', rows);
}


// ---- Base rates ------------------------------------------------------------
function migrateRateBase_() {
  const C = COL.RATE_BASE, rows = [];
  const cur = currencyByModellingId_();
  let id = 0;

  sourceDataRows_(SRC.BASE).forEach(r => {
    const mid = safeInt(r[0]);
    const from = normaliseDate(r[8]), to = normaliseDate(r[9]);
    if (!mid || !from || !to) return;
    const row = blankRow_('RATE_BASE');
    row[C.Rate_ID]      = ++id;
    row[C.Modelling_ID] = mid;
    row[C.Valid_From]   = from;
    row[C.Valid_To]     = to;
    row[C.Base_Rate]    = safeNum(r[10]);
    row[C.Currency]     = cur[mid] || 'GBP';
    row[C.Scenario_ID]  = MIGRATION_SCENARIO_ID;
    row[C.Source_Ref]   = 'Postage_RFQ3_New.xlsx / Base Rate Card';
    row[C.Active]       = true;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('RATE_BASE', rows);
}

function currencyByModellingId_() {
  const out = {};
  const hlCur = {}, hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) hlCur[safeInt(hl[i][H.High_Level_ID])] = safeStr(hl[i][H.Currency]);
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < md.length; i++) out[safeInt(md[i][M.Modelling_ID])] = hlCur[safeInt(md[i][M.High_Level_ID])] || 'GBP';
  return out;
}


// ---- Surcharges (fuel + other), with the two fixes applied -----------------
function migrateRateSurcharge_() {
  const C = COL.RATE_SURCHARGE;
  const cur = currencyByModellingId_();
  const horizonStart = monthStart(configDate('HORIZON_START', '2026-01-01'));
  const months = configInt('HORIZON_MONTHS', 36);
  const lastMonth = new Date(horizonStart.getFullYear(), horizonStart.getMonth() + months - 1, 1);

  // Collect from both source tabs
  const all = [];
  [[SRC.FUEL, 'PCT'], [SRC.OTHER, 'AMT']].forEach(pair => {
    sourceDataRows_(pair[0]).forEach(r => {
      const mid = safeInt(r[0]);
      const from = normaliseDate(r[9]), to = normaliseDate(r[10]);
      if (!mid || !from || !to) return;
      all.push({ modellingId: mid, code: toCode_(r[8]), valueType: pair[1],
                 from: from, to: to, value: safeNum(r[11]), src: pair[0] });
    });
  });

  // --- Fix 1: close overlaps (Peak periods starting before the previous ends)
  const byKey = {};
  all.forEach(x => { const k = x.modellingId + '|' + x.code; (byKey[k] = byKey[k] || []).push(x); });
  let fixedOverlap = 0;
  for (const k in byKey) {
    const g = byKey[k].sort((a, b) => dateKey(a.from) - dateKey(b.from));
    for (let i = 1; i < g.length; i++) {
      if (dateKey(g[i].from) <= dateKey(g[i - 1].to)) {
        const newFrom = new Date(g[i - 1].to.getTime());
        newFrom.setDate(newFrom.getDate() + 1);
        g[i].from = newFrom;
        g[i].fixed = 'overlap';
        fixedOverlap++;
      }
    }
    // --- Fix 2: extend a terminal period that stops mid-month
    const last = g[g.length - 1];
    const isMonthEnd = dateKey(last.to) === dateKey(monthEnd(last.to));
    const insideHorizon = dateKey(last.to) <= dateKey(monthEnd(lastMonth));
    if (!isMonthEnd && insideHorizon) {
      last.to = monthEnd(last.to);
      last.fixed = (last.fixed ? last.fixed + '+' : '') + 'terminal';
    }
  }
  const fixedTerminal = all.filter(x => x.fixed && x.fixed.indexOf('terminal') >= 0).length;

  if (fixedOverlap || fixedTerminal) {
    Logger.log('  MIGRATION FIXES APPLIED:');
    if (fixedOverlap)  Logger.log('    ' + fixedOverlap + ' overlapping periods: Valid_From moved to the day after the previous period');
    if (fixedTerminal) Logger.log('    ' + fixedTerminal + ' terminal periods: Valid_To extended to month end');
    logAction_('MIGRATION_FIX', SHEET.RATE_SURCHARGE, '',
               'overlaps=' + fixedOverlap + ' terminal=' + fixedTerminal);
  } else {
    Logger.log('  No fixes needed — source data already clean.');
  }

  let id = 0;
  const rows = all.map(x => {
    const row = blankRow_('RATE_SURCHARGE');
    row[C.Surcharge_Rate_ID] = ++id;
    row[C.Modelling_ID]  = x.modellingId;
    row[C.Surcharge_Code] = x.code;
    row[C.Valid_From]    = x.from;
    row[C.Valid_To]      = x.to;
    row[C.Value]         = x.value;
    row[C.Currency]      = (x.valueType === 'AMT') ? (cur[x.modellingId] || 'GBP') : '';
    row[C.Scenario_ID]   = MIGRATION_SCENARIO_ID;
    row[C.Source_Ref]    = 'Postage_RFQ3_New.xlsx / ' + x.src + (x.fixed ? ' [fixed: ' + x.fixed + ']' : '');
    row[C.Active]        = true;
    if (x.fixed) row[C.Notes] = 'Corrected during migration: ' + x.fixed;
    return stampNew_(row, C);
  });
  return writeRows_('RATE_SURCHARGE', rows);
}


// ---- Method mix ------------------------------------------------------------
function migrateMixMethod_() {
  const C = COL.MIX_METHOD, rows = [];
  let id = 0;
  sourceDataRows_(SRC.METHOD_MIX).forEach(r => {
    const mid = safeInt(r[0]);
    const from = normaliseDate(r[9]), to = normaliseDate(r[10]);
    if (!mid || !from || !to) return;
    const row = blankRow_('MIX_METHOD');
    row[C.Mix_ID]        = ++id;
    row[C.Modelling_ID]  = mid;
    row[C.Temp_Regime]   = tempRegimeCode_(r[7]);
    row[C.Valid_From]    = from;
    row[C.Valid_To]      = to;
    row[C.Mix_Pct]       = safeNum(r[11]);
    row[C.Scenario_ID]   = MIGRATION_SCENARIO_ID;
    row[C.Active]        = true;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('MIX_METHOD', rows);
}


// ---- Letter/parcel mix -----------------------------------------------------
function migrateMixLetterParcel_() {
  const C = COL.MIX_LETTERPARCEL, rows = [];
  let id = 0;
  sourceDataRows_(SRC.LP_MIX).forEach(r => {
    const hlid = safeInt(r[0]);
    const from = normaliseDate(r[4]), to = normaliseDate(r[5]);
    if (!hlid || !from || !to) return;
    const row = blankRow_('MIX_LETTERPARCEL');
    row[C.LP_Mix_ID]       = ++id;
    row[C.High_Level_ID]   = hlid;
    row[C.Valid_From]      = from;
    row[C.Valid_To]        = to;
    row[C.Letter_Mix_Pct]  = safeNum(r[6]);
    row[C.Scenario_ID]     = MIGRATION_SCENARIO_ID;
    row[C.Active]          = true;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('MIX_LETTERPARCEL', rows);
}


// ---- Cold chain ------------------------------------------------------------
function migrateMixColdChain_() {
  const C = COL.MIX_COLDCHAIN, rows = [];
  let id = 0;
  sourceDataRows_(SRC.COLD_CHAIN).forEach(r => {
    const hlid = safeInt(r[0]);
    const from = normaliseDate(r[5]), to = normaliseDate(r[6]);
    if (!hlid || !from || !to) return;
    const row = blankRow_('MIX_COLDCHAIN');
    row[C.CC_Mix_ID]     = ++id;
    row[C.High_Level_ID] = hlid;
    row[C.Valid_From]    = from;
    row[C.Valid_To]      = to;
    row[C.CC_Mix_Pct]    = safeNum(r[7]);
    row[C.Scenario_ID]   = MIGRATION_SCENARIO_ID;
    row[C.Active]        = true;
    rows.push(stampNew_(row, C));
  });
  return writeRows_('MIX_COLDCHAIN', rows);
}


// ---- FX rates: wide source -> long target ----------------------------------
function migrateFxRates_() {
  const C = COL.FX_RATES;
  const all = readSourceTab_(SRC.FX);
  if (all.length < 3) { Logger.log('  FX_Rates: nothing to migrate'); return 0; }

  // Header row is row 2 (index 1): col A month no, col B month date, C.. currencies
  const hdr = all[1] || [];
  const cols = [];
  const seen = {};
  for (let c = 2; c < hdr.length; c++) {
    const code = safeStr(hdr[c]).toUpperCase();
    if (!/^[A-Z]{3}$/.test(code)) continue;      // skips the 'cad/usd' helper column
    if (seen[code]) continue;                    // EUR appears three times
    seen[code] = true;
    cols.push({ idx: c, code: code });
  }

  const rows = [];
  let id = 0;
  for (let i = 2; i < all.length; i++) {
    const r = all[i] || [];
    const ms = monthStart(normaliseDate(r[1]));
    if (!ms) continue;
    cols.forEach(c => {
      const v = safeNum(r[c.idx]);
      if (!v) return;
      const row = blankRow_('FX_RATES');
      row[C.FX_ID]        = ++id;
      row[C.Month_Start]  = ms;
      row[C.Currency]     = c.code;
      row[C.Rate_To_GBP]  = v;
      row[C.Source]       = 'Postage_RFQ3_New.xlsx / FX Rates';
      row[C.Active]       = true;
      rows.push(row);
    });
  }

  addMissingCurrencies_(cols.map(c => c.code));
  return writeRows_('FX_RATES', rows);
}

/** FX brings currencies the seed didn't know about (AUD, BRL, DKK, SEK). */
function addMissingCurrencies_(codes) {
  const data = getAllData_(SHEET.DIM_REFERENCE), C = COL.DIM_REFERENCE;
  const have = {};
  let maxId = 0, maxSort = 0;
  for (let i = 1; i < data.length; i++) {
    if (safeStr(data[i][C.List_Name]) === 'CURRENCY') {
      have[safeStr(data[i][C.Code]).toUpperCase()] = true;
      maxSort = Math.max(maxSort, safeInt(data[i][C.Sort_Order]));
    }
    maxId = Math.max(maxId, safeInt(data[i][C.Ref_ID]));
  }
  const add = [];
  codes.forEach(code => {
    if (have[code]) return;
    have[code] = true;
    const row = blankRow_('DIM_REFERENCE');
    row[C.Ref_ID] = ++maxId; row[C.List_Name] = 'CURRENCY';
    row[C.Code] = code;      row[C.Label] = code;
    row[C.Sort_Order] = (maxSort += 10);
    row[C.Active] = true;
    row[C.Notes] = 'Added during migration from the FX Rates tab';
    add.push(row);
  });
  if (add.length) {
    const sh = getSheet_(SHEET.DIM_REFERENCE);
    sh.getRange(sh.getLastRow() + 1, 1, add.length, TABLES.DIM_REFERENCE.headers.length).setValues(add);
    invalidateSheetCache_(SHEET.DIM_REFERENCE);
    Logger.log('  Dim_Reference: +' + add.length + ' currencies (' + add.map(r => r[C.Code]).join(', ') + ')');
  }
}


/** Minimal audit writer, so migration is recorded like any other change. */
function logAction_(action, entity, entityId, detail) {
  try {
    const sh = getSheet_(SHEET.AUDIT_LOG), C = COL.AUDIT_LOG;
    const row = blankRow_('AUDIT_LOG');
    row[C.Log_ID]    = getNextId_(SHEET.AUDIT_LOG, C.Log_ID);
    row[C.TS]        = new Date();
    row[C.Email]     = (Session.getActiveUser && Session.getActiveUser().getEmail()) || 'migration';
    row[C.Action]    = action;
    row[C.Entity]    = entity;
    row[C.Entity_ID] = entityId;
    row[C.Detail]    = detail;
    row[C.Success]   = true;
    sh.appendRow(row);
  } catch (e) {
    Logger.log('  (could not write audit row: ' + e.message + ')');
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// 3. VERIFY
// ─────────────────────────────────────────────────────────────────────────────

function migrateVerify() {
  requireMaintenance_();
  Logger.log('=== MIGRATION VERIFICATION ===');
  const problems = [];

  const expect = { HIGH_LEVEL_IDS: 17, MODELLING_IDS: 232, RATE_BASE: 264,
                   RATE_SURCHARGE: 1418, MIX_METHOD: 330, MIX_LETTERPARCEL: 17,
                   MIX_COLDCHAIN: 139 };
  Logger.log('--- row counts ---');
  for (const k in expect) {
    const n = Math.max(getSheet_(SHEET[k]).getLastRow() - 1, 0);
    const ok = (n === expect[k]);
    Logger.log('  ' + (ok ? 'OK   ' : 'CHECK') + ' ' + SHEET[k] + ': ' + n + ' (expected ' + expect[k] + ')');
    if (!ok) problems.push(SHEET[k] + ' has ' + n + ' rows, expected ' + expect[k]);
  }
  const fx = Math.max(getSheet_(SHEET.FX_RATES).getLastRow() - 1, 0);
  Logger.log('  info  ' + SHEET.FX_RATES + ': ' + fx + ' rows (reference only, not used by the engine)');

  // -- foreign keys ---------------------------------------------------------
  Logger.log('');
  Logger.log('--- referential integrity ---');
  const hlIds = {}, hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) hlIds[safeInt(hl[i][H.High_Level_ID])] = true;

  const mdIds = {}, md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  let badHl = 0, badMethodCarrier = 0;
  const methodCarrier = {}, dm = getAllData_(SHEET.DIM_METHOD), DM = COL.DIM_METHOD;
  for (let i = 1; i < dm.length; i++) methodCarrier[safeStr(dm[i][DM.Method_Code])] = safeStr(dm[i][DM.Carrier_Code]);

  const mismatches = [];
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    mdIds[id] = true;
    if (!hlIds[safeInt(md[i][M.High_Level_ID])]) badHl++;
    const mc = safeStr(md[i][M.Method_Code]), cc = safeStr(md[i][M.Carrier_Code]);
    if (methodCarrier[mc] && methodCarrier[mc] !== cc) {
      badMethodCarrier++;
      if (mismatches.length < 5) mismatches.push('Modelling ID ' + id + ': carrier ' + cc +
                                                 ' but ' + mc + ' belongs to ' + methodCarrier[mc]);
    }
  }
  Logger.log('  Modelling_IDs with unknown High_Level_ID: ' + badHl + (badHl ? '  PROBLEM' : '  OK'));
  if (badHl) problems.push(badHl + ' Modelling_IDs point at a missing High_Level_ID');

  Logger.log('  METHOD_CARRIER mismatches: ' + badMethodCarrier + (badMethodCarrier ? '  (expected: 2)' : ''));
  mismatches.forEach(m => Logger.log('      ' + m));
  if (badMethodCarrier > 2) problems.push('More METHOD_CARRIER mismatches than the 2 expected');

  ['RATE_BASE', 'RATE_SURCHARGE', 'MIX_METHOD'].forEach(k => {
    const data = getAllData_(SHEET[k]), c = COL[k];
    let bad = 0;
    for (let i = 1; i < data.length; i++) if (!mdIds[safeInt(data[i][c.Modelling_ID])]) bad++;
    Logger.log('  ' + SHEET[k] + ' rows with unknown Modelling_ID: ' + bad + (bad ? '  PROBLEM' : '  OK'));
    if (bad) problems.push(bad + ' rows in ' + SHEET[k] + ' point at a missing Modelling_ID');
  });

  // -- the 100% mix invariant ----------------------------------------------
  Logger.log('');
  Logger.log('--- method mix totals (the critical invariant) ---');
  const bad = checkMixSums_();
  if (!bad.length) {
    Logger.log('  OK   every High Level ID / regime / class / month totals 100%');
  } else {
    Logger.log('  PROBLEM  ' + bad.length + ' combinations do not total 100%:');
    bad.slice(0, 10).forEach(b => Logger.log('      HLID ' + b.hlid + ' ' + b.regime + ' ' +
        b.lp + ' ' + b.month + ': ' + (b.total * 100).toFixed(2) + '%'));
    problems.push(bad.length + ' method mix combinations do not total 100%');
  }

  Logger.log('');
  if (problems.length) {
    Logger.log('VERIFICATION FAILED — ' + problems.length + ' problem(s):');
    problems.forEach(p => Logger.log('  - ' + p));
  } else {
    Logger.log('VERIFICATION PASSED — ready for step 3 (the engine)');
  }
  return { ok: !problems.length, problems: problems };
}


function checkMixSums_() {
  const cal = getAllData_(SHEET.DIM_CALENDAR), K = COL.DIM_CALENDAR;
  const months = [];
  for (let i = 1; i < cal.length; i++) {
    if (safeBool(cal[i][K.In_Horizon])) months.push({
      start: normaliseDate(cal[i][K.Month_Start]), label: safeStr(cal[i][K.Month_Label]) });
  }

  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const meta = {};
  for (let i = 1; i < md.length; i++) {
    meta[safeInt(md[i][M.Modelling_ID])] = {
      hlid: safeInt(md[i][M.High_Level_ID]), lp: safeStr(md[i][M.Letter_Parcel]) };
  }

  const mix = getAllData_(SHEET.MIX_METHOD), X = COL.MIX_METHOD;
  const rows = [];
  for (let i = 1; i < mix.length; i++) {
    const m = meta[safeInt(mix[i][X.Modelling_ID])];
    if (!m) continue;
    rows.push({ hlid: m.hlid, lp: m.lp, regime: safeStr(mix[i][X.Temp_Regime]),
                from: dateKey(mix[i][X.Valid_From]), to: dateKey(mix[i][X.Valid_To]),
                pct: safeNum(mix[i][X.Mix_Pct]) });
  }

  const bad = [];
  months.forEach(mo => {
    const k = dateKey(mo.start);
    const totals = {};
    rows.forEach(r => {
      if (r.from <= k && r.to >= k) {
        const key = r.hlid + '|' + r.regime + '|' + r.lp;
        totals[key] = (totals[key] || 0) + r.pct;
      }
    });
    for (const key in totals) {
      if (Math.abs(totals[key] - 1) > 0.0005) {
        const p = key.split('|');
        bad.push({ hlid: p[0], regime: p[1], lp: p[2], month: mo.label, total: totals[key] });
      }
    }
  });
  return bad;
}


// ─────────────────────────────────────────────────────────────────────────────
// RESET — clears only the migrated tables, leaves dimensions and config alone
// ─────────────────────────────────────────────────────────────────────────────

function migrateDANGERClearMigratedData() {
  requireMaintenance_();
  const CONFIRM = false;   // <- set to true to actually run
  if (!CONFIRM) {
    Logger.log('Refused. Set CONFIRM = true inside migrateDANGERClearMigratedData to proceed.');
    return;
  }
  ['HIGH_LEVEL_IDS','MODELLING_IDS','RATE_BASE','RATE_SURCHARGE',
   'MIX_METHOD','MIX_LETTERPARCEL','MIX_COLDCHAIN','FX_RATES'].forEach(k => {
    const sh = getSheet_(SHEET[k]);
    if (sh.getLastRow() > 1) {
      clearDataRows_(SHEET[k]);
      Logger.log('Cleared ' + SHEET[k]);
    }
  });
  Logger.log('Done. Dimensions, calendar, roles and config are untouched.');
}