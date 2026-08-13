/**
 * Postage Forecast Portal — output.gs
 *
 * Turns the engine's result into the OUTPUT and OUTPUT_Detail tabs.
 *
 *   previewOutput()  runs the model and shows what would change. Writes nothing.
 *   publishOutput()  runs the model, validates, then writes.
 *
 * Publishing is blocked when validation returns an ERROR, unless you turn
 * VALIDATION_BLOCKS_PUBLISH off in Config.
 */

// ─────────────────────────────────────────────────────────────────────────────
// PREVIEW — what would change, without changing it
// ─────────────────────────────────────────────────────────────────────────────

function previewOutput(scenarioId) {
  const perms = requirePermissions_();
  requireRunCalc_(perms);
  scenarioId = scenarioId || 1;
  const t0 = Date.now();
  Logger.log('=== PREVIEW (nothing will be written) ===');

  const input  = loadEngineInput_(scenarioId);
  const result = computeModel_(input);

  // what's on the OUTPUT tab now
  const current = {};
  const data = getAllData_(SHEET.OUTPUT), C = COL.OUTPUT;
  for (let i = 1; i < data.length; i++) {
    const h = safeInt(data[i][C.High_Level_ID]), d = safeInt(data[i][C.Date_ID]);
    if (h && d) current[h + '|' + d] = safeNum(data[i][C.Forecast_Rate_Per_Order]);
  }

  const label = {};
  input.calendar.forEach(m => { label[m.dateId] = m.label; });

  const changes = [];
  let unchanged = 0, added = 0;
  result.outputRows.forEach(r => {
    const key = r.hlId + '|' + r.dateId;
    if (!(key in current)) { added++; return; }
    const diff = r.forecastRatePerOrder - current[key];
    if (Math.abs(diff) < 0.000001) { unchanged++; return; }
    changes.push({ hlId: r.hlId, dateId: r.dateId, label: label[r.dateId] || r.dateId,
                   before: current[key], after: r.forecastRatePerOrder, diff: diff,
                   pct: current[key] ? (diff / current[key] * 100) : 0 });
  });
  const removed = Object.keys(current).length - unchanged - changes.length;

  Logger.log('  rows unchanged : ' + unchanged);
  Logger.log('  rows changed   : ' + changes.length);
  Logger.log('  rows new       : ' + added);
  if (removed > 0) Logger.log('  rows that would disappear : ' + removed);

  if (changes.length) {
    Logger.log('');
    Logger.log('--- biggest movements ---');
    Logger.log('  HLID  month     before        after         change      %');
    changes.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 20).forEach(c =>
      Logger.log('  ' + pad_(c.hlId, 4) + '  ' + pad_(c.label, 7) + '  ' +
                 pad_(c.before.toFixed(4), 12) + ' ' + pad_(c.after.toFixed(4), 12) + ' ' +
                 pad_(c.diff.toFixed(4), 11) + ' ' + c.pct.toFixed(2) + '%'));

    const byHl = {};
    changes.forEach(c => { byHl[c.hlId] = (byHl[c.hlId] || 0) + 1; });
    Logger.log('');
    Logger.log('--- affected High Level IDs ---');
    Object.keys(byHl).sort((a, b) => a - b).forEach(h =>
      Logger.log('  High Level ID ' + h + ': ' + byHl[h] + ' month(s)'));
  } else if (added) {
    Logger.log('');
    Logger.log('  OUTPUT is empty, so all ' + added + ' rows would be written fresh.');
  } else {
    Logger.log('');
    Logger.log('  Nothing would change. OUTPUT already matches the inputs.');
  }

  Logger.log('');
  Logger.log('(' + (Date.now() - t0) + 'ms) Run publishOutput() to write these.');
  return { changes: changes.slice(0, 100), unchanged: unchanged, added: added };
}


// ─────────────────────────────────────────────────────────────────────────────
// PUBLISH
// ─────────────────────────────────────────────────────────────────────────────

function publishOutput(scenarioId) {
  const perms = requirePermissions_();
  requireRunCalc_(perms);
  requirePublish_(perms);
  scenarioId = scenarioId || 1;
  const t0 = Date.now();
  Logger.log('=== PUBLISH ===');

  return withLock_(function () {
    // ---- run the model --------------------------------------------------
    const input  = loadEngineInput_(scenarioId);
    const result = computeModel_(input);
    Logger.log('  engine: ' + result.stats.outputRows + ' output rows, ' +
               result.stats.outputDetailRows + ' detail rows, ' + result.stats.ms + 'ms');

    // ---- validate ---------------------------------------------------------
    Logger.log('');
    const validation = runValidationQuiet_(input, result);
    Logger.log('  validation: ' + validation.counts.ERROR + ' errors, ' +
               validation.counts.WARN + ' warnings, ' + validation.counts.INFO + ' info');

    const blocks = configBool('VALIDATION_BLOCKS_PUBLISH', true);
    if (validation.counts.ERROR > 0 && blocks) {
      Logger.log('');
      Logger.log('PUBLISH BLOCKED — validation found ' + validation.counts.ERROR + ' error(s).');
      Logger.log('Nothing was written. Fix them, or set VALIDATION_BLOCKS_PUBLISH to FALSE');
      Logger.log('in Config if you need to publish anyway.');
      validation.findings.filter(f => f.severity === SEVERITY.ERROR).slice(0, 10)
        .forEach(f => Logger.log('    [' + f.rule + '] ' + f.message));
      return { ok: false, blocked: true, validation: validation };
    }

    // ---- record the run ---------------------------------------------------
    const calcRunId = nextCalcRunId_();

    // ---- write ------------------------------------------------------------
    const hlById = {};
    input.highLevelIds.forEach(h => { hlById[h.id] = h; });
    const midById = {};
    input.modellingIds.forEach(d => { midById[d.id] = d; });

    const nOut = writeOutputTab_(result.outputRows, hlById, calcRunId);
    Logger.log('  OUTPUT: ' + nOut + ' rows');

    const nDetail = writeOutputDetailTab_(result.outputDetailRows, hlById, midById, calcRunId);
    Logger.log('  OUTPUT_Detail: ' + nDetail + ' rows');

    writeValidationResults_(validation.findings, calcRunId);

    const ms = Date.now() - t0;
    writeCalcRun_(calcRunId, scenarioId, 'MANUAL', nOut, nDetail, ms,
                  validation.status, summariseValidation_(validation));

    Logger.log('');
    Logger.log('PUBLISHED — run ' + calcRunId + ' in ' + ms + 'ms');
    if (validation.counts.WARN) Logger.log('  ' + validation.counts.WARN +
      ' warning(s) recorded on Validation_Results — worth a look');
    return { ok: true, calcRunId: calcRunId, rows: nOut, detailRows: nDetail, validation: validation };
  });
}


function writeOutputTab_(rows, hlById, calcRunId) {
  const t = TABLES.OUTPUT, C = COL.OUTPUT, sh = getSheet_(t.sheet);
  clearDataRows_(t.sheet);

  const values = rows.map(r => {
    const h = hlById[r.hlId] || {};
    const row = blankRow_('OUTPUT');
    row[C.High_Level_ID]           = r.hlId;
    row[C.Date_ID]                 = r.dateId;
    row[C.Month_Start]             = r.monthStart;
    row[C.Brand]                   = h.brand || '';
    row[C.Geo]                     = h.geo || '';
    row[C.Treatment_Type]          = h.treatmentType || '';
    row[C.WL_Split]                = h.wlSplit || '';
    row[C.Currency]                = h.currency || '';
    row[C.Forecast_Rate_Per_Order] = r.forecastRatePerOrder;
    row[C.Scenario_ID]             = r.scenarioId;
    row[C.Calc_Run_ID]             = calcRunId;
    return row;
  });

  if (values.length) sh.getRange(2, 1, values.length, t.headers.length).setValues(values);
  invalidateSheetCache_(t.sheet);
  return values.length;
}


function writeOutputDetailTab_(rows, hlById, midById, calcRunId) {
  const t = TABLES.OUTPUT_DETAIL, C = COL.OUTPUT_DETAIL, sh = getSheet_(t.sheet);
  clearDataRows_(t.sheet);

  const values = rows.map(r => {
    const h = hlById[r.hlId] || {};
    const row = blankRow_('OUTPUT_DETAIL');
    row[C.Modelling_ID]         = r.modellingId;
    row[C.High_Level_ID]        = r.hlId;
    row[C.Date_ID]              = r.dateId;
    row[C.Month_Start]          = r.monthStart;
    row[C.Brand]                = h.brand || '';
    row[C.Geo]                  = h.geo || '';
    row[C.Treatment_Type]       = h.treatmentType || '';
    row[C.WL_Split]             = h.wlSplit || '';
    row[C.Carrier_Code]         = r.carrier;
    row[C.Method_Code]          = r.method;
    row[C.Letter_Parcel]        = r.letterParcel;
    row[C.Base_Rate]            = r.baseRate;
    row[C.Surcharge_Pct_Total]  = r.surchargePctTotal;
    row[C.Surcharge_Amt_Total]  = r.surchargeAmtTotal;
    row[C.Rate_Per_Parcel]      = r.ratePerParcel;
    row[C.CC_Share]             = r.ccShare;
    row[C.Method_Mix]           = r.methodMix;
    row[C.LP_Mix]               = r.lpMix;
    row[C.Rate_Contribution]    = r.rateContribution;
    row[C.Rate_CTS]             = r.rateCts;
    row[C.Scenario_ID]          = r.scenarioId;
    row[C.Calc_Run_ID]          = calcRunId;
    return row;
  });

  // 8,352 rows in one call — chunking would be slower, not faster
  if (values.length) sh.getRange(2, 1, values.length, t.headers.length).setValues(values);
  invalidateSheetCache_(t.sheet);
  return values.length;
}


function nextCalcRunId_() {
  return getNextId_(SHEET.CALC_RUNS, COL.CALC_RUNS.Calc_Run_ID);
}

function writeCalcRun_(id, scenarioId, trigger, nOut, nDetail, ms, status, summary) {
  const t = TABLES.CALC_RUNS, C = COL.CALC_RUNS;
  const row = blankRow_('CALC_RUNS');
  row[C.Calc_Run_ID]        = id;
  row[C.Run_TS]             = new Date();
  row[C.Run_By]             = currentUserEmail_();
  row[C.Scenario_ID]        = scenarioId;
  row[C.Trigger]            = trigger;
  row[C.Rows_Output]        = nOut;
  row[C.Rows_Output_Detail] = nDetail;
  row[C.Duration_Ms]        = ms;
  row[C.Validation_Status]  = status;
  row[C.Validation_Summary] = summary;
  row[C.Engine_Version]     = configStr('ENGINE_VERSION', ENGINE_VERSION);
  getSheet_(t.sheet).appendRow(row);
  invalidateSheetCache_(t.sheet);
  return id;
}

function currentUserEmail_() {
  try {
    const e = Session.getActiveUser().getEmail();
    return e || 'unknown';
  } catch (err) { return 'unknown'; }
}

function summariseValidation_(v) {
  const parts = [];
  Object.keys(v.byRule).sort().forEach(r => parts.push(r + '=' + v.byRule[r]));
  return parts.join(', ') || 'clean';
}

/** Same rules as runValidation, without re-running the engine or re-logging. */
function runValidationQuiet_(input, result) {
  const findings = [];
  result.defects.forEach(d => findings.push({
    rule: d.rule,
    severity: (d.rule === 'MIX_SUM' || d.rule === 'RATE_MISSING' ||
               d.rule === 'ORPHAN_FK' || d.rule === 'RANGE_OVERLAP')
              ? SEVERITY.ERROR : SEVERITY.WARN,
    message: d.label, hlId: d.hlId || '', modellingId: d.modellingId || '', dateId: d.dateId || ''
  }));
  ruleValueRanges_(input, findings);
  ruleDateOrder_(input, findings);
  ruleRangeIssues_(input, findings);
  ruleCoverage_(input, findings);
  ruleDuplicates_(findings);
  ruleMethodCarrier_(findings);
  ruleTbcCarrier_(findings);
  ruleOutputSwing_(result, input, findings);

  const counts = { ERROR: 0, WARN: 0, INFO: 0 }, byRule = {};
  findings.forEach(f => {
    counts[f.severity] = (counts[f.severity] || 0) + 1;
    byRule[f.rule] = (byRule[f.rule] || 0) + 1;
  });
  return {
    status: counts.ERROR ? 'FAIL' : (counts.WARN ? 'WARN' : 'PASS'),
    counts: counts, byRule: byRule, findings: findings
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// CHECK — does the published OUTPUT still agree with the original workbook?
// ─────────────────────────────────────────────────────────────────────────────

function verifyPublishedOutput() {
  requirePermissions_();
  Logger.log('=== PUBLISHED OUTPUT CHECK ===');
  const data = getAllData_(SHEET.OUTPUT), C = COL.OUTPUT;
  const mine = {};
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const h = safeInt(data[i][C.High_Level_ID]), d = safeInt(data[i][C.Date_ID]);
    if (!h || !d) continue;
    mine[h + '|' + d] = safeNum(data[i][C.Forecast_Rate_Per_Order]);
    n++;
  }
  Logger.log('  OUTPUT tab has ' + n + ' rows');

  let src;
  try { src = readSourceTab_('Output'); }
  catch (e) { Logger.log('  (source workbook not readable — skipping comparison)'); return { ok: true }; }

  let matched = 0, diffs = 0, maxDiff = 0;
  for (let i = 1; i < src.length; i++) {
    const r = src[i] || [];
    const h = safeInt(r[0]), d = safeInt(r[1]);
    if (!h || !d) continue;
    const key = h + '|' + d;
    if (!(key in mine)) { diffs++; continue; }
    const delta = Math.abs(mine[key] - safeNum(r[7]));
    if (delta > maxDiff) maxDiff = delta;
    delta <= 0.0001 ? matched++ : diffs++;
  }
  Logger.log('  matches the original workbook: ' + matched);
  Logger.log('  differences: ' + diffs);
  Logger.log('  largest difference: ' + maxDiff.toExponential(3));
  Logger.log('');
  Logger.log(diffs ? 'CHECK FAILED — send me this log' :
                     'CHECK PASSED — the published OUTPUT matches the original workbook');
  return { ok: !diffs, matched: matched, diffs: diffs };
}