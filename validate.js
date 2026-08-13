/**
 * Postage Forecast Portal — validate.gs
 *
 * The rule pack. Runs on demand and before every publish.
 *
 *   ERROR  blocks publishing — the forecast would be wrong
 *   WARN   worth looking at — probably a mistake, but the maths still works
 *   INFO   noted, expected, no action needed
 *
 * Everything here is read only apart from writing its own results.
 */

const SEVERITY = { ERROR: 'ERROR', WARN: 'WARN', INFO: 'INFO' };


function runValidation(scenarioId) {
  const perms = requirePermissions_();
  requireRunCalc_(perms);
  scenarioId = scenarioId || 1;
  const t0 = Date.now();
  Logger.log('=== VALIDATION ===');

  const input  = loadEngineInput_(scenarioId);
  const result = computeModel_(input);
  const findings = [];

  // Defects the engine spotted while calculating
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

  writeValidationResults_(findings, null);
  return reportFindings_(findings, Date.now() - t0);
}


// ─────────────────────────────────────────────────────────────────────────────
// RULES
// ─────────────────────────────────────────────────────────────────────────────

/** Percentages inside 0-100, rates not negative. */
function ruleValueRanges_(input, out) {
  input.mixMethod.forEach(r => {
    if (r.v < 0 || r.v > 1) out.push({ rule: 'MIX_NEG', severity: SEVERITY.ERROR,
      message: 'Modelling ID ' + r.modellingId + ' ' + r.regime + ' mix is ' +
               (r.v * 100).toFixed(2) + '% — must be between 0 and 100',
      modellingId: r.modellingId });
  });
  input.mixLetterParcel.forEach(r => {
    if (r.v < 0 || r.v > 1) out.push({ rule: 'LP_RANGE', severity: SEVERITY.ERROR,
      message: 'High Level ID ' + r.hlId + ' letter mix is ' + (r.v * 100).toFixed(2) + '%',
      hlId: r.hlId });
  });
  input.mixColdChain.forEach(r => {
    if (r.v < 0 || r.v > 1) out.push({ rule: 'CC_RANGE', severity: SEVERITY.ERROR,
      message: 'High Level ID ' + r.hlId + ' cold-chain mix is ' + (r.v * 100).toFixed(2) + '%',
      hlId: r.hlId });
  });
  input.rateBase.forEach(r => {
    if (r.v < 0) out.push({ rule: 'RATE_NEG', severity: SEVERITY.ERROR,
      message: 'Modelling ID ' + r.modellingId + ' has a negative base rate (' + r.v + ')',
      modellingId: r.modellingId });
  });
  input.rateSurcharge.forEach(r => {
    if (r.v < 0) out.push({ rule: 'RATE_NEG', severity: SEVERITY.WARN,
      message: 'Modelling ID ' + r.modellingId + ' ' + r.code + ' surcharge is negative (' + r.v + ')',
      modellingId: r.modellingId });
  });
}

/** Valid_From must not be after Valid_To. */
function ruleDateOrder_(input, out) {
  [['rateBase', 'Rate_Base'], ['rateSurcharge', 'Rate_Surcharge'],
   ['mixMethod', 'Mix_Method'], ['mixLetterParcel', 'Mix_LetterParcel'],
   ['mixColdChain', 'Mix_ColdChain']].forEach(pair => {
    input[pair[0]].forEach(r => {
      if (r.f > r.t) out.push({ rule: 'DATE_ORDER', severity: SEVERITY.ERROR,
        message: pair[1] + ': a row for ' + (r.modellingId ? 'Modelling ID ' + r.modellingId
                 : 'High Level ID ' + r.hlId) + ' starts after it ends',
        modellingId: r.modellingId || '', hlId: r.hlId || '' });
    });
  });
}

/**
 * Overlaps, gaps, and periods that stop mid-month.
 *
 * TERMINAL_MID_MONTH is an ERROR for day-weighted surcharges — that is the
 * December 2028 bug, where a period ending on the 1st counted for one day of
 * the month. For point-in-time tables it is only a WARN, because the value on
 * the 1st is all that matters, so nothing is actually understated.
 */
function ruleRangeIssues_(input, out) {
  const lastMonth = input.calendar.filter(m => m.inHorizon).slice(-1)[0];
  if (!lastMonth) return;
  const horizonEnd = lastMonth.meNum;

  function examine(rows, keyFn, label, dayWeighted) {
    const groups = {};
    rows.forEach(r => { const k = keyFn(r); (groups[k] = groups[k] || []).push(r); });

    for (const k in groups) {
      const g = groups[k].sort((a, b) => a.f - b.f);

      for (let i = 1; i < g.length; i++) {
        if (g[i].f <= g[i - 1].t) out.push({ rule: 'RANGE_OVERLAP', severity: SEVERITY.ERROR,
          message: label + ' ' + k + ': a period starts on or before the previous one ends' });
        else if (g[i].f > g[i - 1].t + 1 && g[i - 1].t < horizonEnd)
          out.push({ rule: 'RANGE_GAP', severity: SEVERITY.WARN,
            message: label + ' ' + k + ': ' + (g[i].f - g[i - 1].t - 1) +
                     ' day(s) with no value, inside the horizon' });
      }

      // Only day-weighted values can be understated by stopping mid-month.
      // For point-in-time tables (base rates, mixes) the value live on the 1st
      // applies to the whole month, so a period ending on the 15th is harmless —
      // and any month left genuinely uncovered is caught by RANGE_GAP instead.
      if (dayWeighted) {
        const last = g[g.length - 1];
        if (last.t <= horizonEnd) {
          const d = normaliseDate(new Date(last.t * 86400000));
          if (dayNum(monthEnd(d)) !== last.t) out.push({
            rule: 'TERMINAL_MID_MONTH', severity: SEVERITY.ERROR,
            message: label + ' ' + k + ': the final period ends ' + fmtDate(d) +
                     ', part-way through the month — that month will be understated'
          });
        }
      }
    }
  }

  examine(input.rateBase, r => 'Modelling ID ' + r.modellingId, 'Rate_Base', false);
  examine(input.rateSurcharge, r => 'Modelling ID ' + r.modellingId + ' ' + r.code,
          'Rate_Surcharge', true);
  examine(input.mixMethod, r => 'Modelling ID ' + r.modellingId + ' ' + r.regime,
          'Mix_Method', false);
  examine(input.mixLetterParcel, r => 'High Level ID ' + r.hlId, 'Mix_LetterParcel', false);
  examine(input.mixColdChain, r => 'High Level ID ' + r.hlId, 'Mix_ColdChain', false);
}

/** Every active High Level ID needs mix values covering every month. */
function ruleCoverage_(input, out) {
  const months = input.calendar.filter(m => m.inHorizon);
  const active = input.highLevelIds.filter(h => h.active !== false);

  function coverage(rows, ruleCode, severity, what) {
    const byHl = groupBy_(rows, 'hlId');
    active.forEach(h => {
      const g = byHl[h.id] || [];
      const gaps = months.filter(m => !g.some(r => r.f <= m.msNum && r.t >= m.msNum));
      if (!gaps.length) return;
      out.push({ rule: ruleCode, severity: severity, hlId: h.id,
        message: 'High Level ID ' + h.id + ' has no ' + what + ' for ' + gaps.length +
                 ' month(s), starting ' + gaps[0].label });
    });
  }

  coverage(input.mixLetterParcel, 'LP_COVERAGE', SEVERITY.ERROR, 'letter/parcel mix');
  // Cold chain is out of scope for now, so a gap is INFO: it simply means 100% ambient.
  coverage(input.mixColdChain, 'CC_COVERAGE', SEVERITY.INFO, 'cold-chain mix (treated as 100% ambient)');
}

/** The same business segment or the same route defined twice. */
function ruleDuplicates_(out) {
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS, seenHl = {};
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id || !safeBool(hl[i][H.Active])) continue;
    const k = [normKey(hl[i][H.Brand]), normKey(hl[i][H.Geo]),
               normKey(hl[i][H.Treatment_Type]), normKey(hl[i][H.WL_Split])].join('|');
    if (seenHl[k]) out.push({ rule: 'DUP_HIGH_LEVEL', severity: SEVERITY.ERROR, hlId: id,
      message: 'High Level IDs ' + seenHl[k] + ' and ' + id + ' are the same combination (' + k + ')' });
    else seenHl[k] = id;
  }

  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS, seenMd = {};
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id || !safeBool(md[i][M.Active])) continue;
    const k = [safeInt(md[i][M.High_Level_ID]), normKey(md[i][M.Carrier_Code]),
               normKey(md[i][M.Method_Code]), normKey(md[i][M.Letter_Parcel])].join('|');
    if (seenMd[k]) out.push({ rule: 'DUP_MODELLING', severity: SEVERITY.ERROR, modellingId: id,
      message: 'Modelling IDs ' + seenMd[k] + ' and ' + id + ' are the same route' });
    else seenMd[k] = id;
  }
}

/** A method filed under a carrier it doesn't belong to. */
function ruleMethodCarrier_(out) {
  const owner = {}, dm = getAllData_(SHEET.DIM_METHOD), DM = COL.DIM_METHOD;
  for (let i = 1; i < dm.length; i++) owner[safeStr(dm[i][DM.Method_Code])] = safeStr(dm[i][DM.Carrier_Code]);

  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < md.length; i++) {
    const mc = safeStr(md[i][M.Method_Code]), cc = safeStr(md[i][M.Carrier_Code]);
    if (!mc || !cc) continue;
    if (!owner[mc]) {
      out.push({ rule: 'ORPHAN_FK', severity: SEVERITY.ERROR,
        modellingId: safeInt(md[i][M.Modelling_ID]),
        message: 'Modelling ID ' + safeInt(md[i][M.Modelling_ID]) + ' uses method "' + mc +
                 '", which is not in Dim_Method' });
    } else if (owner[mc] !== cc) {
      out.push({ rule: 'METHOD_CARRIER', severity: SEVERITY.WARN,
        modellingId: safeInt(md[i][M.Modelling_ID]),
        message: 'Modelling ID ' + safeInt(md[i][M.Modelling_ID]) + ' is filed under carrier ' +
                 cc + ', but ' + mc + ' belongs to ' + owner[mc] });
    }
  }
}

/** Routes still on the placeholder carrier. */
function ruleTbcCarrier_(out) {
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const ids = [];
  for (let i = 1; i < md.length; i++) {
    if (safeStr(md[i][M.Carrier_Code]) === 'TBC' && safeBool(md[i][M.Active])) {
      ids.push(safeInt(md[i][M.Modelling_ID]));
    }
  }
  if (ids.length) out.push({ rule: 'TBC_CARRIER', severity: SEVERITY.INFO,
    message: ids.length + ' route(s) still use the placeholder carrier TBC: ' +
             ids.slice(0, 10).join(', ') + (ids.length > 10 ? '...' : '') });
}

/**
 * A rate per order jumping more than 20% month on month.
 *
 * This is the rule that catches a fat-finger. A base rate typed as 32.00 instead
 * of 3.20 passes every structural check above — it's a positive number in a valid
 * date range — but moves the forecast 900%.
 */
function ruleOutputSwing_(result, input, out) {
  const THRESHOLD = 0.20;
  const byHl = {};
  result.outputRows.forEach(r => { (byHl[r.hlId] = byHl[r.hlId] || []).push(r); });
  const label = {};
  input.calendar.forEach(m => { label[m.dateId] = m.label; });

  for (const hlId in byHl) {
    const rows = byHl[hlId].sort((a, b) => a.dateId - b.dateId);
    for (let i = 1; i < rows.length; i++) {
      const prev = rows[i - 1].forecastRatePerOrder, cur = rows[i].forecastRatePerOrder;
      if (prev <= 0) continue;
      const move = (cur - prev) / prev;
      if (Math.abs(move) > THRESHOLD) out.push({ rule: 'OUTPUT_SWING', severity: SEVERITY.WARN,
        hlId: safeInt(hlId), dateId: rows[i].dateId,
        message: 'High Level ID ' + hlId + ': rate per order moves ' + (move * 100).toFixed(1) +
                 '% into ' + (label[rows[i].dateId] || rows[i].dateId) +
                 ' (' + prev.toFixed(4) + ' to ' + cur.toFixed(4) + ')' });
    }
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// OUTPUT
// ─────────────────────────────────────────────────────────────────────────────

function writeValidationResults_(findings, calcRunId) {
  const t = TABLES.VALIDATION_RESULTS, C = COL.VALIDATION_RESULTS;
  const sh = getSheet_(t.sheet);
  clearDataRows_(t.sheet);
  if (!findings.length) return 0;

  const order = { ERROR: 0, WARN: 1, INFO: 2 };
  const sorted = findings.slice().sort((a, b) =>
    (order[a.severity] - order[b.severity]) || (a.rule < b.rule ? -1 : 1));

  const rows = sorted.map((f, i) => {
    const row = blankRow_('VALIDATION_RESULTS');
    row[C.Result_ID]     = i + 1;
    row[C.Calc_Run_ID]   = calcRunId || '';
    row[C.Rule_Code]     = f.rule;
    row[C.Severity]      = f.severity;
    row[C.Entity]        = f.entity || '';
    row[C.Entity_ID]     = f.entityId || '';
    row[C.High_Level_ID] = f.hlId || '';
    row[C.Modelling_ID]  = f.modellingId || '';
    row[C.Date_ID]       = f.dateId || '';
    row[C.Message]       = f.message;
    row[C.Resolved]      = false;
    return row;
  });
  sh.getRange(2, 1, rows.length, t.headers.length).setValues(rows);
  invalidateSheetCache_(t.sheet);
  return rows.length;
}


function reportFindings_(findings, ms) {
  const counts = { ERROR: 0, WARN: 0, INFO: 0 };
  const byRule = {};
  findings.forEach(f => {
    counts[f.severity] = (counts[f.severity] || 0) + 1;
    byRule[f.rule] = (byRule[f.rule] || 0) + 1;
  });

  Logger.log('');
  Logger.log('  ERROR : ' + counts.ERROR + (counts.ERROR ? '   <- these block publishing' : ''));
  Logger.log('  WARN  : ' + counts.WARN);
  Logger.log('  INFO  : ' + counts.INFO);

  if (findings.length) {
    Logger.log('');
    Logger.log('--- by rule ---');
    Object.keys(byRule).sort().forEach(r => Logger.log('  ' + pad_(byRule[r], 5) + '  ' + r));

    ['ERROR', 'WARN', 'INFO'].forEach(sev => {
      const list = findings.filter(f => f.severity === sev);
      if (!list.length) return;
      Logger.log('');
      Logger.log('--- ' + sev + ' (' + list.length + (list.length > 10 ? ', showing 10' : '') + ') ---');
      list.slice(0, 10).forEach(f => Logger.log('  [' + f.rule + '] ' + f.message));
    });
  }

  const status = counts.ERROR ? 'FAIL' : (counts.WARN ? 'WARN' : 'PASS');
  Logger.log('');
  if (status === 'PASS') Logger.log('VALIDATION PASSED — no problems found');
  else if (status === 'WARN') Logger.log('VALIDATION PASSED WITH WARNINGS — safe to publish, but read them');
  else Logger.log('VALIDATION FAILED — ' + counts.ERROR + ' error(s) must be fixed before publishing');
  Logger.log('Full results are on the Validation_Results tab. (' + ms + 'ms)');

  return { status: status, counts: counts, byRule: byRule, findings: findings };
}