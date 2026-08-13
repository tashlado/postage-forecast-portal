/**
 * Postage Forecast Portal — engine.gs
 *
 * This replaces the 8,352-row `Modelling` tab from the original workbook.
 *
 * computeModel_() is PURE: plain objects in, plain objects out. It never touches
 * SpreadsheetApp, Session or any Google service. That is what lets us run it
 * against your existing numbers and prove it agrees before anything depends on it.
 *
 * The maths, restated from the spec:
 *
 *   base(d,m)      = sum of base rates live on the 1st of month m          [point in time]
 *   fuelPct(d,m)   = day-weighted average % across month m                 [prorated]
 *   otherAmt(d,m)  = day-weighted average currency amount across month m   [prorated]
 *   rate(d,m)      = base x (1 + fuelPct) + otherAmt
 *   methodMix(d,m) = mixCC x ccShare + mixAmbient x (1 - ccShare)
 *   lpMix(d,m)     = letterMix if the route is a LETTER, else 1 - letterMix
 *   contribution   = rate x methodMix x lpMix
 *   OUTPUT(h,m)    = sum of contributions across every Modelling ID in h
 */

// ─────────────────────────────────────────────────────────────────────────────
// DAY NUMBERS
//
// Every date becomes an integer day number computed from year/month/day via
// Date.UTC. Comparisons and day counts are then plain integer arithmetic, so
// no timezone, daylight-saving or serial-number issue can shift a boundary.
// ─────────────────────────────────────────────────────────────────────────────

const MS_DAY = 86400000;

/**
 * A month label such as "Jan-26". If the stored value has been coerced into a
 * date or a serial number by Sheets, rebuild it from the month start instead.
 */
function monthLabel_(stored, monthStartDate) {
  const s = safeStr(stored);
  if (s && !/^[0-9.]+$/.test(s) && !isDate_(stored)) return s;
  return monthStartDate ? fmtMonthLabel(monthStartDate) : s;
}

function dayNum(value) {
  const d = normaliseDate(value);
  if (!d) return null;
  return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / MS_DAY);
}


// ─────────────────────────────────────────────────────────────────────────────
// RESOLVERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Value in force on the first of the month.
 *
 * Additive by design — see spec 2.9. On clean data it should never match more
 * than one row, so a multi-match is recorded as a defect rather than silently
 * summed into a wrong answer.
 */
function pointInTime_(rows, ms, defects, label) {
  if (!rows) return 0;
  let total = 0, matches = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r.f <= ms && r.t >= ms) { total += r.v; matches++; }
  }
  if (matches > 1 && defects) {
    defects.push({ rule: 'RANGE_OVERLAP', label: label, matches: matches, monthStart: ms });
  }
  return total;
}

/**
 * Day-weighted average across the month.
 *
 *   sum over rows of  (overlapping days x value)  /  days in month
 *
 * A rate changing on the 4th of a 31-day month therefore contributes 28/31.
 */
function dayWeighted_(rows, ms, me, dim) {
  if (!rows || !dim) return 0;
  let total = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const from = r.f > ms ? r.f : ms;
    const to   = r.t < me ? r.t : me;
    if (to < from) continue;
    total += (to - from + 1) * r.v;
  }
  return total / dim;
}


// ─────────────────────────────────────────────────────────────────────────────
// THE MODEL
// ─────────────────────────────────────────────────────────────────────────────

/**
 * @param {Object} input  { calendar, highLevelIds, modellingIds, dimSurcharge,
 *                          rateBase, rateSurcharge, mixMethod, mixLetterParcel,
 *                          mixColdChain, scenarioId }
 * @return {Object} { outputRows, outputDetailRows, defects, stats }
 */
function computeModel_(input) {
  const t0 = Date.now();
  const defects = [];

  const months = (input.calendar || []).filter(m => m.inHorizon !== false);
  const hls    = (input.highLevelIds || []).filter(h => h.active !== false);
  const mids   = (input.modellingIds || []).filter(d => d.active !== false);

  // ---- metadata lookups ---------------------------------------------------
  const hlById = {};
  hls.forEach(h => { hlById[h.id] = h; });

  const midsByHl = {};
  const validMids = [];
  mids.forEach(d => {
    if (!hlById[d.hlId]) {
      defects.push({ rule: 'ORPHAN_FK', label: 'Modelling ID ' + d.id +
                     ' points at High Level ID ' + d.hlId + ', which is missing or inactive' });
      return;
    }
    validMids.push(d);
    (midsByHl[d.hlId] = midsByHl[d.hlId] || []).push(d);
  });

  // ---- surcharge definitions ---------------------------------------------
  const surchargeDef = input.dimSurcharge || {};

  // ---- indexes: one pass, then pure lookups inside the month loop ---------
  const baseByMid = groupBy_(input.rateBase, 'modellingId');
  const mixLPByHl = groupBy_(input.mixLetterParcel, 'hlId');
  const mixCCByHl = groupBy_(input.mixColdChain, 'hlId');

  const mixMethodByMid = {};
  (input.mixMethod || []).forEach(r => {
    const k = r.modellingId;
    const slot = mixMethodByMid[k] || (mixMethodByMid[k] = { CC: [], AMBIENT: [] });
    (slot[r.regime] || (slot[r.regime] = [])).push(r);
  });

  // Modelling ID -> surcharge code -> rows, split by how they are applied
  const surchargeByMid = {};
  (input.rateSurcharge || []).forEach(r => {
    const def = surchargeDef[r.code];
    if (!def) {
      defects.push({ rule: 'ORPHAN_FK', label: 'Surcharge code "' + r.code +
                     '" on Modelling ID ' + r.modellingId + ' is not defined in Dim_Surcharge' });
      return;
    }
    const slot = surchargeByMid[r.modellingId] ||
                 (surchargeByMid[r.modellingId] = { pct: [], amt: [] });
    (def.valueType === 'PCT' ? slot.pct : slot.amt).push(
      { f: r.f, t: r.t, v: r.v, proration: def.proration });
  });

  // ---- the loop -----------------------------------------------------------
  const outputRows = [], outputDetailRows = [];
  const scenarioId = input.scenarioId || 1;

  for (let mi = 0; mi < months.length; mi++) {
    const m = months[mi];
    const ms = m.msNum, me = m.meNum, dim = m.daysInMonth;

    // per-High-Level-ID values, resolved once per month
    const ccShare = {}, letterMix = {}, hlTotal = {};
    hls.forEach(h => {
      ccShare[h.id]   = pointInTime_(mixCCByHl[h.id], ms, defects, 'Mix_ColdChain HLID ' + h.id);
      letterMix[h.id] = pointInTime_(mixLPByHl[h.id], ms, defects, 'Mix_LetterParcel HLID ' + h.id);
      hlTotal[h.id]   = 0;
    });

    for (let di = 0; di < validMids.length; di++) {
      const d = validMids[di];

      const base = pointInTime_(baseByMid[d.id], ms, defects, 'Rate_Base MID ' + d.id);
      const sur  = surchargeByMid[d.id];

      let pctTotal = 0, amtTotal = 0;
      if (sur) {
        pctTotal = resolveSurcharge_(sur.pct, ms, me, dim, defects, 'PCT MID ' + d.id);
        amtTotal = resolveSurcharge_(sur.amt, ms, me, dim, defects, 'AMT MID ' + d.id);
      }

      const rate = base * (1 + pctTotal) + amtTotal;

      const mm  = mixMethodByMid[d.id];
      const cc  = ccShare[d.hlId];
      const mixCC  = mm ? pointInTime_(mm.CC,      ms, defects, 'Mix_Method CC MID ' + d.id)      : 0;
      const mixAmb = mm ? pointInTime_(mm.AMBIENT, ms, defects, 'Mix_Method AMBIENT MID ' + d.id) : 0;
      const methodMix = mixCC * cc + mixAmb * (1 - cc);

      const lp    = letterMix[d.hlId];
      const lpMix = (d.letterParcel === 'LETTER') ? lp : (1 - lp);

      const contribution = rate * methodMix * lpMix;
      hlTotal[d.hlId] += contribution;

      // A route carrying volume but priced at zero is the silent-failure mode
      // that SUMIFS made invisible in the workbook. Surface it.
      if (base === 0 && methodMix > 0) {
        defects.push({ rule: 'RATE_MISSING', label: 'Modelling ID ' + d.id +
                       ' has mix ' + (methodMix * 100).toFixed(2) + '% in ' + m.label +
                       ' but no base rate — it is being priced at zero',
                       modellingId: d.id, dateId: m.dateId });
      }

      outputDetailRows.push({
        modellingId: d.id, hlId: d.hlId, dateId: m.dateId, monthStart: m.monthStart,
        carrier: d.carrier, method: d.method, letterParcel: d.letterParcel,
        baseRate: base, surchargePctTotal: pctTotal, surchargeAmtTotal: amtTotal,
        ratePerParcel: rate, ccShare: cc, methodMix: methodMix, lpMix: lpMix,
        rateContribution: contribution, rateCts: rate * lpMix,
        scenarioId: scenarioId
      });
    }

    hls.forEach(h => {
      outputRows.push({
        hlId: h.id, dateId: m.dateId, monthStart: m.monthStart,
        brand: h.brand, geo: h.geo, treatmentType: h.treatmentType, wlSplit: h.wlSplit,
        currency: h.currency, forecastRatePerOrder: hlTotal[h.id],
        scenarioId: scenarioId
      });
    });
  }

  // ---- the 100% invariant -------------------------------------------------
  checkMixInvariant_(months, validMids, hlById, mixMethodByMid, mixCCByHl, defects);

  return {
    outputRows: outputRows,
    outputDetailRows: outputDetailRows,
    defects: defects,
    stats: {
      months: months.length,
      modellingIds: validMids.length,
      highLevelIds: hls.length,
      outputRows: outputRows.length,
      outputDetailRows: outputDetailRows.length,
      ms: Date.now() - t0
    }
  };
}


/** Apply each surcharge row by whichever proration its type specifies. */
function resolveSurcharge_(rows, ms, me, dim, defects, label) {
  if (!rows || !rows.length) return 0;
  let total = 0;
  const pit = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].proration === 'POINT_IN_TIME') pit.push(rows[i]);
  }
  if (pit.length !== rows.length) {
    const dw = rows.filter(r => r.proration !== 'POINT_IN_TIME');
    total += dayWeighted_(dw, ms, me, dim);
  }
  if (pit.length) total += pointInTime_(pit, ms, defects, label + ' (point in time)');
  return total;
}


/**
 * For every High Level ID / temperature regime / letter-or-parcel / month, the
 * delivery method percentages must total 100%. Nothing in the original workbook
 * enforced this: set a mix to 90% and the forecast quietly came out 10% low.
 */
function checkMixInvariant_(months, mids, hlById, mixMethodByMid, mixCCByHl, defects) {
  const REGIMES = ['CC', 'AMBIENT'];

  for (let mi = 0; mi < months.length; mi++) {
    const m = months[mi], ms = m.msNum;
    const totals = {};

    for (let di = 0; di < mids.length; di++) {
      const d = mids[di];
      const mm = mixMethodByMid[d.id];
      if (!mm) continue;
      for (let ri = 0; ri < REGIMES.length; ri++) {
        const regime = REGIMES[ri];
        const pct = pointInTime_(mm[regime], ms, null, null);
        if (!pct) continue;
        const key = d.hlId + '|' + regime + '|' + d.letterParcel;
        totals[key] = (totals[key] || 0) + pct;
      }
    }

    for (const key in totals) {
      if (Math.abs(totals[key] - 1) > 0.0005) {
        const p = key.split('|');
        defects.push({
          rule: 'MIX_SUM',
          label: 'High Level ID ' + p[0] + ' ' + p[1] + ' ' + p[2] + ' in ' + m.label +
                 ' totals ' + (totals[key] * 100).toFixed(2) + '%, should be 100%',
          hlId: safeInt(p[0]), dateId: m.dateId, total: totals[key]
        });
      }
    }
  }
}


function groupBy_(rows, key) {
  const out = {};
  if (!rows) return out;
  for (let i = 0; i < rows.length; i++) {
    const k = rows[i][key];
    (out[k] = out[k] || []).push(rows[i]);
  }
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// LOADING — turns the spreadsheet into the plain objects computeModel_ wants
// ─────────────────────────────────────────────────────────────────────────────

function loadEngineInput_(scenarioId) {
  scenarioId = scenarioId || 1;
  perfMark('loadEngineInput_ start');

  prewarmSheetCache_([
    SHEET.DIM_CALENDAR, SHEET.DIM_SURCHARGE, SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS,
    SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.MIX_METHOD,
    SHEET.MIX_LETTERPARCEL, SHEET.MIX_COLDCHAIN, SHEET.CONFIG
  ]);
  perfMark('sheets fetched');

  // ---- calendar -----------------------------------------------------------
  const calendar = [];
  const cal = getAllData_(SHEET.DIM_CALENDAR), K = COL.DIM_CALENDAR;
  for (let i = 1; i < cal.length; i++) {
    const msDate = normaliseDate(cal[i][K.Month_Start]);
    if (!msDate) continue;
    calendar.push({
      dateId: safeInt(cal[i][K.Date_ID]),
      monthStart: msDate,
      msNum: dayNum(msDate),
      meNum: dayNum(normaliseDate(cal[i][K.Month_End])),
      daysInMonth: safeInt(cal[i][K.Days_In_Month]),
      label: monthLabel_(cal[i][K.Month_Label], msDate),
      inHorizon: safeBool(cal[i][K.In_Horizon])
    });
  }
  calendar.sort((a, b) => a.msNum - b.msNum);

  // ---- surcharge definitions ---------------------------------------------
  const dimSurcharge = {};
  const ds = getAllData_(SHEET.DIM_SURCHARGE), S = COL.DIM_SURCHARGE;
  for (let i = 1; i < ds.length; i++) {
    const code = safeStr(ds[i][S.Surcharge_Code]);
    if (!code || !safeBool(ds[i][S.Active])) continue;
    dimSurcharge[code] = {
      valueType:  safeStr(ds[i][S.Value_Type]) || 'AMT',
      appliesTo:  safeStr(ds[i][S.Applies_To]) || 'BASE',
      applyOrder: safeInt(ds[i][S.Apply_Order]),
      proration:  safeStr(ds[i][S.Proration]) || 'DAY_WEIGHTED'
    };
  }

  // ---- structure ----------------------------------------------------------
  const highLevelIds = [];
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id) continue;
    highLevelIds.push({
      id: id, code: safeStr(hl[i][H.High_Level_Code]),
      brand: safeStr(hl[i][H.Brand]), geo: safeStr(hl[i][H.Geo]),
      treatmentType: safeStr(hl[i][H.Treatment_Type]), wlSplit: safeStr(hl[i][H.WL_Split]),
      currency: safeStr(hl[i][H.Currency]), active: safeBool(hl[i][H.Active])
    });
  }

  const modellingIds = [];
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id) continue;
    modellingIds.push({
      id: id, hlId: safeInt(md[i][M.High_Level_ID]),
      carrier: safeStr(md[i][M.Carrier_Code]), method: safeStr(md[i][M.Method_Code]),
      letterParcel: safeStr(md[i][M.Letter_Parcel]), active: safeBool(md[i][M.Active])
    });
  }

  // ---- dated rows ---------------------------------------------------------
  const rateBase = readDated_(SHEET.RATE_BASE, COL.RATE_BASE, scenarioId,
    (r, C) => ({ modellingId: safeInt(r[C.Modelling_ID]), v: safeNum(r[C.Base_Rate]) }),
    COL.RATE_BASE.Valid_From, COL.RATE_BASE.Valid_To, COL.RATE_BASE.Active, COL.RATE_BASE.Scenario_ID);

  const rateSurcharge = readDated_(SHEET.RATE_SURCHARGE, COL.RATE_SURCHARGE, scenarioId,
    (r, C) => ({ modellingId: safeInt(r[C.Modelling_ID]),
                 code: safeStr(r[C.Surcharge_Code]), v: safeNum(r[C.Value]) }),
    COL.RATE_SURCHARGE.Valid_From, COL.RATE_SURCHARGE.Valid_To,
    COL.RATE_SURCHARGE.Active, COL.RATE_SURCHARGE.Scenario_ID);

  const mixMethod = readDated_(SHEET.MIX_METHOD, COL.MIX_METHOD, scenarioId,
    (r, C) => ({ modellingId: safeInt(r[C.Modelling_ID]),
                 regime: safeStr(r[C.Temp_Regime]), v: safeNum(r[C.Mix_Pct]) }),
    COL.MIX_METHOD.Valid_From, COL.MIX_METHOD.Valid_To,
    COL.MIX_METHOD.Active, COL.MIX_METHOD.Scenario_ID);

  const mixLetterParcel = readDated_(SHEET.MIX_LETTERPARCEL, COL.MIX_LETTERPARCEL, scenarioId,
    (r, C) => ({ hlId: safeInt(r[C.High_Level_ID]), v: safeNum(r[C.Letter_Mix_Pct]) }),
    COL.MIX_LETTERPARCEL.Valid_From, COL.MIX_LETTERPARCEL.Valid_To,
    COL.MIX_LETTERPARCEL.Active, COL.MIX_LETTERPARCEL.Scenario_ID);

  const mixColdChain = readDated_(SHEET.MIX_COLDCHAIN, COL.MIX_COLDCHAIN, scenarioId,
    (r, C) => ({ hlId: safeInt(r[C.High_Level_ID]), v: safeNum(r[C.CC_Mix_Pct]) }),
    COL.MIX_COLDCHAIN.Valid_From, COL.MIX_COLDCHAIN.Valid_To,
    COL.MIX_COLDCHAIN.Active, COL.MIX_COLDCHAIN.Scenario_ID);

  perfMark('loadEngineInput_ done');

  return {
    calendar: calendar, dimSurcharge: dimSurcharge,
    highLevelIds: highLevelIds, modellingIds: modellingIds,
    rateBase: rateBase, rateSurcharge: rateSurcharge,
    mixMethod: mixMethod, mixLetterParcel: mixLetterParcel, mixColdChain: mixColdChain,
    scenarioId: scenarioId
  };
}

/** Shared reader for the effective-dated tables. */
function readDated_(sheetName, C, scenarioId, build, fromCol, toCol, activeCol, scenarioCol) {
  const data = getAllData_(sheetName);
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (activeCol !== undefined && !safeBool(r[activeCol])) continue;
    if (scenarioCol !== undefined && safeInt(r[scenarioCol]) !== scenarioId) continue;
    const f = dayNum(r[fromCol]), t = dayNum(r[toCol]);
    if (f === null || t === null) continue;
    const o = build(r, C);
    o.f = f; o.t = t;
    out.push(o);
  }
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// CONVENIENCE — run the model and print a summary, without writing anything
// ─────────────────────────────────────────────────────────────────────────────

function runEnginePreview() {
  requireMaintenance_();
  perfReset();
  const input  = loadEngineInput_(1);
  const result = computeModel_(input);
  perfReport();

  Logger.log('=== ENGINE PREVIEW (nothing written) ===');
  Logger.log('  months            : ' + result.stats.months);
  Logger.log('  Modelling IDs     : ' + result.stats.modellingIds);
  Logger.log('  OUTPUT rows       : ' + result.stats.outputRows + ' (expected 612)');
  Logger.log('  OUTPUT_Detail rows: ' + result.stats.outputDetailRows + ' (expected 8352)');
  Logger.log('  compute time      : ' + result.stats.ms + 'ms');

  Logger.log('');
  if (!result.defects.length) {
    Logger.log('  no defects found');
  } else {
    const byRule = {};
    result.defects.forEach(d => { byRule[d.rule] = (byRule[d.rule] || 0) + 1; });
    Logger.log('  defects:');
    for (const rule in byRule) Logger.log('    ' + rule + ': ' + byRule[rule]);
    Logger.log('  first few:');
    result.defects.slice(0, 5).forEach(d => Logger.log('    ' + d.rule + ' — ' + d.label));
  }

  Logger.log('');
  Logger.log('  sample — High Level ID 1:');
  result.outputRows.filter(r => r.hlId === 1).slice(0, 6).forEach(r =>
    Logger.log('    ' + fmtDate(r.monthStart) + '  ' + r.forecastRatePerOrder.toFixed(6)));

  return result.stats;
}