/**
 * Postage Forecast Portal — mixes.gs
 *
 * Editing the delivery method mix, and the letter/parcel split.
 *
 * THE INVARIANT
 *
 * For every High Level ID, every month, cold-chain and ambient separately,
 * letters and parcels separately, the method percentages must total 100%.
 *
 * Nothing in the original workbook enforced this. Set a mix to 90% and the
 * forecast quietly came out 10% low, with no error anywhere.
 *
 * That is why mixes are saved as a GRID, not row by row. Saving one route at a
 * time cannot maintain the total — you would have to pass through an invalid
 * state to get from one valid one to another. The whole set is validated and
 * written together, or nothing is.
 */

// ─────────────────────────────────────────────────────────────────────────────
// READING THE GRID
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Every route in one High Level ID, with its mix for a regime and class,
 * as at a given date — plus the rate, so the page can show the consequence.
 *
 * @param {Object} p { hlId, regime: 'CC'|'AMBIENT', letterParcel: 'LETTER'|'PARCEL',
 *                     asAt: 'yyyy-mm-dd', scenarioId }
 */
function getMethodMixGrid(p) {
  const perms = requirePermissions_();
  const hlId = safeInt(p.hlId);
  if (!canSeeHighLevelId_(perms, hlId, false)) {
    throw new Error('High Level ID ' + hlId + ' is outside the area you can see.');
  }

  prewarmSheetCache_([SHEET.MODELLING_IDS, SHEET.MIX_METHOD, SHEET.DIM_CALENDAR,
                      SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.DIM_SURCHARGE,
                      SHEET.HIGH_LEVEL_IDS, SHEET.MIX_LETTERPARCEL, SHEET.MIX_COLDCHAIN]);

  const regime = safeStr(p.regime).toUpperCase() || 'AMBIENT';
  const lp     = safeStr(p.letterParcel).toUpperCase() || 'PARCEL';
  const scenarioId = safeInt(p.scenarioId) || 1;
  const asAt = normaliseDate(p.asAt) || monthStart(configDate('HORIZON_START', '2026-01-01'));
  const asAtKey = dateKey(asAt);

  // routes in this High Level ID of this class
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const routes = [];
  for (let i = 1; i < md.length; i++) {
    if (safeInt(md[i][M.High_Level_ID]) !== hlId) continue;
    if (safeStr(md[i][M.Letter_Parcel]).toUpperCase() !== lp) continue;
    if (!safeBool(md[i][M.Active])) continue;
    routes.push({ modellingId: safeInt(md[i][M.Modelling_ID]),
                  carrier: safeStr(md[i][M.Carrier_Code]),
                  method: safeStr(md[i][M.Method_Code]) });
  }

  // the mix in force on that date
  const mix = getAllData_(SHEET.MIX_METHOD), X = COL.MIX_METHOD;
  const byRoute = {};
  for (let i = 1; i < mix.length; i++) {
    if (!safeBool(mix[i][X.Active])) continue;
    if (safeInt(mix[i][X.Scenario_ID]) !== scenarioId) continue;
    if (safeStr(mix[i][X.Temp_Regime]).toUpperCase() !== regime) continue;
    const f = dateKey(mix[i][X.Valid_From]), t = dateKey(mix[i][X.Valid_To]);
    if (f > asAtKey || t < asAtKey) continue;
    byRoute[safeInt(mix[i][X.Modelling_ID])] = {
      id: safeInt(mix[i][X.Mix_ID]), value: safeNum(mix[i][X.Mix_Pct]),
      validFrom: fmtDate(mix[i][X.Valid_From]), validTo: fmtDate(mix[i][X.Valid_To])
    };
  }

  // rate per parcel for the same month, so the page can show contributions
  const rateByRoute = {};
  try {
    const input  = loadEngineInput_(scenarioId);
    const result = computeModel_(input);
    const wantMonth = dateKey(monthStart(asAt));
    result.outputDetailRows.forEach(r => {
      if (dateKey(r.monthStart) === wantMonth) rateByRoute[r.modellingId] = r.ratePerParcel;
    });
  } catch (e) { /* rates are a nicety here, not required */ }

  let total = 0;
  const rows = routes.map(r => {
    const m = byRoute[r.modellingId] || { id: 0, value: 0, validFrom: '', validTo: '' };
    total += m.value;
    return {
      modellingId: r.modellingId, carrier: r.carrier, method: r.method,
      mixId: m.id, value: m.value, validFrom: m.validFrom, validTo: m.validTo,
      ratePerParcel: rateByRoute[r.modellingId] || 0,
      contribution: (rateByRoute[r.modellingId] || 0) * m.value
    };
  });

  return {
    hlId: hlId, regime: regime, letterParcel: lp,
    asAt: fmtDate(asAt), scenarioId: scenarioId,
    rows: rows,
    total: total,
    isValid: Math.abs(total - 1) <= 0.0005,
    canEdit: canSeeHighLevelId_(perms, hlId, true)
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// SAVING THE GRID
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Replace the method mix for one High Level ID / regime / class over a period.
 *
 * Refuses unless the values total 100%. That refusal is the single most
 * important guard in the system: everything else can be recalculated, but a mix
 * that does not add up produces a plausible-looking forecast that is quietly
 * wrong.
 *
 * @param {Object} p { hlId, regime, letterParcel, validFrom, validTo,
 *                     rows: [{ modellingId, value }], scenarioId, notes }
 */
function saveMethodMixGrid(p) {
  prewarmForWrite_([SHEET.MIX_METHOD, SHEET.MIX_METHOD_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const hlId = safeInt(p.hlId);
  assertCanEditHighLevelId_(perms, hlId);

  const regime = safeStr(p.regime).toUpperCase();
  const lp     = safeStr(p.letterParcel).toUpperCase();
  if (regime !== 'CC' && regime !== 'AMBIENT') throw new Error('Regime must be CC or AMBIENT.');
  if (lp !== 'LETTER' && lp !== 'PARCEL')      throw new Error('Class must be LETTER or PARCEL.');

  const d = validateDates_(p.validFrom, p.validTo);
  const scenarioId = safeInt(p.scenarioId) || 1;
  const rows = p.rows || [];
  if (!rows.length) throw new Error('No rows to save.');

  // ---- the routes this grid is allowed to contain -------------------------
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const allowed = {};
  for (let i = 1; i < md.length; i++) {
    if (safeInt(md[i][M.High_Level_ID]) !== hlId) continue;
    if (safeStr(md[i][M.Letter_Parcel]).toUpperCase() !== lp) continue;
    if (!safeBool(md[i][M.Active])) continue;
    allowed[safeInt(md[i][M.Modelling_ID])] = true;
  }

  // ---- validate before touching anything ---------------------------------
  // Completeness is checked BEFORE the total. Leaving a route out also changes
  // the total, so checking the total first reports "85%" when the real problem
  // is a missing row — a refusal either way, but an unhelpful one.
  const supplied = {};
  rows.forEach(r => { supplied[safeInt(r.modellingId)] = true; });
  const missing = Object.keys(allowed).filter(id => !supplied[id]);
  if (missing.length) {
    throw new Error('These routes are missing from the grid: ' + missing.join(', ') +
                    '. Include them with 0% rather than leaving them out.');
  }

  let total = 0;
  rows.forEach(r => {
    const mid = safeInt(r.modellingId);
    if (!allowed[mid]) {
      throw new Error('Modelling ID ' + mid + ' does not belong to High Level ID ' +
                      hlId + ' as a ' + lp.toLowerCase() + ' route.');
    }
    const v = safeNum(r.value);
    if (v < 0 || v > 1) {
      throw new Error('Modelling ID ' + mid + ' has a mix of ' + (v * 100).toFixed(2) +
                      '% — it must be between 0 and 100.');
    }
    total += v;
  });

  if (Math.abs(total - 1) > 0.0005) {
    throw new Error('The percentages total ' + (total * 100).toFixed(2) +
                    '%. They must total 100% before this can be saved.');
  }

  // ---- write --------------------------------------------------------------
  const t = TABLES.MIX_METHOD, C = COL.MIX_METHOD, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    const cleared = clearMixWindowBatch_(Object.keys(allowed).map(safeInt),
                                        regime, scenarioId, d.from, d.to, perms);

    invalidateSheetCache_(t.sheet);
    const sh = getSheet_(t.sheet);
    const newRows = [];
    rows.forEach(r => {
      const row = blankRow_('MIX_METHOD');
      row[C.Mix_ID]       = getNextId_(t.sheet, C.Mix_ID);
      row[C.Modelling_ID] = safeInt(r.modellingId);
      row[C.Temp_Regime]  = regime;
      row[C.Valid_From]   = d.from;
      row[C.Valid_To]     = d.to;
      row[C.Mix_Pct]      = safeNum(r.value);
      row[C.Scenario_ID]  = scenarioId;
      row[C.Active]       = true;
      row[C.Notes]        = safeStr(p.notes);
      row[C.Created_TS]   = new Date();
      row[C.Created_By]   = perms.email;
      row[C.Updated_TS]   = new Date();
      row[C.Updated_By]   = perms.email;
      newRows.push(row);
    });

    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, width).setValues(newRows);
    invalidateSheetCache_(t.sheet);

    recordChangesBatch_('MIX_METHOD', newRows.map(row =>
      ({ id: row[C.Mix_ID], before: null, after: row, type: 'CREATE' })));
    logAudit_('UPDATE', 'MIX_GRID', hlId + '/' + regime + '/' + lp, '', '',
              newRows.length + ' routes',
              'total 100% for ' + fmtDate(d.from) + ' to ' + fmtDate(d.to), true);

    return { ok: true, written: newRows.length, cleared: cleared, total: total };
  });
}


/**
 * Make room for a new period, across every route in the grid at once.
 *
 *   overlapping, starts earlier          -> shortened to end the day before
 *   overlapping, entirely inside         -> deactivated
 *   overlapping, continues past the end  -> start moved to the day after
 *   spans the whole window               -> shortened to end the day before
 *   no overlap                           -> untouched
 *
 * The rows needing adjustment are scattered, so they cannot be written as one
 * contiguous range. Instead the whole data range is read once, changed in
 * memory, and written back once. Row by row this meant a setValues plus two
 * appendRow calls each — 66 writes to clear 22 rows, about 28 seconds. This is
 * four writes however many rows are touched.
 */
function clearMixWindowBatch_(modellingIds, regime, scenarioId, from, to, perms) {
  const t = TABLES.MIX_METHOD, C = COL.MIX_METHOD, width = t.headers.length;
  const sh = getSheet_(t.sheet);
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const want = {};
  (modellingIds || []).forEach(id => { want[safeInt(id)] = true; });

  const fromKey = dateKey(from), toKey = dateKey(to);
  const dayBefore = new Date(from.getFullYear(), from.getMonth(), from.getDate() - 1);
  const dayAfter  = new Date(to.getFullYear(),   to.getMonth(),   to.getDate() + 1);

  const data = sh.getRange(2, 1, last - 1, width).getValues();
  const changes = [];
  const now = new Date();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!want[safeInt(r[C.Modelling_ID])]) continue;
    if (safeStr(r[C.Temp_Regime]).toUpperCase() !== regime) continue;
    if (safeInt(r[C.Scenario_ID]) !== scenarioId) continue;
    if (!safeBool(r[C.Active])) continue;

    const f = dateKey(r[C.Valid_From]), tt = dateKey(r[C.Valid_To]);
    if (f > toKey || tt < fromKey) continue;          // no overlap

    const before = r.slice();

    if (f < fromKey) {
      // starts earlier, or spans the whole window: keep the earlier part.
      // Splitting a spanning row in two would be tidier, but silently creating
      // a row nobody asked for is worse than a visible gap they can fill.
      r[C.Valid_To] = dayBefore;
    } else if (tt > toKey) {
      r[C.Valid_From] = dayAfter;
    } else {
      r[C.Active] = false;
    }

    r[C.Updated_TS] = now;
    r[C.Updated_By] = perms.email;
    changes.push({ id: safeInt(before[C.Mix_ID]), before: before,
                   after: r.slice(), type: 'UPDATE' });
  }

  if (changes.length) {
    sh.getRange(2, 1, data.length, width).setValues(data);
    invalidateSheetCache_(t.sheet);
    recordChangesBatch_('MIX_METHOD', changes);
  }
  return changes.length;
}


/**
 * Scale entered values so they total exactly 100%.
 *
 * Rounding each value independently does not get there: 0.5, 0.3, 0.3 scaled by
 * 1.1 and rounded to six places sums to 99.9999%. So every row but the largest
 * is rounded, and the largest absorbs whatever is left over. The largest is
 * chosen because a residual of a millionth is least visible there.
 */
function normaliseMixRows(rows) {
  requirePermissions_();
  const list = (rows || []).map(r => ({
    modellingId: safeInt(r.modellingId), raw: safeNum(r.value)
  }));
  const total = list.reduce((s, r) => s + r.raw, 0);
  if (!total) throw new Error('All values are zero — there is nothing to scale.');

  let biggest = 0;
  for (let i = 1; i < list.length; i++) if (list[i].raw > list[biggest].raw) biggest = i;

  let allocated = 0;
  const out = list.map((r, i) => {
    if (i === biggest) return { modellingId: r.modellingId, value: 0 };
    const v = Math.round((r.raw / total) * 1000000) / 1000000;
    allocated += v;
    return { modellingId: r.modellingId, value: v };
  });
  out[biggest].value = Math.round((1 - allocated) * 1000000) / 1000000;
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// BULK: ONE METHOD MIX ACROSS MANY SEGMENTS
//
// saveMethodMixGrid writes one grid: one High Level ID, one regime, one class.
// This writes the same mix into many of them at once, chosen by dimension the
// way a bulk rate change is.
//
// A mix is not a rate. A rate is one independent number, so a bulk rate change
// can set it wherever it finds it. A mix is a SET that has to total 100%, and
// each segment carries its own routes — so the same percentages do not
// necessarily fit every segment they are pointed at.
//
// The rule here is exact application. A route the mix does not name is set to
// 0%; that is safe, because the named ones still total 100%. But a route the
// mix names with a NON-ZERO share, in a segment that does not carry it, cannot
// be applied: those percentages have nowhere to go and the grid would land
// under 100%. Such a grid is REFUSED and named in the preview, never quietly
// renormalised to fit. Renormalising would be the exact failure the 100% guard
// exists to prevent — a plausible-looking forecast that is silently not the
// mix anybody asked for.
// ─────────────────────────────────────────────────────────────────────────────

/** A route's identity across segments: the pair the grid actually shows. */
function mixRouteKey_(carrier, method) {
  return normKey(carrier) + '|' + normKey(method);
}

/**
 * Every grid a dimension selection matches, and the routes inside them.
 *
 * Brand, geo, treatment type and WL split select High Level IDs; regime and
 * class multiply them out into grids, because a grid IS one segment-regime-class.
 * Routes belong to a segment and a class, not to a regime, so CC and AMBIENT for
 * the same segment and class share a route list and differ only in their values.
 *
 * Dimension values reuse dimensionMatches_ from rates.gs, so a tick-box list
 * means the same thing on both forms — the union of what is ticked, ANDed across
 * dimensions, with nothing ticked meaning All.
 */
function resolveMixTargets_(dims, perms) {
  const sel = dims || {};
  const regimes = pickFromList_(sel.regime, ['CC', 'AMBIENT']);
  const classes = pickFromList_(sel.letterParcel, ['LETTER', 'PARCEL']);

  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  const hlOk = {};
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id || !safeBool(hl[i][H.Active])) continue;
    if (!dimensionMatches_(sel.brand,         hl[i][H.Brand]))          continue;
    if (!dimensionMatches_(sel.geo,           hl[i][H.Geo]))            continue;
    if (!dimensionMatches_(sel.treatmentType, hl[i][H.Treatment_Type])) continue;
    if (!dimensionMatches_(sel.wlSplit,       hl[i][H.WL_Split]))       continue;
    hlOk[id] = true;
  }

  // routes, grouped by segment and class
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const bySegClass = {}, seenSeg = {};
  let skippedNoScope = 0;
  const noScopeSeen = {};
  for (let i = 1; i < md.length; i++) {
    const mid = safeInt(md[i][M.Modelling_ID]);
    if (!mid || !safeBool(md[i][M.Active])) continue;
    const hlId = safeInt(md[i][M.High_Level_ID]);
    if (!hlOk[hlId]) continue;
    const lp = normKey(md[i][M.Letter_Parcel]);
    if (classes.indexOf(lp) === -1) continue;

    // Scope is checked per segment, not per route: a grid is a segment's grid,
    // and half of one is not a thing that can total 100%.
    if (!canSeeHighLevelId_(perms, hlId, true)) {
      if (!noScopeSeen[hlId]) { noScopeSeen[hlId] = true; skippedNoScope++; }
      continue;
    }

    const k = hlId + '|' + lp;
    if (!bySegClass[k]) bySegClass[k] = { hlId: hlId, letterParcel: lp, routes: [] };
    bySegClass[k].routes.push({ modellingId: mid,
                                carrier: safeStr(md[i][M.Carrier_Code]),
                                method: safeStr(md[i][M.Method_Code]) });
    seenSeg[hlId] = true;
  }

  // multiply out by regime
  const grids = [], routeTally = {};
  Object.keys(bySegClass).forEach(function (k) {
    const sc = bySegClass[k];
    regimes.forEach(function (regime) {
      grids.push({ hlId: sc.hlId, letterParcel: sc.letterParcel, regime: regime,
                   routes: sc.routes });
    });
    sc.routes.forEach(function (r) {
      const key = mixRouteKey_(r.carrier, r.method);
      if (!routeTally[key]) {
        routeTally[key] = { key: key, carrier: r.carrier, method: r.method, grids: 0 };
      }
      routeTally[key].grids += regimes.length;
    });
  });

  const routes = Object.keys(routeTally).map(function (k) { return routeTally[k]; });
  routes.sort(function (a, b) {
    return a.carrier === b.carrier ? (a.method < b.method ? -1 : a.method > b.method ? 1 : 0)
                                   : (a.carrier < b.carrier ? -1 : 1);
  });

  return { grids: grids, routes: routes, regimes: regimes, classes: classes,
           segmentCount: Object.keys(seenSeg).length, skippedNoScope: skippedNoScope };
}

/**
 * A locked-list dimension resolved against the codes the engine branches on.
 *
 * Regime and class are not free reference data: engine.gs reads CC/AMBIENT and
 * LETTER/PARCEL literally, which is why LOCKED_REFERENCE_LISTS exists. Anything
 * else that arrives is dropped rather than matched, so a typo narrows nothing
 * instead of silently selecting nothing.
 */
function pickFromList_(selected, allowed) {
  const wanted = dimensionWanted_(selected);
  if (!wanted.length) return allowed.slice();
  const out = allowed.filter(function (v) { return wanted.indexOf(v) > -1; });
  return out.length ? out : allowed.slice();
}

/** A human description of a bulk mix selection, for the preview and the audit. */
function describeMixDimensions_(dims) {
  const sel = dims || {};
  const parts = [];
  [['brand', 'brand'], ['geo', 'geo'], ['treatmentType', 'treatment'],
   ['wlSplit', 'WL split'], ['regime', 'regime'],
   ['letterParcel', 'class']].forEach(function (pair) {
    const wanted = dimensionWanted_(sel[pair[0]]);
    parts.push(pair[1] + '=' + (wanted.length ? wanted.slice().sort().join('+') : 'All'));
  });
  return parts.join(', ');
}

/**
 * The grids a selection matches and the routes to put percentages against.
 *
 * Called before the value form is built, because the useful question is not
 * "which of the 35 methods in Dim_Method" but "which routes do the segments you
 * just chose actually carry" — usually a handful.
 *
 * @param {Object} p { dimensions }
 */
function bulkMixScope(p) {
  prewarmSheetCache_([SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const t = resolveMixTargets_(p.dimensions, perms);
  return {
    grids: t.grids.length, segments: t.segmentCount,
    skippedNoScope: t.skippedNoScope,
    regimes: t.regimes, classes: t.classes,
    routes: t.routes,
    selection: describeMixDimensions_(p.dimensions)
  };
}

/**
 * Plan the write: which grids take the mix exactly, and which cannot.
 *
 * The entered mix is checked for 100% ONCE, here, rather than per grid. It is a
 * property of what was typed, not of any segment, and reporting it per grid
 * would bury one typo under a hundred identical complaints.
 */
function planBulkMix_(p, perms) {
  const d = validateDates_(p.validFrom, p.validTo);
  const scenarioId = safeInt(p.scenarioId) || 1;

  const entered = {};
  let total = 0;
  (p.mix || []).forEach(function (m) {
    const key = mixRouteKey_(m.carrier, m.method);
    const v = safeNum(m.value);
    if (v < 0 || v > 1) {
      throw new Error(safeStr(m.carrier) + ' ' + safeStr(m.method) + ' is ' +
        (v * 100).toFixed(2) + '% — it must be between 0 and 100.');
    }
    entered[key] = (entered[key] || 0) + v;
    total += v;
  });
  if (!Object.keys(entered).length) throw new Error('No mix was given.');
  if (Math.abs(total - 1) > 0.0005) {
    throw new Error('The percentages total ' + (total * 100).toFixed(2) +
      '%. They must total 100% before this can be applied to anything.');
  }

  const named = Object.keys(entered).filter(function (k) { return entered[k] > 0; });
  const t = resolveMixTargets_(p.dimensions, perms);

  const mix = getAllData_(SHEET.MIX_METHOD), X = COL.MIX_METHOD;
  const fromKey = dateKey(d.from);
  const current = {};                       // modellingId|regime -> value in force
  for (let i = 1; i < mix.length; i++) {
    if (!safeBool(mix[i][X.Active])) continue;
    if (safeInt(mix[i][X.Scenario_ID]) !== scenarioId) continue;
    const f = dateKey(mix[i][X.Valid_From]), tt = dateKey(mix[i][X.Valid_To]);
    if (f > fromKey || tt < fromKey) continue;
    current[safeInt(mix[i][X.Modelling_ID]) + '|' + normKey(mix[i][X.Temp_Regime])] =
      safeNum(mix[i][X.Mix_Pct]);
  }

  const planned = [], refused = [];
  t.grids.forEach(function (g) {
    const have = {};
    g.routes.forEach(function (r) { have[mixRouteKey_(r.carrier, r.method)] = true; });

    const absent = named.filter(function (k) { return !have[k]; });
    if (absent.length) {
      refused.push({ hlId: g.hlId, regime: g.regime, letterParcel: g.letterParcel,
                     missing: absent.map(function (k) { return k.replace('|', ' '); }) });
      return;
    }

    let changed = false;
    const rows = g.routes.map(function (r) {
      const v = entered[mixRouteKey_(r.carrier, r.method)] || 0;
      const was = current[r.modellingId + '|' + g.regime];
      if (was === undefined || Math.abs(was - v) > 0.0000005) changed = true;
      return { modellingId: r.modellingId, carrier: r.carrier, method: r.method,
               value: v, was: was === undefined ? null : was };
    });
    planned.push({ hlId: g.hlId, regime: g.regime, letterParcel: g.letterParcel,
                   rows: rows, changed: changed });
  });

  return { dates: d, scenarioId: scenarioId, entered: entered,
           planned: planned, refused: refused,
           routes: t.routes, segments: t.segmentCount,
           skippedNoScope: t.skippedNoScope,
           selection: describeMixDimensions_(p.dimensions),
           planKey: mixPlanKey_(planned) };
}

/**
 * A fingerprint of the planned grids and the values they currently hold.
 *
 * Same purpose as planKey_ in rates.gs: the client previews, the user confirms,
 * and the apply call sends this back. If a mix moved in between, the recomputed
 * key differs and the write is refused rather than applied to a set nobody saw.
 */
function mixPlanKey_(planned) {
  const parts = [];
  planned.forEach(function (g) {
    g.rows.forEach(function (r) {
      parts.push(g.regime + ':' + r.modellingId + ':' +
                 Math.round((r.was === null ? -1 : r.was) * 1000000) + ':' +
                 Math.round(r.value * 1000000));
    });
  });
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
                                        parts.sort().join('|') || 'empty');
  return bytes.map(function (b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('').slice(0, 12);
}

/**
 * Apply one method mix to every grid a dimension selection matches.
 *
 * Previews by default. Pass preview:false with the planKey the preview returned
 * to write.
 *
 * @param {Object} p { dimensions: { brand, geo, treatmentType, wlSplit,
 *                                   regime, letterParcel },
 *                     mix: [{ carrier, method, value }],   // value 0..1, totals 1
 *                     validFrom, validTo, scenarioId, notes, preview, planKey }
 */
function bulkUpdateMethodMix(p) {
  prewarmForWrite_([SHEET.MIX_METHOD, SHEET.MIX_METHOD_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const plan = planBulkMix_(p, perms);
  const changing = plan.planned.filter(function (g) { return g.changed; });

  if (p.preview !== false) {
    return {
      preview: true, selection: plan.selection,
      validFrom: fmtDate(plan.dates.from), validTo: fmtDate(plan.dates.to),
      grids: plan.planned.length, changing: changing.length,
      unchanged: plan.planned.length - changing.length,
      segments: plan.segments, skippedNoScope: plan.skippedNoScope,
      rowsToWrite: plan.planned.reduce(function (n, g) { return n + g.rows.length; }, 0),
      refused: plan.refused.map(function (r) {
        return { label: r.hlId + ' · ' + r.regime + ' · ' + r.letterParcel,
                 missing: r.missing.join(', ') };
      }),
      sample: plan.planned.slice(0, 8).map(function (g) {
        return { label: g.hlId + ' · ' + g.regime + ' · ' + g.letterParcel,
                 changed: g.changed,
                 rows: g.rows.map(function (r) {
                   return { route: r.carrier + ' ' + r.method,
                            was: r.was, now: r.value };
                 }) };
      }),
      planKey: plan.planKey
    };
  }

  if (!plan.planned.length) {
    throw new Error('Nothing matches that selection, so there is nothing to update. ' +
      plan.refused.length + ' grid(s) do not carry every method you gave a share to, ' +
      'and ' + plan.skippedNoScope + ' segment(s) are outside what you can edit.');
  }
  if (safeStr(p.planKey) !== plan.planKey) {
    throw new Error('These mixes have changed since the preview was taken, so ' +
      'nothing was applied. Run the preview again and check it still says what ' +
      'you expect.');
  }

  return withLock_(function () { return applyBulkMix_(p, plan, perms); });
}

/**
 * Write a planned bulk mix. Assumes the caller holds the lock.
 *
 * Every grid for one regime is cleared in a single pass, because
 * clearMixWindowBatch_ reads and rewrites the whole Mix_Method sheet each time
 * it is called — once per regime is two passes however many segments are in
 * play, where once per grid would be two per segment. The new rows then go down
 * as one setValues rather than one per grid.
 */
function applyBulkMix_(p, plan, perms) {
  const t = TABLES.MIX_METHOD, C = COL.MIX_METHOD, width = t.headers.length;
  const now = new Date();
  const batchRef = bulkBatchRef_(now);

  invalidateSheetCache_(t.sheet);

  const byRegime = {};
  plan.planned.forEach(function (g) {
    if (!byRegime[g.regime]) byRegime[g.regime] = [];
    g.rows.forEach(function (r) { byRegime[g.regime].push(r.modellingId); });
  });

  let cleared = 0;
  Object.keys(byRegime).forEach(function (regime) {
    cleared += clearMixWindowBatch_(byRegime[regime], regime, plan.scenarioId,
                                    plan.dates.from, plan.dates.to, perms);
  });

  invalidateSheetCache_(t.sheet);
  const sh = getSheet_(t.sheet);
  const newRows = [];
  plan.planned.forEach(function (g) {
    g.rows.forEach(function (r) {
      const row = blankRow_('MIX_METHOD');
      row[C.Mix_ID]       = getNextId_(t.sheet, C.Mix_ID);
      row[C.Modelling_ID] = r.modellingId;
      row[C.Temp_Regime]  = g.regime;
      row[C.Valid_From]   = plan.dates.from;
      row[C.Valid_To]     = plan.dates.to;
      row[C.Mix_Pct]      = r.value;
      row[C.Scenario_ID]  = plan.scenarioId;
      row[C.Active]       = true;
      row[C.Notes]        = safeStr(p.notes) ? (batchRef + ' ' + safeStr(p.notes)) : batchRef;
      row[C.Created_TS]   = now;
      row[C.Created_By]   = perms.email;
      row[C.Updated_TS]   = now;
      row[C.Updated_By]   = perms.email;
      newRows.push(row);
    });
  });

  sh.getRange(sh.getLastRow() + 1, 1, newRows.length, width).setValues(newRows);
  invalidateSheetCache_(t.sheet);

  recordChangesBatch_('MIX_METHOD', newRows.map(function (row) {
    return { id: row[C.Mix_ID], before: null, after: row, type: 'CREATE' };
  }));
  logAudit_('UPDATE', 'MIX_GRID', batchRef, '', '',
            plan.planned.length + ' grids, ' + newRows.length + ' routes',
            batchRef + ' bulk mix from ' + fmtDate(plan.dates.from) + ' — ' + plan.selection,
            true);

  return { ok: true, batchRef: batchRef, grids: plan.planned.length,
           written: newRows.length, cleared: cleared,
           refused: plan.refused.length };
}


// ─────────────────────────────────────────────────────────────────────────────
// LETTER / PARCEL SPLIT
// ─────────────────────────────────────────────────────────────────────────────

/**
 * The letter / parcel split for one segment as a two-row grid, with the blended
 * rate each class produces.
 *
 * The two levels of mix are easy to confuse, so this makes the relationship
 * visible: the split decides how many orders go each way, and the method mix
 * decides which courier carries them once they are on that side.
 *
 * @param {Object} p { hlId, asAt, scenarioId }
 */
function getLetterParcelGrid(p) {
  const perms = requirePermissions_();
  const hlId = safeInt(p.hlId);
  if (!canSeeHighLevelId_(perms, hlId, false)) {
    throw new Error('High Level ID ' + hlId + ' is outside the area you can see.');
  }

  prewarmSheetCache_([SHEET.MODELLING_IDS, SHEET.MIX_LETTERPARCEL, SHEET.MIX_METHOD,
                      SHEET.DIM_CALENDAR, SHEET.RATE_BASE, SHEET.RATE_SURCHARGE,
                      SHEET.DIM_SURCHARGE, SHEET.HIGH_LEVEL_IDS, SHEET.MIX_COLDCHAIN]);

  const scenarioId = safeInt(p.scenarioId) || 1;
  const asAt = normaliseDate(p.asAt) || monthStart(configDate('HORIZON_START', '2026-01-01'));
  const asAtKey = dateKey(asAt);

  // how many routes serve each class
  const counts = { LETTER: 0, PARCEL: 0 };
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < md.length; i++) {
    if (safeInt(md[i][M.High_Level_ID]) !== hlId) continue;
    if (!safeBool(md[i][M.Active])) continue;
    const cls = safeStr(md[i][M.Letter_Parcel]).toUpperCase();
    if (counts[cls] !== undefined) counts[cls]++;
  }

  // the split in force on that date
  let letterShare = 0, currentId = 0, from = '', to = '';
  const lp = getAllData_(SHEET.MIX_LETTERPARCEL), L = COL.MIX_LETTERPARCEL;
  for (let i = 1; i < lp.length; i++) {
    if (safeInt(lp[i][L.High_Level_ID]) !== hlId) continue;
    if (!safeBool(lp[i][L.Active])) continue;
    if (safeInt(lp[i][L.Scenario_ID]) !== scenarioId) continue;
    const f = dateKey(lp[i][L.Valid_From]), t = dateKey(lp[i][L.Valid_To]);
    if (f > asAtKey || t < asAtKey) continue;
    letterShare = safeNum(lp[i][L.Letter_Mix_Pct]);
    currentId = safeInt(lp[i][L.LP_Mix_ID]);
    from = fmtDate(lp[i][L.Valid_From]);
    to   = fmtDate(lp[i][L.Valid_To]);
    break;
  }

  // what each class costs once its own method mix is applied
  const blended = { LETTER: 0, PARCEL: 0 };
  try {
    const input  = loadEngineInput_(scenarioId);
    const result = computeModel_(input);
    const wantMonth = dateKey(monthStart(asAt));
    result.outputDetailRows.forEach(r => {
      if (r.hlId !== hlId) return;
      if (dateKey(r.monthStart) !== wantMonth) return;
      const cls = safeStr(r.letterParcel).toUpperCase();
      if (blended[cls] === undefined) return;
      blended[cls] += r.ratePerParcel * r.methodMix;
    });
  } catch (e) { /* rates are context here, not required */ }

  const rows = [
    { cls: 'LETTER', routes: counts.LETTER, share: letterShare,
      blendedRate: blended.LETTER, contribution: blended.LETTER * letterShare },
    { cls: 'PARCEL', routes: counts.PARCEL, share: 1 - letterShare,
      blendedRate: blended.PARCEL, contribution: blended.PARCEL * (1 - letterShare) }
  ];

  // every period on record, so the timeline is visible
  const periods = [];
  for (let i = 1; i < lp.length; i++) {
    if (safeInt(lp[i][L.High_Level_ID]) !== hlId) continue;
    if (!safeBool(lp[i][L.Active])) continue;
    periods.push({ id: safeInt(lp[i][L.LP_Mix_ID]),
                   validFrom: fmtDate(lp[i][L.Valid_From]),
                   validTo: fmtDate(lp[i][L.Valid_To]),
                   letter: safeNum(lp[i][L.Letter_Mix_Pct]) });
  }
  periods.sort((a, b) => a.validFrom < b.validFrom ? -1 : 1);

  return {
    hlId: hlId, asAt: fmtDate(asAt), scenarioId: scenarioId,
    currentId: currentId, validFrom: from, validTo: to,
    rows: rows, periods: periods,
    total: rows[0].share + rows[1].share,
    blendedPerOrder: rows[0].contribution + rows[1].contribution,
    hasLetterRoutes: counts.LETTER > 0,
    canEdit: canSeeHighLevelId_(perms, hlId, true)
  };
}


/**
 * Replace the letter / parcel split over a period.
 *
 * Only the letter share is stored — parcel is whatever is left, so the two can
 * never disagree. Existing periods that overlap are adjusted the same way the
 * method mix grid adjusts them, rather than the save being refused.
 *
 * @param {Object} p { hlId, validFrom, validTo, letterShare, scenarioId, notes }
 */
function saveLetterParcelGrid(p) {
  prewarmForWrite_([SHEET.MIX_LETTERPARCEL, SHEET.MIX_LETTERPARCEL_AMENDS,
                    SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const hlId = safeInt(p.hlId);
  assertCanEditHighLevelId_(perms, hlId);

  const d = validateDates_(p.validFrom, p.validTo);
  const letter = safeNum(p.letterShare);
  if (letter < 0 || letter > 1) {
    throw new Error('The letter share is ' + (letter * 100).toFixed(2) +
                    '%. It must be between 0 and 100 — parcel takes the remainder.');
  }

  // A letter share above zero with no letter routes would send orders nowhere.
  if (letter > 0) {
    const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
    let letterRoutes = 0;
    for (let i = 1; i < md.length; i++) {
      if (safeInt(md[i][M.High_Level_ID]) !== hlId) continue;
      if (!safeBool(md[i][M.Active])) continue;
      if (safeStr(md[i][M.Letter_Parcel]).toUpperCase() === 'LETTER') letterRoutes++;
    }
    if (!letterRoutes) {
      throw new Error('High Level ID ' + hlId + ' has no letter routes, so a letter share ' +
        'of ' + (letter * 100).toFixed(2) + '% would price at zero. Add letter routes ' +
        'under Admin first, or leave the share at 0%.');
    }
  }

  const t = TABLES.MIX_LETTERPARCEL, C = COL.MIX_LETTERPARCEL, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const cleared = clearLetterParcelWindow_(hlId, scenarioIdOf_(p), d.from, d.to, perms);

    invalidateSheetCache_(t.sheet);
    const sh = getSheet_(t.sheet);
    const row = blankRow_('MIX_LETTERPARCEL');
    row[C.LP_Mix_ID]      = getNextId_(t.sheet, C.LP_Mix_ID);
    row[C.High_Level_ID]  = hlId;
    row[C.Valid_From]     = d.from;
    row[C.Valid_To]       = d.to;
    row[C.Letter_Mix_Pct] = letter;
    row[C.Scenario_ID]    = scenarioIdOf_(p);
    row[C.Active]         = true;
    row[C.Notes]          = safeStr(p.notes);
    row[C.Created_TS]     = new Date();
    row[C.Created_By]     = perms.email;
    row[C.Updated_TS]     = new Date();
    row[C.Updated_By]     = perms.email;

    sh.getRange(sh.getLastRow() + 1, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);
    recordChange_('MIX_LETTERPARCEL', row[C.LP_Mix_ID], null, row, 'CREATE');
    logAudit_('UPDATE', 'LP_SPLIT', hlId, '', '',
              (letter * 100).toFixed(2) + '% letter',
              fmtDate(d.from) + ' to ' + fmtDate(d.to), true);

    return { ok: true, id: row[C.LP_Mix_ID], cleared: cleared,
             letter: letter, parcel: 1 - letter };
  });
}

function scenarioIdOf_(p) { return safeInt(p.scenarioId) || 1; }

/** Same window logic as the method mix, for one segment's split. */
function clearLetterParcelWindow_(hlId, scenarioId, from, to, perms) {
  const t = TABLES.MIX_LETTERPARCEL, C = COL.MIX_LETTERPARCEL, width = t.headers.length;
  const sh = getSheet_(t.sheet);
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const fromKey = dateKey(from), toKey = dateKey(to);
  const dayBefore = new Date(from.getFullYear(), from.getMonth(), from.getDate() - 1);
  const dayAfter  = new Date(to.getFullYear(),   to.getMonth(),   to.getDate() + 1);

  const data = sh.getRange(2, 1, last - 1, width).getValues();
  const changes = [];
  const now = new Date();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (safeInt(r[C.High_Level_ID]) !== hlId) continue;
    if (safeInt(r[C.Scenario_ID]) !== scenarioId) continue;
    if (!safeBool(r[C.Active])) continue;

    const f = dateKey(r[C.Valid_From]), tt = dateKey(r[C.Valid_To]);
    if (f > toKey || tt < fromKey) continue;

    const before = r.slice();
    if (f < fromKey)       r[C.Valid_To]   = dayBefore;
    else if (tt > toKey)   r[C.Valid_From] = dayAfter;
    else                   r[C.Active]     = false;

    r[C.Updated_TS] = now;
    r[C.Updated_By] = perms.email;
    changes.push({ id: safeInt(before[C.LP_Mix_ID]), before: before,
                   after: r.slice(), type: 'UPDATE' });
  }

  if (changes.length) {
    sh.getRange(2, 1, data.length, width).setValues(data);
    invalidateSheetCache_(t.sheet);
    recordChangesBatch_('MIX_LETTERPARCEL', changes);
  }
  return changes.length;
}


function saveLetterParcelMix(p) {
  prewarmForWrite_([SHEET.MIX_LETTERPARCEL, SHEET.MIX_LETTERPARCEL_AMENDS,
                    SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const hlId = safeInt(p.hlId);
  assertCanEditHighLevelId_(perms, hlId);

  const d = validateDates_(p.validFrom, p.validTo);
  const value = safeNum(p.value);
  if (value < 0 || value > 1) {
    throw new Error('The letter share must be between 0 and 100%. The parcel share is ' +
                    'whatever is left, so you only enter one of them.');
  }

  const t = TABLES.MIX_LETTERPARCEL, C = COL.MIX_LETTERPARCEL, width = t.headers.length;
  const isNew = !safeInt(p.id);
  const scenarioId = safeInt(p.scenarioId) || 1;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    // The segment the row belongs to now, not only the one it is being pointed
    // at — otherwise an editable hlId in the payload authorises any row's id.
    assertCanEditRowOwner_(perms, t.sheet, C.LP_Mix_ID, p.id,
                           C.High_Level_ID, 'HIGH_LEVEL_ID');

    assertNoOverlap_(t.sheet, C,
      r => safeInt(r[C.High_Level_ID]) === hlId && safeInt(r[C.Scenario_ID]) === scenarioId,
      d.fromKey, d.toKey, p.id, C.LP_Mix_ID, 'High Level ID ' + hlId);

    const sh = getSheet_(t.sheet);
    let rowIndex, before = null, row;

    if (isNew) {
      row = blankRow_('MIX_LETTERPARCEL');
      row[C.LP_Mix_ID]  = getNextId_(t.sheet, C.LP_Mix_ID);
      row[C.Created_TS] = new Date();
      row[C.Created_By] = perms.email;
      rowIndex = sh.getLastRow() + 1;
    } else {
      rowIndex = findRowById_(t.sheet, C.LP_Mix_ID, p.id);
      if (rowIndex < 0) throw new Error('That letter/parcel split no longer exists.');
      before = readRow_(t.sheet, rowIndex, width);
      row = before.slice();
    }

    row[C.High_Level_ID]  = hlId;
    row[C.Valid_From]     = d.from;
    row[C.Valid_To]       = d.to;
    row[C.Letter_Mix_Pct] = value;
    row[C.Scenario_ID]    = scenarioId;
    row[C.Active]         = true;
    row[C.Notes]          = safeStr(p.notes);
    row[C.Updated_TS]     = new Date();
    row[C.Updated_By]     = perms.email;

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);
    recordChange_('MIX_LETTERPARCEL', row[C.LP_Mix_ID], before, row, isNew ? 'CREATE' : 'UPDATE');
    return { ok: true, id: row[C.LP_Mix_ID], isNew: isNew };
  });
}


function deleteLetterParcelMix(id) {
  prewarmForWrite_([SHEET.MIX_LETTERPARCEL, SHEET.MIX_LETTERPARCEL_AMENDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditMixes_(perms);

  const t = TABLES.MIX_LETTERPARCEL, C = COL.MIX_LETTERPARCEL, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.LP_Mix_ID, id);
    if (rowIndex < 0) throw new Error('That letter/parcel split no longer exists.');

    const before = readRow_(t.sheet, rowIndex, width);
    assertCanEditHighLevelId_(perms, safeInt(before[C.High_Level_ID]));

    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('MIX_LETTERPARCEL', safeInt(before[C.LP_Mix_ID]), before, row, 'DELETE');
    return { ok: true };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Read-only check of the bulk mix planner, against whatever is really in the
 * sheet. Writes nothing — it only ever asks for previews, so it is safe to run
 * against production data.
 *
 * What it is actually testing is the rule that a grid missing a route with a
 * non-zero share is refused rather than renormalised. That rule is the whole
 * reason this feature needs a preview at all.
 */
function testBulkMixUpdate() {
  requireMaintenance_();
  const perms = requirePermissions_();
  requireEditMixes_(perms);
  prewarmSheetCache_([SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS, SHEET.MIX_METHOD]);

  const checks = [];
  function check(label, cond) {
    checks.push({ label: label, ok: !!cond });
    Logger.log((cond ? '  PASS  ' : '  FAIL  ') + label);
  }

  Logger.log('=== bulk mix update — read-only ===');

  const all = bulkMixScope({ dimensions: {} });
  Logger.log('  every dimension All : ' + all.grids + ' grids, ' + all.segments +
             ' segments, ' + all.routes.length + ' distinct routes' +
             (all.skippedNoScope ? ', ' + all.skippedNoScope + ' segment(s) out of scope' : ''));
  check('All-everything matches grids', all.grids > 0);
  check('grids are segments multiplied by regime and class',
        all.grids >= all.segments);
  check('routes are reported for the value form', all.routes.length > 0);

  const cc = bulkMixScope({ dimensions: { regime: ['CC'] } });
  Logger.log('  regime=[CC]         : ' + cc.grids + ' grids');
  check('picking one regime halves the grids at most', cc.grids <= all.grids);
  check('picking one regime does not change the segments',
        cc.segments === all.segments);

  const both = bulkMixScope({ dimensions: { regime: ['CC', 'AMBIENT'] } });
  check('ticking both regimes equals ticking neither', both.grids === all.grids);
  check('an empty list means All, not nothing',
        bulkMixScope({ dimensions: { regime: [] } }).grids === all.grids);
  check('an unknown brand matches nothing',
        bulkMixScope({ dimensions: { brand: ['NO_SUCH_BRAND'] } }).grids === 0);

  // ---- the refusal rule, using routes that really exist -------------------
  Logger.log('');
  Logger.log('--- the 100% rule ---');
  const one = all.routes[0];
  const whole = [{ carrier: one.carrier, method: one.method, value: 1 }];
  const base = { dimensions: {}, validFrom: '2035-01-01', validTo: '2035-12-31',
                 scenarioId: 1, preview: true };

  const pv = bulkUpdateMethodMix(Object.assign({}, base, { mix: whole }));
  Logger.log('  100% on ' + one.carrier + ' ' + one.method + ' : ' +
             pv.grids + ' grids would take it, ' + pv.refused.length + ' refused');
  check('grids without that route are refused, not renormalised',
        pv.grids + pv.refused.length === all.grids);
  check('every planned grid keeps its full route list',
        pv.rowsToWrite >= pv.grids);
  check('a preview writes nothing and says so', pv.preview === true);
  check('a preview issues a planKey', !!pv.planKey);

  let under = false, over = false;
  try { bulkUpdateMethodMix(Object.assign({}, base,
    { mix: [{ carrier: one.carrier, method: one.method, value: 0.99 }] })); }
  catch (e) { under = true; Logger.log('  refused, correctly: ' + e.message); }
  try { bulkUpdateMethodMix(Object.assign({}, base,
    { mix: [{ carrier: one.carrier, method: one.method, value: 1.01 }] })); }
  catch (e) { over = true; }
  check('99% is refused before anything is planned', under);
  check('101% is refused before anything is planned', over);

  let noKey = false;
  try {
    bulkUpdateMethodMix(Object.assign({}, base,
      { mix: whole, preview: false, planKey: 'deadbeef0000' }));
  } catch (e) { noKey = true; Logger.log('  refused, correctly: ' + e.message); }
  check('applying with a stale planKey is refused', noKey);

  const failed = checks.filter(function (c) { return !c.ok; });
  Logger.log('');
  Logger.log(failed.length ? failed.length + ' of ' + checks.length + ' FAILED'
                           : 'all ' + checks.length + ' checks passed');
  Logger.log('Nothing was written.');
  return { checks: checks.length, failed: failed.length,
           failures: failed.map(function (c) { return c.label; }) };
}


function testMixWrite() {
  requireMaintenance_();
  Logger.log('=== MIX WRITE TEST ===');
  const perms = requirePermissions_();
  Logger.log('  acting as ' + perms.email + ' (' + perms.role + ')');

  const HLID = 1, REGIME = 'CC', LP = 'PARCEL';

  Logger.log('');
  Logger.log('--- read the grid as at Jan-26 ---');
  const grid = getMethodMixGrid({ hlId: HLID, regime: REGIME, letterParcel: LP,
                                  asAt: '2026-01-01' });
  Logger.log('  High Level ID ' + HLID + ' / ' + REGIME + ' / ' + LP +
             ': ' + grid.rows.length + ' routes');
  grid.rows.filter(r => r.value > 0).slice(0, 6).forEach(r =>
    Logger.log('    ' + (r.carrier + '/' + r.method).slice(0, 34) +
               '   mix ' + (r.value * 100).toFixed(2) + '%' +
               '   rate ' + r.ratePerParcel.toFixed(4) +
               '   contributes ' + r.contribution.toFixed(4)));
  Logger.log('  TOTAL ' + (grid.total * 100).toFixed(2) + '%   ' +
             (grid.isValid ? 'valid' : 'INVALID'));

  Logger.log('');
  Logger.log('--- a grid that does not total 100% must be refused ---');
  let refused = false;
  try {
    saveMethodMixGrid({
      hlId: HLID, regime: REGIME, letterParcel: LP,
      validFrom: '2035-01-01', validTo: '2035-12-31',
      rows: grid.rows.map((r, i) => ({ modellingId: r.modellingId,
                                       value: i === 0 ? r.value + 0.10 : r.value }))
    });
  } catch (e) {
    refused = true;
    Logger.log('  refused, correctly: ' + e.message);
  }
  if (!refused) Logger.log('  PROBLEM — an invalid grid was accepted');

  Logger.log('');
  Logger.log('--- a grid missing a route must be refused ---');
  let refusedMissing = false;
  try {
    saveMethodMixGrid({
      hlId: HLID, regime: REGIME, letterParcel: LP,
      validFrom: '2035-01-01', validTo: '2035-12-31',
      rows: grid.rows.slice(1).map(r => ({ modellingId: r.modellingId, value: r.value }))
    });
  } catch (e) {
    refusedMissing = true;
    Logger.log('  refused, correctly: ' + e.message.slice(0, 120));
  }
  if (!refusedMissing) Logger.log('  PROBLEM — an incomplete grid was accepted');

  Logger.log('');
  Logger.log('--- normalise scales values back to 100% ---');
  const scaled = normaliseMixRows(grid.rows.map((r, i) =>
    ({ modellingId: r.modellingId, value: i === 0 ? r.value + 0.10 : r.value })));
  const scaledTotal = scaled.reduce((s, r) => s + r.value, 0);
  Logger.log('  after scaling: ' + (scaledTotal * 100).toFixed(4) + '%');

  Logger.log('');
  Logger.log('--- save a valid grid into 2035, outside the horizon ---');
  const res = saveMethodMixGrid({
    hlId: HLID, regime: REGIME, letterParcel: LP,
    validFrom: '2035-01-01', validTo: '2035-12-31',
    rows: grid.rows.map(r => ({ modellingId: r.modellingId, value: r.value })),
    notes: 'TEST — safe to delete'
  });
  Logger.log('  wrote ' + res.written + ' rows, adjusted ' + res.cleared +
             ' existing, total ' + (res.total * 100).toFixed(2) + '%');

  Logger.log('');
  const ok = refused && refusedMissing && res.ok && Math.abs(scaledTotal - 1) < 0.000001;
  Logger.log(ok ? 'MIX WRITES WORKING'
                : 'PROBLEM — one of the guards did not fire. Send me this log.');
  Logger.log('');
  Logger.log(res.written + ' test rows were written into 2035, outside your forecast');
  Logger.log('horizon, so no number moves. Filter Mix_Method by Valid_From 2035 to remove them.');
  return { ok: ok };
}