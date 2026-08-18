/**
 * Postage Forecast Portal — actuals.gs
 *
 * Recording what postage actually cost, and comparing it against the forecast.
 *
 * The blended rate is the number that matters, and it can arrive two ways:
 *
 *   spend and orders  -> rate is derived as spend / orders. Preferred, because
 *                        the two inputs are auditable and the rate cannot drift
 *                        out of step with them.
 *   rate only         -> stored as given, for when that is all you have.
 *
 * Comparison is at High Level ID x month, which is exactly the grain of OUTPUT,
 * so forecast and actual line up without any mapping.
 */

// ─────────────────────────────────────────────────────────────────────────────
// WRITING
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Record or update one month of actuals for one segment.
 *
 * @param {Object} p { id, hlId, monthStart, orders, totalSpend, blendedRate,
 *                     currency, source, sourceVersion, notes }
 */
function saveActual(p) {
  prewarmForWrite_([SHEET.ACTUALS, SHEET.ACTUALS_AMENDS,
                    SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CALENDAR]);
  const perms = requirePermissions_();
  requireWrite_(perms);

  const hlId = safeInt(p.hlId);
  assertCanEditHighLevelId_(perms, hlId);

  const month = monthStart(p.monthStart);
  if (!month) throw new Error('The month is missing or not a date.');

  const orders = safeNum(p.orders);
  const spend  = safeNum(p.totalSpend);
  let rate     = safeNum(p.blendedRate);

  // Derive the rate when both parts are present, so the three can never disagree.
  if (orders > 0 && spend > 0) {
    rate = spend / orders;
  } else if (!rate) {
    throw new Error('Enter either orders and total spend, or a blended rate.');
  }
  if (rate < 0) throw new Error('A blended rate cannot be negative.');

  const t = TABLES.ACTUALS, C = COL.ACTUALS, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    // The segment the row belongs to now, not only the one it is being pointed
    // at. Only bites when an id was supplied: the match-on-month path below
    // finds a row for hlId, which is already authorised.
    assertCanEditRowOwner_(perms, t.sheet, C.Actual_ID, p.id,
                           C.High_Level_ID, 'HIGH_LEVEL_ID');

    // one row per segment per month — a second would silently double count
    const data = getAllData_(t.sheet);
    const wantId = safeInt(p.id);
    let rowIndex = -1, before = null;
    for (let i = 1; i < data.length; i++) {
      // Only match on ID when one was actually supplied — safeInt('') is 0, so
      // an unsaved row would otherwise collide with any row whose ID is blank.
      if (wantId && safeInt(data[i][C.Actual_ID]) === wantId) { rowIndex = i + 1; break; }
      if (!wantId && safeBool(data[i][C.Active]) &&
          safeInt(data[i][C.High_Level_ID]) === hlId &&
          dateKey(data[i][C.Month_Start]) === dateKey(month)) {
        rowIndex = i + 1;    // same segment and month: update rather than add
        break;
      }
    }

    const sh = getSheet_(t.sheet);
    let row;
    if (rowIndex < 0) {
      row = blankRow_('ACTUALS');
      row[C.Actual_ID]  = getNextId_(t.sheet, C.Actual_ID);
      row[C.Created_TS] = new Date();
      row[C.Created_By] = perms.email;
      rowIndex = sh.getLastRow() + 1;
    } else {
      before = readRow_(t.sheet, rowIndex, width);
      row = before.slice();
    }

    row[C.High_Level_ID]  = hlId;
    row[C.Month_Start]    = month;
    row[C.Orders]         = orders || '';
    row[C.Total_Spend]    = spend || '';
    row[C.Blended_Rate]   = Math.round(rate * 1000000) / 1000000;
    row[C.Currency]       = safeStr(p.currency).toUpperCase() || currencyForHighLevelId_(hlId);
    row[C.Source]         = safeStr(p.source) || 'manual';
    row[C.Source_Version] = safeStr(p.sourceVersion);
    row[C.Active]         = true;
    row[C.Notes]          = safeStr(p.notes);
    row[C.Updated_TS]     = new Date();
    row[C.Updated_By]     = perms.email;

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);
    recordChange_('ACTUALS', row[C.Actual_ID], before, row, before ? 'UPDATE' : 'CREATE');
    return { ok: true, id: row[C.Actual_ID], blendedRate: row[C.Blended_Rate],
             derived: (orders > 0 && spend > 0) };
  });
}


function deleteActual(id) {
  prewarmForWrite_([SHEET.ACTUALS, SHEET.ACTUALS_AMENDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireWrite_(perms);

  const t = TABLES.ACTUALS, C = COL.ACTUALS, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.Actual_ID, id);
    if (rowIndex < 0) throw new Error('That actual no longer exists.');

    const before = readRow_(t.sheet, rowIndex, width);
    assertCanEditHighLevelId_(perms, safeInt(before[C.High_Level_ID]));

    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('ACTUALS', safeInt(id), before, row, 'DELETE');
    return { ok: true };
  });
}


/**
 * Load many months at once — a paste from a spreadsheet, or a Metabase pull.
 *
 * @param {Array} rows [{ hlId, monthStart, orders, totalSpend, blendedRate }]
 * @param {Object} opts { source, sourceVersion }
 */
function importActuals(rows, opts) {
  prewarmForWrite_([SHEET.ACTUALS, SHEET.ACTUALS_AMENDS,
                    SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CALENDAR]);
  const perms = requirePermissions_();
  requireWrite_(perms);
  opts = opts || {};

  if (!rows || !rows.length) throw new Error('Nothing to import.');

  const results = { written: 0, skipped: 0, errors: [] };
  rows.forEach(function (r, i) {
    try {
      saveActual({ hlId: r.hlId, monthStart: r.monthStart, orders: r.orders,
                   totalSpend: r.totalSpend, blendedRate: r.blendedRate,
                   currency: r.currency,
                   source: opts.source || 'import',
                   sourceVersion: opts.sourceVersion || '' });
      results.written++;
    } catch (e) {
      results.skipped++;
      if (results.errors.length < 10) results.errors.push('row ' + (i + 1) + ': ' + e.message);
    }
  });

  logAudit_('IMPORT', SHEET.ACTUALS, '', '', '',
            results.written + ' rows', (opts.source || 'import'), true);
  return results;
}


function currencyForHighLevelId_(hlId) {
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) {
    if (safeInt(hl[i][H.High_Level_ID]) === safeInt(hlId)) {
      return safeStr(hl[i][H.Currency]) || 'GBP';
    }
  }
  return 'GBP';
}


// ─────────────────────────────────────────────────────────────────────────────
// READING AND COMPARING
// ─────────────────────────────────────────────────────────────────────────────

function loadActualsForClient_(visibleHl) {
  const data = getAllData_(SHEET.ACTUALS), C = COL.ACTUALS, out = [];
  for (let i = 1; i < data.length; i++) {
    const hl = safeInt(data[i][C.High_Level_ID]);
    if (!hl || !visibleHl[hl]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.Actual_ID]), hlId: hl,
               monthStart: fmtDate(data[i][C.Month_Start]),
               orders: safeNum(data[i][C.Orders]),
               totalSpend: safeNum(data[i][C.Total_Spend]),
               blendedRate: safeNum(data[i][C.Blended_Rate]),
               currency: safeStr(data[i][C.Currency]),
               source: safeStr(data[i][C.Source]),
               sourceVersion: safeStr(data[i][C.Source_Version]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out;
}


/**
 * Forecast against actual, per segment per month, for whatever months have both.
 *
 * Variance is actual minus forecast, so a positive number means postage cost
 * more than the model said.
 */
function getActualsComparison(scenarioId) {
  const perms = requirePermissions_();
  prewarmSheetCache_([SHEET.ACTUALS, SHEET.OUTPUT, SHEET.HIGH_LEVEL_IDS,
                      SHEET.DIM_CALENDAR, SHEET.CONFIG, SHEET.PERMISSIONS,
                      SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING]);
  scenarioId = safeInt(scenarioId) || 1;
  const visible = visibleHighLevelIds_(perms);
  const warnPct = safeNum(configStr('ACTUALS_VARIANCE_WARN_PCT', '10')) / 100;

  const monthByKey = {};
  const cal = getAllData_(SHEET.DIM_CALENDAR), K = COL.DIM_CALENDAR;
  for (let i = 1; i < cal.length; i++) {
    const ms = normaliseDate(cal[i][K.Month_Start]);
    if (!ms) continue;
    monthByKey[dateKey(ms)] = { dateId: safeInt(cal[i][K.Date_ID]),
                                label: monthLabel_(cal[i][K.Month_Label], ms) };
  }

  const forecast = {};
  const out = getAllData_(SHEET.OUTPUT), O = COL.OUTPUT;
  for (let i = 1; i < out.length; i++) {
    const hl = safeInt(out[i][O.High_Level_ID]);
    if (!hl || !visible[hl]) continue;
    if (safeInt(out[i][O.Scenario_ID]) !== scenarioId) continue;
    forecast[hl + '|' + dateKey(out[i][O.Month_Start])] =
      safeNum(out[i][O.Forecast_Rate_Per_Order]);
  }

  const rows = [];
  let matched = 0, withinTolerance = 0, orphanActuals = 0;
  const actuals = loadActualsForClient_(visible);

  actuals.forEach(function (a) {
    const key = a.hlId + '|' + dateKey(a.monthStart);
    const f = forecast[key];
    const m = monthByKey[dateKey(a.monthStart)] || { dateId: 0, label: a.monthStart };
    if (f === undefined) {
      orphanActuals++;
      rows.push({ hlId: a.hlId, dateId: m.dateId, month: m.label,
                  monthStart: a.monthStart, forecast: null, actual: a.blendedRate,
                  variance: null, variancePct: null, orders: a.orders,
                  totalSpend: a.totalSpend, source: a.source, id: a.id,
                  flag: 'NO_FORECAST' });
      return;
    }
    matched++;
    const v = a.blendedRate - f;
    const vp = f ? v / f : null;
    const flag = (vp !== null && Math.abs(vp) > warnPct) ? 'OVER_TOLERANCE' : '';
    if (!flag) withinTolerance++;
    rows.push({ hlId: a.hlId, dateId: m.dateId, month: m.label,
                monthStart: a.monthStart, forecast: f, actual: a.blendedRate,
                variance: v, variancePct: vp, orders: a.orders,
                totalSpend: a.totalSpend, source: a.source, id: a.id, flag: flag });
  });

  rows.sort(function (a, b) { return (a.hlId - b.hlId) || (a.dateId - b.dateId); });

  // weighted where orders are known, simple mean otherwise
  let wSpend = 0, wOrders = 0, sumF = 0, sumA = 0, n = 0;
  rows.forEach(function (r) {
    if (r.forecast === null) return;
    n++; sumF += r.forecast; sumA += r.actual;
    if (r.orders > 0) { wOrders += r.orders; wSpend += r.actual * r.orders; }
  });

  return {
    rows: rows,
    summary: {
      months: rows.length, matched: matched, orphanActuals: orphanActuals,
      withinTolerance: withinTolerance,
      tolerancePct: warnPct,
      meanForecast: n ? sumF / n : null,
      meanActual: n ? sumA / n : null,
      meanVariancePct: (n && sumF) ? (sumA - sumF) / sumF : null,
      weightedActual: wOrders ? wSpend / wOrders : null,
      totalOrders: wOrders
    },
    canEdit: perms.allAccess || perms.capabilities.write
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// CHECKING WHAT WAS IMPORTED
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Show every actual on record for one segment, month by month.
 *
 * A blended rate of zero is the symptom worth looking for. It means the cost
 * column did not parse — most often because a mangled currency prefix such as
 * "Â£41.65" left a stray letter that parseFloat could not read, so the total
 * came out as nothing and the rate with it.
 */
function diagnoseActuals(hlId) {
  requireMaintenance_();
  hlId = safeInt(hlId) || 1;
  Logger.log('=== ACTUALS FOR SEGMENT ' + hlId + ' ===');

  prewarmSheetCache_([SHEET.ACTUALS, SHEET.HIGH_LEVEL_IDS, SHEET.OUTPUT, SHEET.CONFIG]);

  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) {
    if (safeInt(hl[i][H.High_Level_ID]) !== hlId) continue;
    Logger.log('  ' + safeStr(hl[i][H.Brand]) + ' ' + safeStr(hl[i][H.Geo]) + ' ' +
               safeStr(hl[i][H.Treatment_Type]) + ' ' + safeStr(hl[i][H.WL_Split]));
    break;
  }

  const fc = {};
  const out = getAllData_(SHEET.OUTPUT), O = COL.OUTPUT;
  for (let i = 1; i < out.length; i++) {
    if (safeInt(out[i][O.High_Level_ID]) !== hlId) continue;
    fc[fmtDate(out[i][O.Month_Start])] = safeNum(out[i][O.Forecast_Rate_Per_Order]);
  }

  const data = getAllData_(SHEET.ACTUALS), C = COL.ACTUALS;
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][C.High_Level_ID]) !== hlId) continue;
    if (!safeBool(data[i][C.Active])) continue;
    rows.push({ month: fmtDate(data[i][C.Month_Start]),
                orders: safeNum(data[i][C.Orders]),
                spend: safeNum(data[i][C.Total_Spend]),
                rate: safeNum(data[i][C.Blended_Rate]),
                source: safeStr(data[i][C.Source]) });
  }
  rows.sort(function (a, b) { return a.month < b.month ? -1 : 1; });

  if (!rows.length) { Logger.log('  No actuals recorded for this segment.'); return { rows: 0 }; }

  Logger.log('');
  Logger.log('  month       orders       spend        rate    forecast   source');
  Logger.log('  ' + new Array(72).join('-'));
  let zeros = 0, noSpend = 0;
  rows.forEach(function (r) {
    const bad = (r.rate === 0);
    if (bad) zeros++;
    if (r.orders > 0 && r.spend === 0) noSpend++;
    Logger.log('  ' + r.month + ' ' + pad_(r.orders ? Math.round(r.orders) : '-', 11) + ' ' +
               pad_(r.spend ? r.spend.toFixed(2) : '-', 12) + ' ' +
               pad_(r.rate.toFixed(4), 10) + '  ' +
               pad_(fc[r.month] !== undefined ? fc[r.month].toFixed(4) : '-', 10) + '   ' +
               r.source + (bad ? '   <- ZERO' : ''));
  });

  Logger.log('');
  Logger.log('  rows: ' + rows.length + ', of which ' + zeros + ' have a rate of zero');
  if (noSpend) {
    Logger.log('  ' + noSpend + ' row(s) have shipments but no spend — the cost column did not parse.');
  }
  if (zeros) {
    Logger.log('');
    Logger.log('  WHAT TO DO: the cost values were not read. That happens when the currency');
    Logger.log('  symbol is mangled, so "Â£41.65" is left with a stray letter that cannot be');
    Logger.log('  turned into a number. The importer now strips anything that is not a digit,');
    Logger.log('  so re-running it fixes this:');
    Logger.log('    1. clearActuals()             removes what is there');
    Logger.log('    2. previewActualsImport()     check the numbers look right');
    Logger.log('    3. runActualsImport()         load them again');
  } else {
    Logger.log('  Every row has a rate. Nothing wrong here.');
  }
  return { rows: rows.length, zeros: zeros };
}


/** Remove every actual, so an import can start clean. */
function clearActuals() {
  requireMaintenance_();
  const before = Math.max(getSheet_(SHEET.ACTUALS).getLastRow() - 1, 0);
  clearDataRows_(SHEET.ACTUALS);
  Logger.log('Removed ' + before + ' actual row(s).');
  Logger.log('The Actuals_Amends history is left alone — it is the record of what happened.');
  Logger.log('Next: previewActualsImport(), then runActualsImport().');
  logAudit_('DELETE', SHEET.ACTUALS, '', '', '', String(before) + ' rows', 'clearActuals', true);
  return { removed: before };
}


// ─────────────────────────────────────────────────────────────────────────────
// IMPORTING FROM THE MONTHLY EXTRACT
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Turn the monthly postage extract into actuals.
 *
 * The extract is one row per country, brand, carrier, method, treatment and
 * month, with a shipment count and a total cost. Three things have to happen to
 * make that comparable with the forecast:
 *
 *   1. Aggregate away carrier and method. The forecast is a blended rate per
 *      order, so the actual has to be blended the same way: total cost divided
 *      by total shipments for the whole segment-month, not per route.
 *
 *   2. Collapse ~50 treatment values into the two the model uses. Anything in
 *      ACTUALS_WL_TREATMENTS is weight loss; everything else is Core Rx.
 *
 *   3. Handle the split pairs. The extract does not distinguish Mounjaro from
 *      WeGovy, but the forecast rate is identical for both, so the measured rate
 *      applies correctly to each. The rate therefore goes on BOTH halves, while
 *      shipments and spend go on the lower ID only — otherwise every shipment
 *      would be counted twice in the volume-weighted comparison.
 */

/**
 * A header reduced to letters and digits.
 *
 * Everything else goes: spaces, punctuation, and any character outside A-Z0-9 —
 * which includes both halves of a mangled "Â£". That makes the match survive
 * encoding damage, and lets "Sum of Cost Of Shipping" and "sum_of_cost_of_shipping"
 * mean the same thing.
 */
function normHeader_(s) {
  return safeStr(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function wlTreatments_() {
  const raw = configStr('ACTUALS_WL_TREATMENTS', 'WeightLossGlp1');
  const set = {};
  raw.split(',').forEach(function (t) {
    const s = safeStr(t).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (s) set[s] = true;
  });
  return set;
}

/** "January, 2026" and similar into a month start. */
function parseExtractMonth_(v) {
  if (isDate_(v)) return monthStart(v);
  const s = safeStr(v).replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
  const m = s.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (m) {
    const idx = MONTH_ABBR.map(function (x) { return x.toUpperCase(); })
                          .indexOf(m[1].slice(0, 3).toUpperCase());
    if (idx >= 0) return new Date(safeInt(m[2]), idx, 1);
  }
  return monthStart(normaliseDate(v));
}

/** Numbers arrive with thousands separators when pasted from a spreadsheet. */
function parseExtractNumber_(v) {
  if (typeof v === 'number') return v;
  // strip separators, currency symbols and any non-ASCII left by a bad decode
  return safeNum(safeStr(v).replace(/[^0-9.\-]/g, ''));
}


/**
 * Read the staging tab, map it, and report. Writes nothing unless told to.
 * @param {boolean} commit
 */
function importActualsFromStaging(commit) {
  requireMaintenance_();
  Logger.log(commit ? '=== IMPORTING ACTUALS ===' : '=== ACTUALS IMPORT PREVIEW (nothing written) ===');

  prewarmSheetCache_([SHEET.ACTUALS_IMPORT, SHEET.HIGH_LEVEL_IDS, SHEET.ACTUALS,
                      SHEET.DIM_CALENDAR, SHEET.CONFIG]);

  const rows = getAllData_(SHEET.ACTUALS_IMPORT);
  if (rows.length < 2) {
    Logger.log('  The Actuals_Import tab is empty.');
    Logger.log('  Paste the extract there, including its header row, then run this again.');
    return { ok: false, reason: 'empty' };
  }

  // Columns are found by name, so a reordered extract still works. Names are
  // compared with everything except letters and digits stripped out, because a
  // header like "Sum of Cost Of Shipping (£)" is fragile in ways that have
  // nothing to do with the data: the pound sign is two bytes in UTF-8 and comes
  // back as "Â£" whenever something in the chain reads it as Latin-1. Spacing,
  // capitalisation and punctuation vary between exports too.
  const hdrRaw = rows[0].map(function (h) { return safeStr(h); });
  const hdrNorm = hdrRaw.map(normHeader_);

  function col(candidates, label) {
    for (let ci = 0; ci < candidates.length; ci++) {
      const want = normHeader_(candidates[ci]);
      const exact = hdrNorm.indexOf(want);
      if (exact >= 0) return exact;
    }
    // nothing exact — accept a header that contains the name, or is contained by it
    for (let ci = 0; ci < candidates.length; ci++) {
      const want = normHeader_(candidates[ci]);
      for (let hi = 0; hi < hdrNorm.length; hi++) {
        if (!hdrNorm[hi]) continue;
        if (hdrNorm[hi].indexOf(want) >= 0 || want.indexOf(hdrNorm[hi]) >= 0) return hi;
      }
    }
    throw new Error('The extract has no ' + label + ' column. It needs one named something ' +
      'like "' + candidates[0] + '". Columns found: ' + hdrRaw.join(' | '));
  }

  const cCountry = col(['Country Code', 'Country'], 'country'),
        cBrand   = col(['Brand'], 'brand'),
        cTreat   = col(['Treatment Type', 'Treatment'], 'treatment type'),
        cMonth   = col(['Dispatched Date: Month', 'Dispatched Month', 'Month'], 'month'),
        cCount   = col(['Count', 'Shipments', 'Orders'], 'shipment count'),
        cCost    = col(['Sum of Cost Of Shipping', 'Cost Of Shipping', 'Total Cost', 'Cost'],
                       'shipping cost');

  Logger.log('  columns matched:');
  [['country', cCountry], ['brand', cBrand], ['treatment', cTreat],
   ['month', cMonth], ['count', cCount], ['cost', cCost]].forEach(function (x) {
    Logger.log('    ' + pad_(x[0], 10) + ' -> "' + hdrRaw[x[1]] + '"');
  });

  const wl = wlTreatments_();

  // ---- segments, keyed on brand / geo / treatment -------------------------
  const byKey = {};
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id || !safeBool(hl[i][H.Active])) continue;
    const key = normKey(hl[i][H.Brand]) + '|' + normKey(hl[i][H.Geo]) + '|' +
                normKey(hl[i][H.Treatment_Type]);
    (byKey[key] = byKey[key] || []).push({ id: id, split: safeStr(hl[i][H.WL_Split]) });
  }
  for (const k in byKey) byKey[k].sort(function (a, b) { return a.id - b.id; });

  // ---- horizon, so out-of-range months can be reported --------------------
  const cal = getAllData_(SHEET.DIM_CALENDAR), K = COL.DIM_CALENDAR;
  const inHorizon = {};
  for (let i = 1; i < cal.length; i++) {
    if (!safeBool(cal[i][K.In_Horizon])) continue;
    const ms = monthStart(normaliseDate(cal[i][K.Month_Start]));
    if (ms) inHorizon[dateKey(ms)] = true;
  }

  // ---- aggregate ----------------------------------------------------------
  const agg = {}, unmatched = {}, outside = {};
  let read = 0, skipped = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const brand = normKey(r[cBrand]), geo = normKey(r[cCountry]);
    if (!brand || !geo) { skipped++; continue; }

    const treatRaw = safeStr(r[cTreat]);
    const treat = wl[treatRaw.toUpperCase().replace(/[^A-Z0-9]/g, '')] ? 'WL' : 'CORE_RX';
    const ms = parseExtractMonth_(r[cMonth]);
    if (!ms) { skipped++; continue; }

    const n = parseExtractNumber_(r[cCount]);
    const cost = parseExtractNumber_(r[cCost]);
    if (!n) { skipped++; continue; }
    read++;

    const segKey = brand + '|' + geo + '|' + treat;
    if (!byKey[segKey]) {
      const u = unmatched[segKey] || (unmatched[segKey] = { ship: 0, cost: 0 });
      u.ship += n; u.cost += cost;
      continue;
    }
    if (!inHorizon[dateKey(ms)]) {
      const o = outside[fmtDate(ms)] || (outside[fmtDate(ms)] = { ship: 0 });
      o.ship += n;
      continue;
    }

    const k = segKey + '|' + dateKey(ms);
    const a = agg[k] || (agg[k] = { segKey: segKey, month: ms, ship: 0, cost: 0 });
    a.ship += n; a.cost += cost;
  }

  Logger.log('  extract rows read      : ' + read + (skipped ? '  (' + skipped + ' skipped as blank)' : ''));
  Logger.log('  segment-months to load : ' + Object.keys(agg).length);

  // ---- what would be written ---------------------------------------------
  const toWrite = [];
  Object.keys(agg).forEach(function (k) {
    const a = agg[k];
    const segs = byKey[a.segKey];
    const rate = a.cost / a.ship;
    segs.forEach(function (s, idx) {
      toWrite.push({
        hlId: s.id, monthStart: fmtDate(a.month), rate: rate,
        // shipments and spend only on the first of a pair, so nothing is double counted
        orders: idx === 0 ? a.ship : 0,
        spend:  idx === 0 ? a.cost : 0,
        pairNote: segs.length > 1
          ? ('Extract does not split ' + segs.map(function (x) { return x.split; }).join(' / ') +
             '; the rate is the same for both. ' +
             (idx === 0 ? 'Shipments and spend recorded here.' : 'Shipments and spend on High Level ID ' + segs[0].id + '.'))
          : ''
      });
    });
  });

  Logger.log('  actual rows to write   : ' + toWrite.length);

  if (Object.keys(unmatched).length) {
    Logger.log('');
    Logger.log('--- no segment for these, so they are left out ---');
    Object.keys(unmatched).sort().forEach(function (k) {
      const p = k.split('|');
      Logger.log('  ' + p[0] + ' ' + p[1] + ' ' + p[2].replace('_', ' ') +
                 ': ' + Math.round(unmatched[k].ship).toLocaleString() + ' shipments, £' +
                 unmatched[k].cost.toFixed(2));
    });
    Logger.log('  Add a High Level ID for these if they should be forecast.');
  }

  if (Object.keys(outside).length) {
    let n = 0;
    Object.keys(outside).forEach(function (k) { n += outside[k].ship; });
    Logger.log('');
    Logger.log('--- ' + Object.keys(outside).length + ' month(s) outside the forecast horizon, ' +
               Math.round(n).toLocaleString() + ' shipments, ignored ---');
  }

  if (!commit) {
    Logger.log('');
    Logger.log('  Nothing written. Run importActualsFromStaging(true) to load them.');
    return { ok: true, preview: true, rows: toWrite.length,
             unmatched: Object.keys(unmatched).length };
  }

  // ---- write --------------------------------------------------------------
  Logger.log('');
  Logger.log('--- writing ---');
  const res = importActuals(toWrite.map(function (w) {
    return { hlId: w.hlId, monthStart: w.monthStart,
             orders: w.orders, totalSpend: w.spend,
             blendedRate: w.rate };
  }), { source: 'extract', sourceVersion: fmtDate(new Date()) });

  Logger.log('  written : ' + res.written);
  if (res.skipped) {
    Logger.log('  skipped : ' + res.skipped);
    res.errors.forEach(function (e) { Logger.log('    ' + e); });
  }
  Logger.log('');
  Logger.log('  Open the Actuals tab in the portal to see them against the forecast.');
  return { ok: true, preview: false, written: res.written, skipped: res.skipped };
}


/** Convenience wrappers, so both appear in the editor's function list. */
function previewActualsImport() { return importActualsFromStaging(false); }
function runActualsImport()     { return importActualsFromStaging(true); }


// ─────────────────────────────────────────────────────────────────────────────
// METABASE — not wired up yet, deliberately
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Check whether this project can reach Metabase at all.
 *
 * Run this BEFORE any work is done on an automatic feed. Two things have to be
 * true and neither can be assumed:
 *
 *   1. Google's servers can resolve and reach the host. If Metabase is
 *      VPN-only or IP-allowlisted, Apps Script cannot get to it and no amount
 *      of code will change that.
 *   2. An API key is stored in Script Properties. It must NOT go in the Config
 *      tab — anyone who can read the spreadsheet could read the key.
 *
 * To set the key: Apps Script > Project Settings > Script Properties >
 * Add property, name METABASE_API_KEY.
 */
function testMetabaseConnection() {
  requireMaintenance_();
  Logger.log('=== METABASE REACHABILITY TEST ===');

  const url = safeStr(configStr('METABASE_URL', ''));
  if (!url) {
    Logger.log('  METABASE_URL is blank in the Config tab.');
    Logger.log('  Set it to something like https://metabase.heliosx.co and run again.');
    return { ok: false, reason: 'no url' };
  }

  let key = '';
  try { key = PropertiesService.getScriptProperties().getProperty('METABASE_API_KEY') || ''; }
  catch (e) { key = ''; }

  Logger.log('  URL     : ' + url);
  Logger.log('  API key : ' + (key ? 'present (' + key.length + ' characters)' : 'NOT SET'));
  if (!key) {
    Logger.log('');
    Logger.log('  Add it under Project Settings > Script Properties, named METABASE_API_KEY.');
    Logger.log('  Do not put it in the Config tab — that is readable by anyone with the sheet.');
    return { ok: false, reason: 'no key' };
  }

  Logger.log('');
  Logger.log('--- calling /api/user/current ---');
  let resp;
  try {
    resp = UrlFetchApp.fetch(url.replace(/\/+$/, '') + '/api/user/current', {
      method: 'get',
      headers: { 'X-API-KEY': key },
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
  } catch (e) {
    Logger.log('  COULD NOT CONNECT: ' + e.message);
    Logger.log('');
    Logger.log('  Almost always means Google cannot reach the host — a private network,');
    Logger.log('  an IP allowlist, or a firewall. That has to be resolved on the Metabase');
    Logger.log('  side before a feed is possible.');
    return { ok: false, reason: 'unreachable', error: e.message };
  }

  const code = resp.getResponseCode();
  Logger.log('  HTTP ' + code);
  if (code === 200) {
    let who = '';
    try { who = JSON.parse(resp.getContentText()).email || ''; } catch (e) {}
    Logger.log('  Connected as: ' + (who || '(unknown)'));
    Logger.log('');
    Logger.log('  Reachable and authenticated. Next question is the data, not the plumbing:');
    Logger.log('  which saved question returns actual spend and orders, and at what grain?');
    Logger.log('  Set METABASE_CARD_ID in Config once you know, then send me a sample row.');
    return { ok: true, user: who };
  }
  if (code === 401 || code === 403) {
    Logger.log('  The key was rejected. Check it is a Metabase API key, not a session token,');
    Logger.log('  and that the account it belongs to has access.');
    return { ok: false, reason: 'auth', code: code };
  }
  Logger.log('  Unexpected response: ' + resp.getContentText().slice(0, 300));
  return { ok: false, reason: 'http ' + code };
}


/**
 * Fetch a saved question and show its shape, without storing anything.
 *
 * The point is to see the column names and one row of data, so the mapping to
 * High Level ID x month can be worked out before any importing happens.
 */
function previewMetabaseCard() {
  requireMaintenance_();
  Logger.log('=== METABASE CARD PREVIEW (nothing stored) ===');

  const url = safeStr(configStr('METABASE_URL', ''));
  const card = safeStr(configStr('METABASE_CARD_ID', ''));
  if (!url || !card) {
    Logger.log('  Set METABASE_URL and METABASE_CARD_ID in Config first.');
    return { ok: false };
  }
  let key = '';
  try { key = PropertiesService.getScriptProperties().getProperty('METABASE_API_KEY') || ''; }
  catch (e) {}
  if (!key) { Logger.log('  METABASE_API_KEY is not set in Script Properties.'); return { ok: false }; }

  let resp;
  try {
    resp = UrlFetchApp.fetch(url.replace(/\/+$/, '') + '/api/card/' + card + '/query/json', {
      method: 'post',
      headers: { 'X-API-KEY': key },
      contentType: 'application/json',
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('  Could not connect: ' + e.message);
    return { ok: false };
  }

  if (resp.getResponseCode() !== 200) {
    Logger.log('  HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 300));
    return { ok: false };
  }

  let data;
  try { data = JSON.parse(resp.getContentText()); }
  catch (e) { Logger.log('  Response was not JSON.'); return { ok: false }; }

  if (!data.length) { Logger.log('  The question returned no rows.'); return { ok: true, rows: 0 }; }

  Logger.log('  rows returned: ' + data.length);
  Logger.log('');
  Logger.log('--- columns ---');
  Object.keys(data[0]).forEach(k => Logger.log('  ' + k + '  =  ' + String(data[0][k]).slice(0, 60)));
  Logger.log('');
  Logger.log('  To turn this into actuals I need to know which column is the month, which');
  Logger.log('  identifies the segment, and which are spend and order count. Send me this');
  Logger.log('  list and I will write the mapping.');
  return { ok: true, rows: data.length, columns: Object.keys(data[0]) };
}