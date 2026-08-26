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


/**
 * Every distinct Treatment Type in the staging extract, and where it lands.
 *
 * Reads only. The import collapses ~50 raw treatment strings into the two the
 * model uses, by testing each against ACTUALS_WL_TREATMENTS — anything that
 * does not match is Core Rx. That default is silent, and silence is the problem:
 * a weight-loss row that fails to match is not dropped, it is ADDED to the Core
 * Rx segment for the same brand and country, where it moves both the volume and
 * the blended rate (the rate is cost/shipments for the whole segment-month). The
 * forecast then fails to reconcile on both sides at once, with nothing in the
 * log to say why.
 *
 * So this prints every distinct value with its money attached, biggest first,
 * and says which bucket today's config puts it in. What it deliberately does NOT
 * do is guess which of them ought to be weight loss. Nothing in the sheet marks
 * that, and a guess printed in a log reads like a finding.
 *
 * It also lists any High Level ID whose Treatment_Type is not literally WL or
 * CORE_RX. The importer can only ever produce those two, so any other value is a
 * segment that can never receive actuals automatically. saveHighLevelId checks
 * Treatment_Type against Dim_Reference, so such a value cannot be created through
 * the portal — but migrate.gs writes the column unchecked, and nothing in
 * validate.gs polices it afterwards.
 */
function diagnoseActualsTreatments() {
  requireMaintenance_();
  Logger.log('=== ACTUALS IMPORT — TREATMENT TYPES (read-only) ===');

  prewarmSheetCache_([SHEET.ACTUALS_IMPORT, SHEET.HIGH_LEVEL_IDS,
                      SHEET.DIM_REFERENCE, SHEET.CONFIG]);

  // ---- what the config says today ----------------------------------------
  const rawCfg = configStr('ACTUALS_WL_TREATMENTS', 'WeightLossGlp1');
  const wl = wlTreatments_();
  Logger.log('  Config ACTUALS_WL_TREATMENTS : ' + rawCfg);
  Logger.log('  counts as weight loss        : ' + Object.keys(wl).sort().join(', '));
  Logger.log('  (matching ignores case, spaces and punctuation — only the letters');
  Logger.log('   and digits are compared, so "Weight Loss GLP-1" and "weightlossglp1"');
  Logger.log('   are the same string as far as the import is concerned.)');

  const rows = getAllData_(SHEET.ACTUALS_IMPORT);
  if (rows.length < 2) {
    Logger.log('');
    Logger.log('  The Actuals_Import tab is empty — paste the extract there first.');
    return { ok: false, reason: 'empty' };
  }

  // ---- columns, found the same way the importer finds them ----------------
  const X = extractStandardColumns_(rows[0]);
  const cCountry = X.cCountry, cBrand = X.cBrand, cTreat = X.cTreat,
        cCount = X.cCount, cCost = X.cCost;
  Logger.log('');
  logExtractColumns_(X.cols);

  // ---- the segments the importer can actually reach -----------------------
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  const segExists = {}, oddSegments = [];
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id) continue;
    const tt = normKey(hl[i][H.Treatment_Type]);
    if (safeBool(hl[i][H.Active])) {
      segExists[normKey(hl[i][H.Brand]) + '|' + normKey(hl[i][H.Geo]) + '|' + tt] = true;
    }
    if (tt !== 'WL' && tt !== 'CORE_RX') {
      oddSegments.push({ id: id, brand: safeStr(hl[i][H.Brand]),
                         geo: safeStr(hl[i][H.Geo]), tt: safeStr(hl[i][H.Treatment_Type]),
                         split: safeStr(hl[i][H.WL_Split]),
                         active: safeBool(hl[i][H.Active]) });
    }
  }

  // ---- every distinct treatment string, with its money --------------------
  const seen = {};
  let readRows = 0, blankTreat = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const brand = normKey(r[cBrand]), geo = normKey(r[cCountry]);
    if (!brand || !geo) continue;
    const n = parseExtractNumber_(r[cCount]);
    if (!n) continue;                       // the importer skips these too
    readRows++;

    const raw  = safeStr(r[cTreat]);
    const norm = raw.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!norm) blankTreat++;
    const bucket = wl[norm] ? 'WL' : 'CORE_RX';

    /* Grouped by the NORMALISED value, not the raw one, because normalised is
       what the config matches on: two spellings differing only in case or
       punctuation are one thing to classify, and listing them apart would invite
       putting both in the config where either alone would do. The spellings
       actually seen are kept, so they can still be recognised below. */
    const e = seen[norm] || (seen[norm] = { norm: norm, bucket: bucket,
                                            spellings: {}, rows: 0, count: 0, cost: 0,
                                            noSeg: 0, noSegCost: 0, noSegWhere: {} });
    e.spellings[raw] = (e.spellings[raw] || 0) + 1;
    e.rows++;
    e.count += n;
    e.cost  += parseExtractNumber_(r[cCost]);
    if (!segExists[brand + '|' + geo + '|' + bucket]) {
      e.noSeg++;
      e.noSegCost += parseExtractNumber_(r[cCost]);
      /* Kept per brand and country, because "21 rows have nowhere to go" is a
         curiosity and "MEDEXPRESS GB has nowhere to go" is an action. */
      const w = e.noSegWhere[brand + ' ' + geo] ||
                (e.noSegWhere[brand + ' ' + geo] = { rows: 0, count: 0, cost: 0 });
      w.rows++; w.count += n; w.cost += parseExtractNumber_(r[cCost]);
    }
  }

  const list = Object.keys(seen).map(function (k) { return seen[k]; });
  list.sort(function (a, b) { return b.cost - a.cost; });

  Logger.log('');
  Logger.log('--- ' + list.length + ' distinct Treatment Type value(s) across ' +
             readRows + ' usable extract row(s) ---');
  Logger.log('  ' + pad_('bucket', 8) + '  ' + pad_('rows', 6) + '  ' +
             pad_('count', 12) + '  ' + pad_('cost', 14) + '  ' +
             pad_('£/ship', 9) + '  ' + pad_('no seg', 7) + '  treatment type');
  list.forEach(function (e) {
    const sp = Object.keys(e.spellings);
    Logger.log('  ' + pad_(e.bucket, 8) + '  ' + pad_(String(e.rows), 6) + '  ' +
               pad_(Math.round(e.count).toLocaleString(), 12) + '  ' +
               pad_('£' + e.cost.toFixed(2), 14) + '  ' +
               /* The implied cost per shipment. It is not used for anything, but a
                  value that is an order of magnitude away from its neighbours is a
                  data problem worth seeing next to the money, not after it. */
               pad_('£' + (e.count ? (e.cost / e.count).toFixed(3) : '0.000'), 9) + '  ' +
               pad_(e.noSeg ? String(e.noSeg) : '-', 7) + '  ' +
               (sp[0] === '' ? '(blank)' : sp[0]) +
               (sp.length > 1 ? '   [+' + (sp.length - 1) + ' spelling' +
                                (sp.length > 2 ? 's' : '') + ']' : ''));
  });

  /* Only printed when it happens: a value spelled more than one way still needs
     just one entry in the config, and seeing the variants is how you know that
     rather than assuming it. */
  const varied = list.filter(function (e) { return Object.keys(e.spellings).length > 1; });
  if (varied.length) {
    Logger.log('');
    Logger.log('--- spelled more than one way (one config entry covers each group) ---');
    varied.forEach(function (e) {
      Logger.log('  ' + e.norm + ':');
      Object.keys(e.spellings).forEach(function (sp) {
        Logger.log('      "' + sp + '"  x' + e.spellings[sp]);
      });
    });
  }

  // ---- totals -------------------------------------------------------------
  const tot = { WL: { v: 0, rows: 0, count: 0, cost: 0 },
                CORE_RX: { v: 0, rows: 0, count: 0, cost: 0 } };
  let silent = 0, silentCost = 0;
  list.forEach(function (e) {
    const t = tot[e.bucket];
    t.v++; t.rows += e.rows; t.count += e.count; t.cost += e.cost;
    // Rows that DO find a segment are the dangerous ones: a misclassified row
    // here is absorbed into the wrong segment rather than reported as unmatched.
    if (e.bucket === 'CORE_RX') { silent += e.rows - e.noSeg; silentCost += e.cost - e.noSegCost; }
  });
  Logger.log('');
  Logger.log('--- totals ---');
  ['WL', 'CORE_RX'].forEach(function (b) {
    const t = tot[b];
    Logger.log('  ' + pad_(b, 8) + '  ' + pad_(String(t.v), 3) + ' value(s), ' +
               pad_(String(t.rows), 6) + ' row(s), ' +
               pad_(Math.round(t.count).toLocaleString(), 12) + ' shipments, £' +
               t.cost.toFixed(2));
  });
  if (blankTreat) Logger.log('  ' + blankTreat + ' row(s) have a blank Treatment Type.');
  Logger.log('');
  Logger.log('  Of the Core Rx rows, ' + silent + ' (£' + silentCost.toFixed(2) +
             ') land on a segment that exists.');
  Logger.log('  Those are the ones that would be absorbed silently if any of them is');
  Logger.log('  really weight loss. The rest have no segment and are already reported');
  Logger.log('  by previewActualsImport as unmatched.');

  const orphaned = list.filter(function (e) { return e.noSeg > 0; });
  if (orphaned.length) {
    Logger.log('');
    Logger.log('--- extract rows with no High Level ID to receive them ---');
    orphaned.forEach(function (e) {
      Logger.log('  ' + e.norm + '  (bucketed as ' + e.bucket + ')');
      Object.keys(e.noSegWhere).sort().forEach(function (k) {
        const w = e.noSegWhere[k];
        Logger.log('      ' + pad_(k, 22) + '  ' + pad_(String(w.rows), 5) + ' row(s), ' +
                   pad_(Math.round(w.count).toLocaleString(), 12) + ' shipments, £' +
                   w.cost.toFixed(2));
      });
    });
    Logger.log('  previewActualsImport already reports these as unmatched. They are');
    Logger.log('  dropped, not misfiled — but a WL row here is volume the forecast');
    Logger.log('  never sees, which looks the same as a segment that is simply quiet.');
  }

  // ---- segments the importer can never reach ------------------------------
  const ref = getAllData_(SHEET.DIM_REFERENCE), R = COL.DIM_REFERENCE;
  const defined = [];
  for (let i = 1; i < ref.length; i++) {
    if (normKey(ref[i][R.List_Name]) !== 'TREATMENT_TYPE') continue;
    if (!safeBool(ref[i][R.Active])) continue;
    defined.push(safeStr(ref[i][R.Code]));
  }
  Logger.log('');
  Logger.log('--- High Level IDs the import can never reach ---');
  Logger.log('  Dim_Reference TREATMENT_TYPE defines: ' + (defined.join(', ') || '(none)'));
  if (!oddSegments.length) {
    Logger.log('  Every High Level ID is WL or CORE_RX. Nothing wrong here.');
  } else {
    oddSegments.forEach(function (o) {
      Logger.log('  ' + pad_(String(o.id), 4) + '  ' + o.brand + ' ' + o.geo + ' ' +
                 o.tt + (o.split ? ' (' + o.split + ')' : '') +
                 (o.active ? '' : '  [inactive]'));
    });
    Logger.log('');
    Logger.log('  The importer only ever produces WL or CORE_RX, so no extract row can');
    Logger.log('  key to these. They will stay empty of actuals until either the');
    Logger.log('  Treatment_Type is corrected in High_Level_IDs, or the value becomes a');
    Logger.log('  real third classification that the import knows how to produce.');
  }

  Logger.log('');
  Logger.log('  Nothing was written.');
  return { values: list.length, rows: readRows,
           wlValues: tot.WL.v, coreValues: tot.CORE_RX.v,
           absorbedRows: silent, absorbedCost: silentCost,
           oddSegments: oddSegments.length };
}


/**
 * The extract broken down by Delivery Carrier and Delivery Method, and one
 * brand/geo/treatment shown month by month.
 *
 * Reads only. Every other diagnostic here aggregates carrier and method away,
 * because the forecast is a blended rate per order and the actual has to be
 * blended the same way. That is right for comparing, and useless for explaining:
 * when a segment's measured rate sits far below its forecast, the two candidates
 * are a cost column that is missing rows and a count column that is counting
 * something other than parcels — and carrier and method are the columns that
 * tell those apart. A month where the carrier is blank and the cost is a
 * rounding error is a gap in how the extract was compiled. A month where the
 * carrier is real and the cost is still small is a gap in what it was charged.
 *
 * With no arguments it summarises the whole extract and then drills into the
 * brand/geo/treatment carrying the most shipments, which is where an
 * order-of-magnitude error is worth the most. Pass all three to choose another;
 * matching ignores case and punctuation, as everywhere else in this file.
 */
function diagnoseActualsExtract(brand, geo, treatment) {
  requireMaintenance_();
  Logger.log('=== ACTUALS IMPORT — CARRIER AND METHOD (read-only) ===');
  prewarmSheetCache_([SHEET.ACTUALS_IMPORT, SHEET.CONFIG]);

  const rows = getAllData_(SHEET.ACTUALS_IMPORT);
  if (rows.length < 2) {
    Logger.log('  The Actuals_Import tab is empty — paste the extract there first.');
    return { ok: false, reason: 'empty' };
  }

  /* Carrier and method are optional rather than required: an older extract that
     predates them should still give the summaries below, not an exception. */
  const X = extractStandardColumns_(rows[0]);
  const hdrRaw = X.cols.raw;
  const cCountry = X.cCountry, cBrand = X.cBrand, cTreat = X.cTreat,
        cMonth = X.cMonth, cCount = X.cCount, cCost = X.cCost,
        cCarrier = X.cCarrier, cMethod = X.cMethod;
  Logger.log('  carrier column : ' + (cCarrier >= 0 ? '"' + hdrRaw[cCarrier] + '"' : 'NOT PRESENT'));
  Logger.log('  method column  : ' + (cMethod  >= 0 ? '"' + hdrRaw[cMethod]  + '"' : 'NOT PRESENT'));
  if (cCarrier < 0 && cMethod < 0) {
    Logger.log('  Neither column is in this extract, so there is nothing to break down.');
    return { ok: false, reason: 'no carrier or method column' };
  }

  /* A separator that cannot occur inside a carrier or method name, built
     rather than written, so no control character sits in the source. */
  const SEP = String.fromCharCode(1);
  function cell(r, i) { return i >= 0 ? (safeStr(r[i]) || '(blank)') : '(n/a)'; }
  function bump(map, key, n, cost) {
    const e = map[key] || (map[key] = { rows: 0, count: 0, cost: 0 });
    e.rows++; e.count += n; e.cost += cost;
  }
  function table(title, map) {
    const keys = Object.keys(map).sort(function (a, b) { return map[b].cost - map[a].cost; });
    Logger.log('');
    Logger.log('--- ' + title + ' ---');
    Logger.log('  ' + pad_('rows', 6) + '  ' + pad_('shipments', 12) + '  ' +
               pad_('cost', 14) + '  ' + pad_('£/ship', 9) + '  value');
    keys.forEach(function (k) {
      const e = map[k];
      Logger.log('  ' + pad_(String(e.rows), 6) + '  ' +
                 pad_(Math.round(e.count).toLocaleString(), 12) + '  ' +
                 pad_('£' + e.cost.toFixed(2), 14) + '  ' +
                 pad_('£' + (e.count ? (e.cost / e.count).toFixed(4) : '0.0000'), 9) + '  ' + k);
    });
  }

  // ---- pass one: the whole extract ---------------------------------------
  const byCarrier = {}, byMethod = {}, combos = {};
  let usable = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const b = normKey(r[cBrand]), g = normKey(r[cCountry]);
    if (!b || !g) continue;
    const n = parseExtractNumber_(r[cCount]);
    if (!n) continue;
    usable++;
    const cost = parseExtractNumber_(r[cCost]);
    bump(byCarrier, cell(r, cCarrier), n, cost);
    bump(byMethod,  cell(r, cMethod),  n, cost);
    const t = safeStr(r[cTreat]).toUpperCase().replace(/[^A-Z0-9]/g, '');
    bump(combos, b + '|' + g + '|' + t, n, cost);
  }
  Logger.log('  usable rows    : ' + usable);
  table('by Delivery Carrier', byCarrier);
  table('by Delivery Method',  byMethod);

  // ---- choose what to drill into -----------------------------------------
  const wantB = normKey(brand), wantG = normKey(geo);
  const wantT = safeStr(treatment).toUpperCase().replace(/[^A-Z0-9]/g, '');
  let target = '';
  if (wantB && wantG && wantT) {
    target = wantB + '|' + wantG + '|' + wantT;
    if (!combos[target]) {
      Logger.log('');
      Logger.log('  No rows for ' + wantB + ' ' + wantG + ' ' + wantT + '.');
      return { ok: true, rows: usable, target: null };
    }
  } else {
    Object.keys(combos).forEach(function (k) {
      if (!target || combos[k].count > combos[target].count) target = k;
    });
    Logger.log('');
    Logger.log('  No brand/geo/treatment given, so drilling into the combination');
    Logger.log('  carrying the most shipments. Pass all three to choose another.');
  }

  // ---- pass two: that combination, month by month ------------------------
  const cells = {}, months = {}, monthTot = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const b = normKey(r[cBrand]), g = normKey(r[cCountry]);
    if (!b || !g) continue;
    const n = parseExtractNumber_(r[cCount]);
    if (!n) continue;
    const t = safeStr(r[cTreat]).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (b + '|' + g + '|' + t !== target) continue;

    const ms = parseExtractMonth_(r[cMonth]);
    const mk = ms ? fmtDate(ms) : '(unparsed)';
    months[mk] = ms ? dateKey(ms) : 0;
    const cost = parseExtractNumber_(r[cCost]);
    bump(cells, mk + SEP + cell(r, cCarrier) + SEP + cell(r, cMethod), n, cost);
    bump(monthTot, mk, n, cost);
  }

  const parts = target.split('|');
  Logger.log('');
  Logger.log('--- ' + parts[0] + ' ' + parts[1] + ' ' + parts[2] + ', month by month ---');
  Logger.log('  ' + pad_('month', 11) + '  ' + pad_('rows', 5) + '  ' +
             pad_('shipments', 12) + '  ' + pad_('cost', 13) + '  ' +
             pad_('£/ship', 9) + '  carrier / method');
  Object.keys(months).sort(function (a, b) { return months[a] - months[b]; })
    .forEach(function (mk) {
      Object.keys(cells).filter(function (k) { return k.split(SEP)[0] === mk; })
        .sort(function (a, b) { return cells[b].cost - cells[a].cost; })
        .forEach(function (k) {
          const e = cells[k], p = k.split(SEP);
          Logger.log('  ' + pad_(mk, 11) + '  ' + pad_(String(e.rows), 5) + '  ' +
                     pad_(Math.round(e.count).toLocaleString(), 12) + '  ' +
                     pad_('£' + e.cost.toFixed(2), 13) + '  ' +
                     pad_('£' + (e.count ? (e.cost / e.count).toFixed(4) : '0.0000'), 9) + '  ' +
                     p[1] + ' / ' + p[2]);
        });
      const m = monthTot[mk];
      Logger.log('  ' + pad_(mk, 11) + '  ' + pad_(String(m.rows), 5) + '  ' +
                 pad_(Math.round(m.count).toLocaleString(), 12) + '  ' +
                 pad_('£' + m.cost.toFixed(2), 13) + '  ' +
                 pad_('£' + (m.count ? (m.cost / m.count).toFixed(4) : '0.0000'), 9) +
                 '  == month total');
    });

  Logger.log('');
  Logger.log('  A month whose carrier or method is (blank) was compiled without one.');
  Logger.log('  If the blank months are also the near-zero-cost months, the cost was');
  Logger.log('  never in the extract, and no change to this code can recover it.');
  Logger.log('');
  Logger.log('  Nothing was written.');
  return { ok: true, rows: usable, target: target,
           carriers: Object.keys(byCarrier).length,
           methods: Object.keys(byMethod).length,
           months: Object.keys(months).length };
}


/**
 * Several segments' actuals against their forecast, side by side.
 *
 * Reads only. diagnoseActuals shows one segment in detail, and the Run menu
 * cannot pass it an argument, so comparing segments means editing the default
 * and running it again — which is how you end up comparing two logs from
 * memory. This runs a fixed list and finishes with one table, because the
 * question is not "what does segment 1 look like" but "is segment 1 unusual".
 *
 * The figure that carries the answer is REPORTED %: the cost the extract holds
 * as a percentage of what the forecast says that many orders should have cost.
 * Near 100 means the extract agrees with the model. Near zero means the cost is
 * not there — and whether that is a portal problem or an extract problem is
 * decided by how the OTHER segments read, not this one.
 *
 * Edit SEGMENTS to compare a different set.
 */
function diagnoseActualsSegments() {
  requireMaintenance_();
  Logger.log('=== ACTUALS vs FORECAST, SEVERAL SEGMENTS (read-only) ===');

  /* Chosen to separate four explanations that look identical from one segment:
     1 and 11 share a brand and country and differ by treatment; 1, 5 and 6 share
     a brand and treatment and differ by country. Whichever dimension the anomaly
     follows names the team that owns it. */
  const SEGMENTS = [1, 5, 6, 11];

  prewarmSheetCache_([SHEET.ACTUALS, SHEET.HIGH_LEVEL_IDS, SHEET.OUTPUT, SHEET.CONFIG]);

  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  const name = {};
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (!id) continue;
    name[id] = safeStr(hl[i][H.Brand]) + ' ' + safeStr(hl[i][H.Geo]) + ' ' +
               safeStr(hl[i][H.Treatment_Type]) +
               (safeStr(hl[i][H.WL_Split]) ? ' ' + safeStr(hl[i][H.WL_Split]) : '');
  }

  const out = getAllData_(SHEET.OUTPUT), O = COL.OUTPUT;
  const fc = {};
  for (let i = 1; i < out.length; i++) {
    fc[safeInt(out[i][O.High_Level_ID]) + '|' + fmtDate(out[i][O.Month_Start])] =
      safeNum(out[i][O.Forecast_Rate_Per_Order]);
  }

  const act = getAllData_(SHEET.ACTUALS), A = COL.ACTUALS;
  const summary = [];

  SEGMENTS.forEach(function (hlId) {
    Logger.log('');
    Logger.log('--- ' + hlId + '  ' + (name[hlId] || '(no such High Level ID)') + ' ---');

    const months = [];
    for (let i = 1; i < act.length; i++) {
      if (safeInt(act[i][A.High_Level_ID]) !== hlId) continue;
      if (!safeBool(act[i][A.Active])) continue;
      months.push({ month: fmtDate(act[i][A.Month_Start]),
                    orders: safeNum(act[i][A.Orders]),
                    spend: safeNum(act[i][A.Total_Spend]),
                    rate: safeNum(act[i][A.Blended_Rate]) });
    }
    months.sort(function (a, b) { return a.month < b.month ? -1 : a.month > b.month ? 1 : 0; });

    if (!months.length) {
      Logger.log('  No actuals on record.');
      summary.push({ id: hlId, months: 0, orders: 0, spend: 0, expected: 0 });
      return;
    }

    Logger.log('  ' + pad_('month', 11) + '  ' + pad_('orders', 10) + '  ' +
               pad_('spend', 13) + '  ' + pad_('rate', 9) + '  ' +
               pad_('forecast', 9) + '  ' + pad_('reported', 9));
    let to = 0, ts = 0, te = 0;
    months.forEach(function (m) {
      const f = fc[hlId + '|' + m.month] || 0;
      /* Orders on the second half of a split pair are deliberately zero — the
         importer puts shipments and spend on the lower ID only, so nothing is
         counted twice. A row like that is shown but contributes no expectation. */
      const exp = m.orders * f;
      to += m.orders; ts += m.spend; te += exp;
      Logger.log('  ' + pad_(m.month, 11) + '  ' +
                 pad_(Math.round(m.orders).toLocaleString(), 10) + '  ' +
                 pad_('£' + m.spend.toFixed(2), 13) + '  ' +
                 pad_('£' + m.rate.toFixed(4), 9) + '  ' +
                 pad_(f ? '£' + f.toFixed(4) : '-', 9) + '  ' +
                 pad_(exp ? ((m.spend / exp) * 100).toFixed(2) + '%' : '-', 9));
    });
    Logger.log('  ' + pad_('TOTAL', 11) + '  ' +
               pad_(Math.round(to).toLocaleString(), 10) + '  ' +
               pad_('£' + ts.toFixed(2), 13) + '  ' +
               pad_(to ? '£' + (ts / to).toFixed(4) : '-', 9) + '  ' +
               pad_('', 9) + '  ' +
               pad_(te ? ((ts / te) * 100).toFixed(2) + '%' : '-', 9));
    if (te) {
      Logger.log('  expected at forecast: £' + te.toFixed(2) +
                 '   shortfall £' + (te - ts).toFixed(2));
    }
    summary.push({ id: hlId, months: months.length, orders: to, spend: ts, expected: te });
  });

  Logger.log('');
  Logger.log('--- side by side ---');
  Logger.log('  ' + pad_('id', 4) + '  ' + pad_('reported', 9) + '  ' +
             pad_('shortfall', 14) + '  segment');
  summary.forEach(function (s) {
    Logger.log('  ' + pad_(String(s.id), 4) + '  ' +
               pad_(s.expected ? ((s.spend / s.expected) * 100).toFixed(2) + '%' : '-', 9) + '  ' +
               pad_(s.expected ? '£' + (s.expected - s.spend).toFixed(2) : '-', 14) + '  ' +
               (name[s.id] || '?'));
  });
  Logger.log('');
  Logger.log('  Reading it: if only 1 is low, the problem is that segment. If 1, 5 and 6');
  Logger.log('  are low and 11 is not, it follows weight loss. If all four are low, it is');
  Logger.log('  the extract as a whole. If 1 and 11 are low but 5 and 6 are not, it');
  Logger.log('  follows the country.');
  Logger.log('');
  Logger.log('  Nothing was written.');
  return { segments: summary };
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


/**
 * Empty the staging tab, so a REPLACE upload starts from nothing.
 *
 * A different tab from clearActuals(), and worth being clear about which is
 * which, because the names are one word apart and the consequences are not.
 * clearActuals() removes the derived Actuals rows — the ones the portal shows
 * against the forecast — and those have an Amends history, so removing them is
 * recoverable. This removes the raw extract in Actuals_Import, which has no
 * history table at all: the staging tab is a scratch space by design, so what is
 * cleared here is gone unless the original file still exists somewhere.
 *
 * Nothing derived is touched. The Actuals rows a previous import produced stay
 * exactly as they are until an import runs again and overwrites them by
 * segment-month.
 */
function clearActualsImport_() {
  const before = Math.max(getSheet_(SHEET.ACTUALS_IMPORT).getLastRow() - 1, 0);
  clearDataRows_(SHEET.ACTUALS_IMPORT);
  logAudit_('DELETE', SHEET.ACTUALS_IMPORT, '', '', '',
            String(before) + ' rows', 'clearActualsImport', true);
  return before;
}

/** The same, from the editor. */
function clearActualsImport() {
  requireMaintenance_();
  const before = clearActualsImport_();
  Logger.log('Removed ' + before + ' staged extract row(s) from ' + SHEET.ACTUALS_IMPORT + '.');
  Logger.log('The Actuals tab itself is untouched — use clearActuals() for that.');
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


/**
 * The extract's columns, located by name rather than by position.
 *
 * One matcher, used by the importer and by every diagnostic that reads the same
 * staging tab. It was four near-copies of the same twenty lines, which is the
 * worst possible arrangement for this particular job: the whole point of the
 * diagnostics is to tell you what the importer will do, and they can only do
 * that if they find columns exactly the way the importer finds them. A tightened
 * candidate list in one copy and not the others makes the diagnostic disagree
 * with the import and there is no way to tell from the output which one is right.
 *
 * Matching is two-pass, in this order, and the order matters:
 *
 *   1. Exact, on the normalised name. Every candidate is tried before any
 *      fuzzy match is considered, so a file that has both "Count" and
 *      "Count Of Orders" binds to "Count" rather than to whichever happened to
 *      be leftmost.
 *   2. Substring, either direction. This is what survives "Sum of Cost Of
 *      Shipping (£)" — the pound sign is two bytes in UTF-8 and comes back as
 *      "Â£" whenever something in the chain decodes it as Latin-1, and the
 *      trailing unit differs between exports.
 *
 * A column, once bound, is CLAIMED, and no later label may bind to it. That is
 * not tidiness, it is the fix for a specific way this used to go wrong: the
 * reverse half of pass 2 lets a candidate contain a header, and "COUNTRYCODE"
 * contains "COUNT". So an extract whose country column was named something
 * unrecognised — "Geo", say — would bind country to the SHIPMENT COUNT column,
 * pass every remaining check, and import with every row's country set to a
 * number. Claiming makes that impossible: Count is bound exactly, before any
 * fuzzy match is tried, and country then finds nothing and says so.
 *
 * Because pass 2 can bind a column on a partial name, `how()` records which pass
 * found each one. Anything matched fuzzily is worth printing before its numbers
 * are believed — a cost column found by substring may be the wrong cost column,
 * and the import cannot tell.
 *
 * @param {Array} headerRow  row 0 of the staging tab or CSV
 * @return {Object} { raw, norm, find, need, resolve, how, matched }
 */
function extractColumns_(headerRow) {
  const raw  = (headerRow || []).map(function (h) { return safeStr(h); });
  const norm = raw.map(normHeader_);
  const how  = {};
  const claimed = {};

  function exactIndex(candidates) {
    for (let ci = 0; ci < candidates.length; ci++) {
      const want = normHeader_(candidates[ci]);
      if (!want) continue;
      for (let hi = 0; hi < norm.length; hi++) {
        if (claimed[hi]) continue;
        if (norm[hi] === want) return { index: hi, candidate: candidates[ci] };
      }
    }
    return null;
  }

  function fuzzyIndex(candidates) {
    for (let ci = 0; ci < candidates.length; ci++) {
      const want = normHeader_(candidates[ci]);
      if (!want) continue;
      for (let hi = 0; hi < norm.length; hi++) {
        if (claimed[hi] || !norm[hi]) continue;
        if (norm[hi].indexOf(want) >= 0 || want.indexOf(norm[hi]) >= 0) {
          return { index: hi, candidate: candidates[ci] };
        }
      }
    }
    return null;
  }

  function bind(hit, label, mode) {
    claimed[hit.index] = true;
    if (label) how[label] = { index: hit.index, mode: mode, candidate: hit.candidate };
    return hit.index;
  }

  function missing(candidates, label) {
    return new Error('The extract has no ' + label + ' column. It needs one named ' +
      'something like "' + candidates[0] + '". Columns found: ' +
      (raw.length ? raw.join(' | ') : '(none — the file has no header row)'));
  }

  /** Index, or -1. Optional columns go through this one. */
  function find(candidates, label) {
    const e = exactIndex(candidates);
    if (e) return bind(e, label, 'exact');
    const f = fuzzyIndex(candidates);
    if (f) return bind(f, label, 'fuzzy');
    return -1;
  }

  /** Index, or an error naming what was wanted and what was actually there. */
  function need(candidates, label) {
    const i = find(candidates, label);
    if (i < 0) throw missing(candidates, label);
    return i;
  }

  /**
   * Resolve several columns together: every exact match claimed first, across
   * all of them, and only then the fuzzy fallbacks. Resolving one at a time
   * cannot get this right, because whichever column is asked for first gets to
   * fuzzy-match before the others have claimed their exact headers.
   *
   * @param {Array} specs  [{ label, candidates, required }]
   * @return {Object} label -> index (-1 for an optional column that is absent)
   */
  function resolve(specs) {
    const out = {}, pending = [];
    specs.forEach(function (s) {
      const e = exactIndex(s.candidates);
      if (e) out[s.label] = bind(e, s.label, 'exact');
      else pending.push(s);
    });
    pending.forEach(function (s) {
      const f = fuzzyIndex(s.candidates);
      if (f) out[s.label] = bind(f, s.label, 'fuzzy');
      else if (s.required) throw missing(s.candidates, s.label);
      else out[s.label] = -1;
    });
    return out;
  }

  return {
    raw: raw, norm: norm, find: find, need: need, resolve: resolve,
    how: function (label) { return how[label] || null; },
    /** [[label, index, mode, header], ...] for the labels resolved so far. */
    matched: function () {
      return Object.keys(how).map(function (k) {
        return [k, how[k].index, how[k].mode, raw[how[k].index]];
      });
    }
  };
}


/**
 * The six columns every reader of the extract needs, plus the two it may not
 * have. Kept here so adding an accepted spelling of a header is one edit, and
 * so the importer and the diagnostics cannot drift apart on what they accept.
 *
 * Carrier and method are optional on purpose: the import aggregates them away
 * anyway, and an older extract that predates the columns should still load
 * rather than throw.
 */
function extractStandardColumns_(headerRow) {
  const c = extractColumns_(headerRow);
  /* Order within this list only decides which of two equally-good fuzzy matches
     wins; every exact match is claimed before any of them is tried. The narrow
     names are first anyway, so a genuinely ambiguous file resolves the specific
     column before the vague one. */
  const r = c.resolve([
    { label: 'shipment count', candidates: ['Count', 'Shipments', 'Orders'], required: true },
    { label: 'shipping cost',  candidates: ['Sum of Cost Of Shipping', 'Cost Of Shipping',
                                            'Total Cost', 'Cost'], required: true },
    { label: 'country',        candidates: ['Country Code', 'Country'], required: true },
    { label: 'brand',          candidates: ['Brand'], required: true },
    { label: 'treatment type', candidates: ['Treatment Type', 'Treatment'], required: true },
    { label: 'month',          candidates: ['Dispatched Date: Month', 'Dispatched Month',
                                            'Month'], required: true },
    { label: 'carrier',        candidates: ['Delivery Carrier', 'Carrier'], required: false },
    { label: 'method',         candidates: ['Delivery Method', 'Method', 'Service'],
                               required: false }
  ]);
  return {
    cols:     c,
    cCountry: r['country'],
    cBrand:   r['brand'],
    cTreat:   r['treatment type'],
    cMonth:   r['month'],
    cCount:   r['shipment count'],
    cCost:    r['shipping cost'],
    cCarrier: r['carrier'],
    cMethod:  r['method']
  };
}


/** The matched columns, and how each was found, into the log. */
function logExtractColumns_(cols) {
  Logger.log('  columns matched:');
  cols.matched().forEach(function (m) {
    Logger.log('    ' + pad_(m[0], 10) + '  ' + pad_(m[2], 6) + '  "' + m[3] + '"');
  });
  const fuzzy = cols.matched().filter(function (m) { return m[2] === 'fuzzy'; });
  if (fuzzy.length) {
    Logger.log('  ** ' + fuzzy.map(function (m) { return m[0]; }).join(', ') +
               ' matched by substring, not by name. Confirm the columns above are');
    Logger.log('     the right ones before believing the numbers below. **');
  }
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
 *
 * The gate and the work are split so the portal's CSV upload can reach the same
 * calculation without inheriting this one's permission check. The upload has its
 * own, at least as strict (see requireImportActuals_ in auth.gs), and if it
 * called this wrapper it would be checked twice against two different rules —
 * the reliable way to end up with a button that is authorised by one rule and
 * refused by the other. There is deliberately no second aggregation path: both
 * entry points run importActualsFromStagingCore_, so an actual row created by
 * the upload is created by the same code that creates it from the editor.
 *
 * @param {boolean} commit
 */
function importActualsFromStaging(commit) {
  requireMaintenance_();
  return importActualsFromStagingCore_(commit);
}


/**
 * The import itself. Assumes the caller has already established the right to run
 * it, and returns a report structured for a screen as well as logged for a human.
 *
 * @param {boolean} commit
 */
function importActualsFromStagingCore_(commit) {
  Logger.log(commit ? '=== IMPORTING ACTUALS ===' : '=== ACTUALS IMPORT PREVIEW (nothing written) ===');

  prewarmSheetCache_([SHEET.ACTUALS_IMPORT, SHEET.HIGH_LEVEL_IDS, SHEET.ACTUALS,
                      SHEET.DIM_CALENDAR, SHEET.CONFIG]);

  const rows = getAllData_(SHEET.ACTUALS_IMPORT);
  if (rows.length < 2) {
    Logger.log('  The Actuals_Import tab is empty.');
    Logger.log('  Paste the extract there, including its header row, then run this again.');
    return { ok: false, reason: 'empty' };
  }

  // Columns are found by name, so a reordered extract still works — see
  // extractColumns_ for why the match is deliberately forgiving.
  const X = extractStandardColumns_(rows[0]);
  const cCountry = X.cCountry, cBrand = X.cBrand, cTreat = X.cTreat,
        cMonth = X.cMonth, cCount = X.cCount, cCost = X.cCost;
  logExtractColumns_(X.cols);

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
  const agg = {}, unmatched = {}, outside = {}, defaulted = {};
  let read = 0, skipped = 0;

  /* Why each skipped row was skipped, not just how many there were.
     "412 rows skipped as blank" is the same message whether the extract had 412
     genuinely empty trailing rows or 412 rows whose month failed to parse, and
     those are an ignorable non-event and a silently halved forecast
     respectively. Counted by reason, with the first few row numbers of each, so
     the difference is visible without opening the tab. */
  const skips = {};
  function skip(reason, rowIndex) {
    skipped++;
    const s = skips[reason] || (skips[reason] = { reason: reason, rows: 0, examples: [] });
    s.rows++;
    if (s.examples.length < 5) s.examples.push(rowIndex + 1);   // 1-based, as the sheet shows it
  }

  /* Rows carrying shipments but no money. Same symptom as a mis-bucketed
     treatment — a blended rate far under forecast with a shipment count that is
     perfectly correct — and a completely different cause, so it is reported
     separately rather than folded into the totals. The parse cannot distinguish
     a genuine zero from a cost that was never exported, so this reports the
     shape of the problem and leaves the judgement to a person. */
  const zeroCost = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const brand = normKey(r[cBrand]), geo = normKey(r[cCountry]);
    if (!brand || !geo) { skip('no brand or country', i); continue; }

    const treatRaw = safeStr(r[cTreat]);
    const treatKey = treatRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const treat = wl[treatKey] ? 'WL' : 'CORE_RX';
    const ms = parseExtractMonth_(r[cMonth]);
    if (!ms) { skip('month could not be read', i); continue; }

    const n = parseExtractNumber_(r[cCount]);
    const cost = parseExtractNumber_(r[cCost]);
    if (!n) { skip('no shipment count', i); continue; }
    read++;

    if (cost === 0) {
      const zk = brand + '|' + geo + '|' + treat + '|' + dateKey(ms);
      const z = zeroCost[zk] ||
                (zeroCost[zk] = { brand: brand, geo: geo, treat: treat,
                                  month: fmtDate(ms), rows: 0, ship: 0, examples: [] });
      z.rows++; z.ship += n;
      if (z.examples.length < 3) z.examples.push(i + 1);
    }

    /* Core Rx is a DEFAULT, not a decision: every treatment string that is not
       in ACTUALS_WL_TREATMENTS lands here, whether or not anyone has ever looked
       at it. A weight-loss value missing from that config is therefore not
       dropped — it is added to the Core Rx segment for the same brand and
       country, moving that segment's volume AND its blended rate, since the rate
       is cost divided by shipments for the whole segment-month. Both sides then
       fail to reconcile at once.

       Nothing in the sheet records which of ~50 conditions ought to be weight
       loss, so this cannot be validated — but it can refuse to be quiet. Every
       defaulted value is counted here and printed below with its money, the way
       RATE_MISSING reports a route priced at zero rather than letting the
       forecast come out low with nothing said. */
    if (treat === 'CORE_RX') {
      const df = defaulted[treatKey] ||
                 (defaulted[treatKey] = { raw: treatRaw, rows: 0, ship: 0, cost: 0 });
      df.rows++; df.ship += n; df.cost += cost;
    }

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

  Logger.log('  extract rows read      : ' + read + (skipped ? '  (' + skipped + ' skipped)' : ''));
  Logger.log('  segment-months to load : ' + Object.keys(agg).length);

  const skipList = Object.keys(skips).map(function (k) { return skips[k]; });
  skipList.sort(function (a, b) { return b.rows - a.rows; });
  if (skipList.length) {
    Logger.log('');
    Logger.log('--- ' + skipped + ' row(s) skipped ---');
    skipList.forEach(function (s) {
      Logger.log('  ' + pad_(String(s.rows), 6) + '  ' + pad_(s.reason, 24) +
                 '  e.g. sheet row ' + s.examples.join(', '));
    });
    Logger.log('  Trailing blank rows are normal. A month that could not be read is not:');
    Logger.log('  those shipments are missing from the comparison entirely.');
  }

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

  // ---- what fell through to Core Rx without being asked about -------------
  const dfList = Object.keys(defaulted).map(function (k) { return defaulted[k]; });
  dfList.sort(function (a, b) { return b.cost - a.cost; });
  let dfRows = 0, dfCost = 0;
  dfList.forEach(function (d) { dfRows += d.rows; dfCost += d.cost; });
  if (dfList.length) {
    Logger.log('');
    Logger.log('--- ' + dfList.length + ' treatment value(s) fell through to Core Rx by ' +
               'default: ' + dfRows + ' row(s), £' + dfCost.toFixed(2) + ' ---');
    Logger.log('  These matched nothing in ACTUALS_WL_TREATMENTS. That is correct for a');
    Logger.log('  genuine Core Rx condition and wrong for a weight-loss one, and nothing');
    Logger.log('  in the sheet says which is which. Read the list.');
    Logger.log('  ' + pad_('rows', 6) + '  ' + pad_('shipments', 12) + '  ' +
               pad_('cost', 14) + '  ' + pad_('£/ship', 9) + '  treatment type');
    dfList.forEach(function (d) {
      Logger.log('  ' + pad_(String(d.rows), 6) + '  ' +
                 pad_(Math.round(d.ship).toLocaleString(), 12) + '  ' +
                 pad_('£' + d.cost.toFixed(2), 14) + '  ' +
                 pad_('£' + (d.ship ? (d.cost / d.ship).toFixed(3) : '0.000'), 9) + '  ' +
                 (d.raw === '' ? '(blank)' : d.raw));
    });
    Logger.log('  diagnoseActualsTreatments() shows the same list with where each row');
    Logger.log('  would land. Add any weight-loss value to Config ACTUALS_WL_TREATMENTS');
    Logger.log('  and run the preview again before committing.');
  }

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

  // ---- shipments that arrived with no money -------------------------------
  const zcList = Object.keys(zeroCost).map(function (k) { return zeroCost[k]; });
  zcList.sort(function (a, b) { return b.ship - a.ship; });
  let zcRows = 0, zcShip = 0;
  zcList.forEach(function (z) { zcRows += z.rows; zcShip += z.ship; });
  if (zcList.length) {
    Logger.log('');
    Logger.log('--- ' + zcRows + ' row(s) carried shipments but a cost of 0: ' +
               Math.round(zcShip).toLocaleString() + ' shipments ---');
    Logger.log('  These still count towards the blended rate, as a divisor with nothing');
    Logger.log('  on top, so they pull the measured rate down. A genuine zero and a cost');
    Logger.log('  that never made it into the export look identical here.');
    Logger.log('  ' + pad_('rows', 6) + '  ' + pad_('shipments', 12) + '  segment-month');
    zcList.slice(0, 20).forEach(function (z) {
      Logger.log('  ' + pad_(String(z.rows), 6) + '  ' +
                 pad_(Math.round(z.ship).toLocaleString(), 12) + '  ' +
                 z.brand + ' ' + z.geo + ' ' + z.treat.replace('_', ' ') + ' ' + z.month +
                 '   e.g. sheet row ' + z.examples.join(', '));
    });
    if (zcList.length > 20) {
      Logger.log('  ... and ' + (zcList.length - 20) + ' more segment-month(s).');
    }
  }

  if (Object.keys(outside).length) {
    let n = 0;
    Object.keys(outside).forEach(function (k) { n += outside[k].ship; });
    Logger.log('');
    Logger.log('--- ' + Object.keys(outside).length + ' month(s) outside the forecast horizon, ' +
               Math.round(n).toLocaleString() + ' shipments, ignored ---');
  }

  /* One report shape for both outcomes, so a caller rendering it does not have
     to know whether it previewed or committed — only `written` differs. The
     lists are capped where a screen could not show more anyway; the log above is
     the complete record. */
  const outsideList = Object.keys(outside).sort().map(function (k) {
    return { month: k, ship: outside[k].ship };
  });
  const unmatchedList = Object.keys(unmatched).sort().map(function (k) {
    const p = k.split('|');
    return { brand: p[0], geo: p[1], treat: p[2].replace('_', ' '),
             ship: unmatched[k].ship, cost: unmatched[k].cost };
  });
  const report = {
    ok: true,
    preview: !commit,
    rowsRead: read,
    rowsSkipped: skipped,
    skipReasons: skipList.map(function (s) {
      return { reason: s.reason, rows: s.rows, examples: s.examples };
    }),
    segmentMonths: Object.keys(agg).length,
    rows: toWrite.length,
    unmatched: unmatchedList.length,
    unmatchedSegments: unmatchedList.slice(0, 50),
    defaultedValues: dfList.length,
    defaultedRows: dfRows,
    defaultedCost: dfCost,
    defaultedTreatments: dfList.slice(0, 50).map(function (d) {
      return { treatment: d.raw === '' ? '(blank)' : d.raw,
               rows: d.rows, ship: d.ship, cost: d.cost };
    }),
    zeroCostRows: zcRows,
    zeroCostShipments: zcShip,
    zeroCost: zcList.slice(0, 50).map(function (z) {
      return { brand: z.brand, geo: z.geo, treat: z.treat.replace('_', ' '),
               month: z.month, rows: z.rows, ship: z.ship };
    }),
    outsideHorizon: outsideList
  };

  if (!commit) {
    Logger.log('');
    Logger.log('  Nothing written. Run importActualsFromStaging(true) to load them.');
    return report;
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

  report.written = res.written;
  report.writeSkipped = res.skipped;
  /* Whatever importActuals could not write, verbatim. Almost always a scope
     refusal — the import touches every segment, and saveActual still checks each
     one against Scope_Mapping — and a row refused for that reason is missing
     from the comparison with nothing on screen to say so. */
  report.writeErrors = res.errors.slice(0, 10);
  return report;
}


/** Convenience wrappers, so both appear in the editor's function list. */
function previewActualsImport() { return importActualsFromStaging(false); }
function runActualsImport()     { return importActualsFromStaging(true); }


// ─────────────────────────────────────────────────────────────────────────────
// UPLOADING THE EXTRACT FROM THE PORTAL
//
// The same import, reached from a file picker instead of a paste into the
// staging tab followed by a run from the Apps Script editor. What it replaces is
// four manual steps with three places to go wrong: open the workbook, paste into
// the right tab, open the editor, run the right function — and then read a log
// that nothing forces you to read.
//
// There is no preview step here, by request, so the report is the only thing
// standing between a bad file and a wrong forecast. That is why the report is
// returned as data rather than only logged: the manual route has caught real
// problems (a treatment type quietly defaulting to Core Rx, a cost column that
// did not parse) and it caught them because a person happened to read the log.
// A screen that says nothing is worse than a log nobody opens, because it looks
// like a clean run.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * The largest file this will accept.
 *
 * Not a performance limit — the staging write is one setValues call regardless.
 * It is a limit on how much can cross google.script.run in one argument and how
 * many rows importActuals can get through inside the 6-minute ceiling, since
 * that part is one saveActual per segment-month. A monthly extract is a few
 * thousand rows; something an order of magnitude past that is a wrong file, and
 * failing on it with a sentence is better than timing out half-written.
 */
const ACTUALS_UPLOAD_MAX_ROWS = 20000;

/**
 * Load a CSV of the monthly extract, then run the import.
 *
 * Two modes, and the difference is only what happens to Actuals_Import first:
 *
 *   APPEND  — add this file's months to what is already staged, having first
 *             removed any rows already staged for the months the file covers.
 *   REPLACE — empty the staging tab and stage only this file.
 *
 * APPEND does not simply append, and it must not. The aggregation sums every
 * staged row for a segment-month, so a month uploaded twice would produce a
 * segment-month with double the shipments and double the spend. The blended rate
 * would survive that untouched — it is cost over shipments, and both doubled —
 * so the error would not show up as a wrong rate on any screen. It would show up
 * as a volume-weighted comparison quietly weighting that month twice as heavily
 * as it should, which is the kind of thing nobody finds for a quarter. So months
 * present in the file are cleared from the staging tab before the file is
 * staged, and the report names them.
 *
 * @param {string} csvText  the file's raw text
 * @param {string} mode     'APPEND' or 'REPLACE'
 */
function uploadActualsCsv(csvText, mode) {
  prewarmForWrite_([SHEET.ACTUALS_IMPORT, SHEET.ACTUALS, SHEET.ACTUALS_AMENDS,
                    SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CALENDAR]);
  const perms = requirePermissions_();
  requireImportActuals_(perms);

  const upMode = normKey(mode);
  if (upMode !== 'APPEND' && upMode !== 'REPLACE') {
    throw new Error('Unknown upload mode "' + safeStr(mode) + '". Expected APPEND or REPLACE.');
  }

  // ---- the file -----------------------------------------------------------
  const csv = parseCsv_(csvText);
  if (!csv.length) {
    throw new Error('That file is empty. It needs a header row and at least one row of data.');
  }
  if (csv.length < 2) {
    throw new Error('That file has a header row but no data. Columns found: ' +
                    csv[0].map(function (h) { return safeStr(h); }).join(' | '));
  }
  if (csv.length - 1 > ACTUALS_UPLOAD_MAX_ROWS) {
    throw new Error('That file has ' + (csv.length - 1).toLocaleString() + ' rows, and the ' +
      'limit is ' + ACTUALS_UPLOAD_MAX_ROWS.toLocaleString() + '. A monthly extract is ' +
      'normally a few thousand — check this is the right file, or split it by month.');
  }

  /* Header check before anything is written. extractStandardColumns_ throws with
     the columns it did find, which is the only useful thing to say to someone
     holding the wrong export. */
  const src = extractStandardColumns_(csv[0]);

  // ---- map the file's columns onto the staging tab's own columns -----------
  /* The tab's header row is fixed by TABLES.ACTUALS_IMPORT and the importer
     matches on it, so rows are written in the TAB's order, not the file's. A
     file with its columns rearranged, or carrying extra columns the model does
     not use, therefore stages correctly rather than landing a month in the count
     column.
     Both sides are located through extractColumns_ rather than by literal header
     text, so neither the tab's position order nor the pound sign in its cost
     header is written down twice — the one place that spelling exists is TABLES. */
  const t = TABLES.ACTUALS_IMPORT;
  const tab = extractStandardColumns_(t.headers);
  const map = new Array(t.headers.length).fill(-1);
  map[tab.cCountry] = src.cCountry;
  map[tab.cBrand]   = src.cBrand;
  map[tab.cTreat]   = src.cTreat;
  map[tab.cMonth]   = src.cMonth;
  map[tab.cCount]   = src.cCount;
  map[tab.cCost]    = src.cCost;
  /* Carrier and method are aggregated away by the import, so a file without them
     stages fine — the columns are just left blank and reported as absent. */
  if (tab.cCarrier >= 0) map[tab.cCarrier] = src.cCarrier;
  if (tab.cMethod  >= 0) map[tab.cMethod]  = src.cMethod;
  const numericCols = {};
  numericCols[tab.cCount] = true;
  numericCols[tab.cCost]  = true;

  const missingOptional = t.headers.filter(function (h, i) { return map[i] < 0; });

  // ---- the months this file covers ---------------------------------------
  const monthKeys = {}, unreadableMonths = [];
  for (let i = 1; i < csv.length; i++) {
    const raw = safeStr(csv[i][src.cMonth]);
    if (!raw) continue;
    const ms = parseExtractMonth_(raw);
    if (ms) monthKeys[dateKey(ms)] = fmtDate(ms);
    else if (unreadableMonths.length < 5 && unreadableMonths.indexOf(raw) < 0) {
      unreadableMonths.push(raw);
    }
  }
  if (!Object.keys(monthKeys).length) {
    throw new Error('No month in that file could be read. The "' +
      safeStr(csv[0][src.cMonth]) + '" column needs values like "January, 2026". ' +
      'Found instead: ' + (unreadableMonths.join(', ') || '(all blank)'));
  }

  // ---- build the staging rows --------------------------------------------
  const width = t.headers.length;
  const staged = [];
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    const row = new Array(width).fill('');
    for (let c = 0; c < width; c++) {
      if (map[c] < 0) continue;
      const raw = safeStr(r[map[c]]);
      /* Count and cost go in as numbers so the tab is usable by eye and so a
         thousands separator cannot survive to be re-parsed. A blank stays blank
         rather than becoming 0: the difference between "this cost nothing" and
         "no cost was exported" is exactly what the zero-cost report is for, and
         writing 0 here would destroy it before the report ever ran. */
      if (numericCols[c]) row[c] = (raw === '' ? '' : parseExtractNumber_(raw));
      else                row[c] = raw;
    }
    staged.push(row);
  }

  // ---- stage it -----------------------------------------------------------
  const result = withLock_(function () {
    const sh = getSheet_(t.sheet);
    invalidateSheetCache_(t.sheet);

    let removedRows = 0, replacedMonths = [], kept = [];

    if (upMode === 'REPLACE') {
      removedRows = clearActualsImport_();
    } else {
      /* Read what is staged, keep the months this file does not cover, and
         rewrite the tab as kept-then-new. One clear and one setValues, rather
         than deleting rows one at a time — and it cannot leave the tab
         half-cleared if something throws between the two, because the rows to
         keep are already in memory before anything is removed. */
      const existing = getAllData_(t.sheet);
      if (existing.length > 1) {
        const ex = extractColumns_(existing[0]);
        const exMonth = ex.find(['Dispatched Date: Month', 'Dispatched Month', 'Month'], null);
        const hit = {};
        for (let i = 1; i < existing.length; i++) {
          const row = existing[i];
          if (!row.some(function (c) { return safeStr(c) !== ''; })) continue;   // blank
          const ms = exMonth >= 0 ? parseExtractMonth_(row[exMonth]) : null;
          const k = ms ? dateKey(ms) : '';
          if (k && monthKeys[k]) { hit[k] = (hit[k] || 0) + 1; removedRows++; continue; }
          kept.push(row.slice(0, width));
        }
        replacedMonths = Object.keys(hit).sort().map(function (k) {
          return { month: monthKeys[k], rowsRemoved: hit[k] };
        });
      }
      clearDataRows_(t.sheet);
    }

    const all = kept.concat(staged);
    if (all.length) {
      sh.getRange(2, 1, all.length, width)
        .setValues(all.map(function (row) {
          const out = row.slice(0, width);
          while (out.length < width) out.push('');
          return out;
        }));
    }
    invalidateSheetCache_(t.sheet);

    return { removedRows: removedRows, replacedMonths: replacedMonths,
             stagedRows: staged.length, totalStaged: all.length };
  });

  logAudit_('IMPORT', SHEET.ACTUALS_IMPORT, '', 'mode', upMode,
            result.stagedRows + ' rows staged, ' + result.removedRows + ' removed',
            'uploadActualsCsv ' + upMode + ' — months ' +
            Object.keys(monthKeys).sort().map(function (k) { return monthKeys[k]; }).join(', '),
            true);

  // ---- and run the same import the editor runs ---------------------------
  Logger.log('=== ACTUALS CSV UPLOAD (' + upMode + ') ===');
  Logger.log('  staged  : ' + result.stagedRows + ' row(s) from the file');
  Logger.log('  removed : ' + result.removedRows + ' row(s) already staged for those months');
  Logger.log('  months  : ' + Object.keys(monthKeys).sort().map(function (k) {
    return monthKeys[k];
  }).join(', '));

  const report = importActualsFromStagingCore_(true);

  report.mode = upMode;
  report.fileRows = csv.length - 1;
  report.stagedRows = result.stagedRows;
  report.stagingRowsRemoved = result.removedRows;
  report.totalStaged = result.totalStaged;
  report.replacedMonths = result.replacedMonths;
  report.monthsInFile = Object.keys(monthKeys).sort().map(function (k) { return monthKeys[k]; });
  report.unreadableMonths = unreadableMonths;
  /* Which columns were matched loosely, so a report that looks wrong can be
     traced to the header that was guessed rather than to the data. */
  report.columnsMatched = src.cols.matched().map(function (m) {
    return { label: m[0], header: m[3], mode: m[2] };
  });
  report.columnsMissing = missingOptional;
  return report;
}


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