/**
 * Postage Forecast Portal — structure.gs
 *
 * Adding and changing the shape of the model: High Level IDs, Modelling IDs,
 * and the dimension lists behind them.
 *
 * This is what makes "add a new courier" or "add a new modelling area" a data
 * operation rather than a code change. Nothing here hard-codes a brand, a geo,
 * a carrier or a method — they all come from Dim_Reference, Dim_Carrier and
 * Dim_Method, and adding a row to one of those is enough.
 */

// ─────────────────────────────────────────────────────────────────────────────
// HIGH LEVEL IDS
// ─────────────────────────────────────────────────────────────────────────────

/** MEDEXPRESS_GB_WL_MOUNJARO — stable, readable, and independent of the number. */
function buildHighLevelCode_(brand, geo, treatmentType, wlSplit) {
  return [brand, geo, treatmentType, wlSplit]
    .map(v => safeStr(v).toUpperCase())
    .filter(v => v && v !== '*')
    .join('_');
}

function saveHighLevelId(p) {
  prewarmForWrite_([SHEET.HIGH_LEVEL_IDS, SHEET.HIGH_LEVEL_IDS_AMENDS, SHEET.DIM_REFERENCE]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const brand = safeStr(p.brand).toUpperCase();
  const geo   = safeStr(p.geo).toUpperCase();
  const tt    = safeStr(p.treatmentType).toUpperCase();
  const wl    = safeStr(p.wlSplit).toUpperCase() || '*';
  if (!brand || !geo || !tt) throw new Error('Brand, geo and treatment type are all required.');

  assertInReferenceList_('BRAND', brand);
  assertInReferenceList_('GEO', geo);
  assertInReferenceList_('TREATMENT_TYPE', tt);
  if (wl !== '*') assertInReferenceList_('WL_SPLIT', wl);

  const t = TABLES.HIGH_LEVEL_IDS, C = COL.HIGH_LEVEL_IDS, width = t.headers.length;
  const isNew = !safeInt(p.id);

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    // the same combination twice would make every SUMIFS-style lookup ambiguous
    const data = getAllData_(t.sheet);
    for (let i = 1; i < data.length; i++) {
      if (safeInt(data[i][C.High_Level_ID]) === safeInt(p.id)) continue;
      if (!safeBool(data[i][C.Active])) continue;
      if (normKey(data[i][C.Brand]) === brand && normKey(data[i][C.Geo]) === geo &&
          normKey(data[i][C.Treatment_Type]) === tt && normKey(data[i][C.WL_Split]) === wl) {
        throw new Error('High Level ID ' + safeInt(data[i][C.High_Level_ID]) +
                        ' is already ' + [brand, geo, tt, wl].join(' / ') + '.');
      }
    }

    const sh = getSheet_(t.sheet);
    let rowIndex, before = null, row;

    if (isNew) {
      row = blankRow_('HIGH_LEVEL_IDS');
      row[C.High_Level_ID] = getNextId_(t.sheet, C.High_Level_ID);
      row[C.Created_TS]    = new Date();
      row[C.Created_By]    = perms.email;
      rowIndex = sh.getLastRow() + 1;
    } else {
      rowIndex = findRowById_(t.sheet, C.High_Level_ID, p.id);
      if (rowIndex < 0) throw new Error('That High Level ID no longer exists.');
      before = readRow_(t.sheet, rowIndex, width);
      assertCanEditHighLevelId_(perms, safeInt(p.id));
      row = before.slice();
    }

    row[C.High_Level_Code] = buildHighLevelCode_(brand, geo, tt, wl);
    row[C.Brand]           = brand;
    row[C.Geo]             = geo;
    row[C.Treatment_Type]  = tt;
    row[C.WL_Split]        = wl;
    row[C.Currency]        = safeStr(p.currency).toUpperCase() || 'GBP';
    row[C.Active]          = (p.active === undefined) ? true : !!p.active;
    row[C.Sort_Order]      = safeInt(p.sortOrder) || (safeInt(row[C.High_Level_ID]) * 10);
    row[C.Notes]           = safeStr(p.notes);
    row[C.Updated_TS]      = new Date();
    row[C.Updated_By]      = perms.email;

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);
    recordChange_('HIGH_LEVEL_IDS', row[C.High_Level_ID], before, row, isNew ? 'CREATE' : 'UPDATE');
    return { ok: true, id: row[C.High_Level_ID], code: row[C.High_Level_Code], isNew: isNew };
  });
}


function deleteHighLevelId(id) {
  prewarmForWrite_([SHEET.HIGH_LEVEL_IDS, SHEET.HIGH_LEVEL_IDS_AMENDS, SHEET.MODELLING_IDS]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);
  assertCanEditHighLevelId_(perms, id);

  const t = TABLES.HIGH_LEVEL_IDS, C = COL.HIGH_LEVEL_IDS, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.High_Level_ID, id);
    if (rowIndex < 0) throw new Error('That High Level ID no longer exists.');

    // refuse rather than orphan its routes
    const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
    let live = 0;
    for (let i = 1; i < md.length; i++) {
      if (safeInt(md[i][M.High_Level_ID]) === safeInt(id) && safeBool(md[i][M.Active])) live++;
    }
    if (live) throw new Error('That High Level ID still has ' + live +
      ' active Modelling ID(s). Deactivate those first.');

    const before = readRow_(t.sheet, rowIndex, width);
    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('HIGH_LEVEL_IDS', safeInt(id), before, row, 'DELETE');
    return { ok: true };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// MODELLING IDS
// ─────────────────────────────────────────────────────────────────────────────

function saveModellingId(p) {
  prewarmForWrite_([SHEET.MODELLING_IDS, SHEET.MODELLING_IDS_AMENDS,
                    SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CARRIER, SHEET.DIM_METHOD]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const hlId    = safeInt(p.hlId);
  const carrier = safeStr(p.carrier).toUpperCase();
  const method  = safeStr(p.method).toUpperCase();
  const lp      = safeStr(p.letterParcel).toUpperCase();

  assertCanEditHighLevelId_(perms, hlId);
  if (lp !== 'LETTER' && lp !== 'PARCEL') throw new Error('Class must be LETTER or PARCEL.');

  const carriers = {}, cData = getAllData_(SHEET.DIM_CARRIER), CC = COL.DIM_CARRIER;
  for (let i = 1; i < cData.length; i++) {
    if (safeBool(cData[i][CC.Active])) carriers[safeStr(cData[i][CC.Carrier_Code])] = true;
  }
  if (!carriers[carrier]) throw new Error('Carrier "' + carrier +
    '" is not in Dim_Carrier. Add it there first.');

  const mData = getAllData_(SHEET.DIM_METHOD), MM = COL.DIM_METHOD;
  let methodOwner = null;
  for (let i = 1; i < mData.length; i++) {
    if (safeStr(mData[i][MM.Method_Code]) === method && safeBool(mData[i][MM.Active])) {
      methodOwner = safeStr(mData[i][MM.Carrier_Code]); break;
    }
  }
  if (methodOwner === null) throw new Error('Method "' + method +
    '" is not in Dim_Method. Add it there first.');

  const t = TABLES.MODELLING_IDS, C = COL.MODELLING_IDS, width = t.headers.length;
  const isNew = !safeInt(p.id);

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    // The segment the route sits under now, not only the one it is being moved
    // to — re-parenting a route is exactly the move this has to refuse.
    assertCanEditRowOwner_(perms, t.sheet, C.Modelling_ID, p.id,
                           C.High_Level_ID, 'HIGH_LEVEL_ID');

    const data = getAllData_(t.sheet);
    for (let i = 1; i < data.length; i++) {
      if (safeInt(data[i][C.Modelling_ID]) === safeInt(p.id)) continue;
      if (!safeBool(data[i][C.Active])) continue;
      if (safeInt(data[i][C.High_Level_ID]) === hlId &&
          normKey(data[i][C.Carrier_Code]) === carrier &&
          normKey(data[i][C.Method_Code]) === method &&
          normKey(data[i][C.Letter_Parcel]) === lp) {
        throw new Error('Modelling ID ' + safeInt(data[i][C.Modelling_ID]) +
                        ' is already that route.');
      }
    }

    const hlCode = highLevelCodeFor_(hlId);
    const sh = getSheet_(t.sheet);
    let rowIndex, before = null, row;

    if (isNew) {
      row = blankRow_('MODELLING_IDS');
      row[C.Modelling_ID] = getNextId_(t.sheet, C.Modelling_ID);
      row[C.Created_TS]   = new Date();
      row[C.Created_By]   = perms.email;
      rowIndex = sh.getLastRow() + 1;
    } else {
      rowIndex = findRowById_(t.sheet, C.Modelling_ID, p.id);
      if (rowIndex < 0) throw new Error('That Modelling ID no longer exists.');
      before = readRow_(t.sheet, rowIndex, width);
      row = before.slice();
    }

    row[C.High_Level_ID]  = hlId;
    row[C.Carrier_Code]   = carrier;
    row[C.Method_Code]    = method;
    row[C.Letter_Parcel]  = lp;
    row[C.Modelling_Code] = hlCode + '__' + carrier + '__' + method + '__' + lp.charAt(0);
    row[C.Active]         = (p.active === undefined) ? true : !!p.active;
    row[C.Notes]          = safeStr(p.notes);
    row[C.Updated_TS]     = new Date();
    row[C.Updated_By]     = perms.email;

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);
    recordChange_('MODELLING_IDS', row[C.Modelling_ID], before, row, isNew ? 'CREATE' : 'UPDATE');

    const warning = (methodOwner && methodOwner !== carrier)
      ? 'Saved, but note: ' + method + ' is listed under carrier ' + methodOwner +
        ' in Dim_Method. Validation will flag this as METHOD_CARRIER.'
      : '';
    return { ok: true, id: row[C.Modelling_ID], code: row[C.Modelling_Code],
             isNew: isNew, warning: warning };
  });
}


function deleteModellingId(id) {
  prewarmForWrite_([SHEET.MODELLING_IDS, SHEET.MODELLING_IDS_AMENDS, SHEET.MIX_METHOD]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);
  assertCanEditModellingId_(perms, id);

  const t = TABLES.MODELLING_IDS, C = COL.MODELLING_IDS, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.Modelling_ID, id);
    if (rowIndex < 0) throw new Error('That Modelling ID no longer exists.');

    // A route carrying volume cannot just disappear — the remaining mixes would
    // no longer total 100% and every month would come out low.
    const mix = getAllData_(SHEET.MIX_METHOD), X = COL.MIX_METHOD;
    let carrying = 0;
    for (let i = 1; i < mix.length; i++) {
      if (safeInt(mix[i][X.Modelling_ID]) !== safeInt(id)) continue;
      if (!safeBool(mix[i][X.Active])) continue;
      if (safeNum(mix[i][X.Mix_Pct]) > 0) carrying++;
    }
    if (carrying) throw new Error('That route still carries volume in ' + carrying +
      ' mix period(s). Set its mix to 0% and redistribute to the others first, ' +
      'or the totals will no longer add to 100%.');

    const before = readRow_(t.sheet, rowIndex, width);
    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('MODELLING_IDS', safeInt(id), before, row, 'DELETE');
    return { ok: true };
  });
}


/**
 * Copy one High Level ID's set of routes onto another.
 *
 * Adding a modelling area usually means "the same carriers as the one next to
 * it". Doing that by hand across 44 routes is where mistakes come from.
 * Rates and mixes are NOT copied — those are decisions, not boilerplate.
 */
function cloneModellingIds(fromHlId, toHlId) {
  prewarmForWrite_([SHEET.MODELLING_IDS, SHEET.MODELLING_IDS_AMENDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);
  assertCanEditHighLevelId_(perms, toHlId);
  if (!canSeeHighLevelId_(perms, fromHlId, false)) {
    throw new Error('High Level ID ' + fromHlId + ' is outside the area you can see.');
  }

  const t = TABLES.MODELLING_IDS, C = COL.MODELLING_IDS, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const data = getAllData_(t.sheet);

    const existing = {};
    const source = [];
    for (let i = 1; i < data.length; i++) {
      if (!safeBool(data[i][C.Active])) continue;
      const key = normKey(data[i][C.Carrier_Code]) + '|' + normKey(data[i][C.Method_Code]) +
                  '|' + normKey(data[i][C.Letter_Parcel]);
      if (safeInt(data[i][C.High_Level_ID]) === safeInt(toHlId))   existing[key] = true;
      if (safeInt(data[i][C.High_Level_ID]) === safeInt(fromHlId)) {
        source.push({ key: key, carrier: safeStr(data[i][C.Carrier_Code]),
                      method: safeStr(data[i][C.Method_Code]),
                      lp: safeStr(data[i][C.Letter_Parcel]) });
      }
    }
    if (!source.length) throw new Error('High Level ID ' + fromHlId + ' has no active routes.');

    const hlCode = highLevelCodeFor_(toHlId);
    const sh = getSheet_(t.sheet);
    let nextId = getNextId_(t.sheet, C.Modelling_ID) - 1;
    const rows = [];

    source.forEach(s => {
      if (existing[s.key]) return;
      const row = blankRow_('MODELLING_IDS');
      row[C.Modelling_ID]   = ++nextId;
      row[C.High_Level_ID]  = safeInt(toHlId);
      row[C.Carrier_Code]   = s.carrier;
      row[C.Method_Code]    = s.method;
      row[C.Letter_Parcel]  = s.lp;
      row[C.Modelling_Code] = hlCode + '__' + s.carrier + '__' + s.method + '__' + s.lp.charAt(0);
      row[C.Active]         = true;
      row[C.Notes]          = 'Cloned from High Level ID ' + fromHlId;
      row[C.Created_TS]     = new Date();
      row[C.Created_By]     = perms.email;
      row[C.Updated_TS]     = new Date();
      row[C.Updated_By]     = perms.email;
      rows.push(row);
    });

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows);
      invalidateSheetCache_(t.sheet);
      recordChangesBatch_('MODELLING_IDS', rows.map(r =>
        ({ id: r[C.Modelling_ID], before: null, after: r, type: 'CREATE' })));
      logAudit_('CREATE', 'CLONE_ROUTES', toHlId, '', '', rows.length + ' routes',
                'cloned from High Level ID ' + fromHlId, true);
    }

    return { ok: true, created: rows.length, skipped: source.length - rows.length,
             note: 'Routes only. Add rates and mixes before this High Level ID will price.' };
  });
}


/**
 * Copy rates, surcharges and mixes from one segment to another.
 *
 * cloneModellingIds deliberately copies routes only, on the grounds that a
 * copied rate nobody notices is worse than a missing one the checks shout about.
 * That holds when the new segment is genuinely different. It does not hold when
 * it is an equivalent one — copying 297 rows by hand is its own kind of risk.
 *
 * So this exists, but with three things the route clone does not need:
 *
 *   - it previews by default, and writes nothing until asked
 *   - it matches routes on carrier, method and class, and reports any that do
 *     not pair up rather than guessing
 *   - everything it creates is marked as copied, in Source_Ref and Notes, so a
 *     rate that was never actually negotiated is findable later
 *
 * @param {Object} p { fromHlId, toHlId, rates, surcharges, mixes, letterParcel,
 *                     replace, preview, scenarioId }
 */
function copySegmentData(p) {
  prewarmForWrite_([SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS,
                    SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS,
                    SHEET.RATE_SURCHARGE, SHEET.RATE_SURCHARGE_AMENDS,
                    SHEET.MIX_METHOD, SHEET.MIX_METHOD_AMENDS,
                    SHEET.MIX_LETTERPARCEL, SHEET.MIX_LETTERPARCEL_AMENDS]);
  const perms = requirePermissions_();
  requireEditRates_(perms);
  requireEditMixes_(perms);

  const fromHl = safeInt(p.fromHlId), toHl = safeInt(p.toHlId);
  if (!fromHl || !toHl) throw new Error('Pick a segment to copy from and one to copy into.');
  if (fromHl === toHl) throw new Error('Those are the same segment.');
  assertCanEditHighLevelId_(perms, toHl);
  if (!canSeeHighLevelId_(perms, fromHl, false)) {
    throw new Error('High Level ID ' + fromHl + ' is outside the area you can see.');
  }

  const scenarioId = safeInt(p.scenarioId) || 1;
  const want = { rates: p.rates !== false, surcharges: p.surcharges !== false,
                 mixes: p.mixes !== false, letterParcel: p.letterParcel !== false };

  // ---- pair the routes up -------------------------------------------------
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const srcByKey = {}, tgtByKey = {};
  const srcRoutes = [], tgtRoutes = [];
  for (let i = 1; i < md.length; i++) {
    if (!safeBool(md[i][M.Active])) continue;
    const hl = safeInt(md[i][M.High_Level_ID]);
    if (hl !== fromHl && hl !== toHl) continue;
    const key = normKey(md[i][M.Carrier_Code]) + '|' + normKey(md[i][M.Method_Code]) +
                '|' + normKey(md[i][M.Letter_Parcel]);
    const r = { id: safeInt(md[i][M.Modelling_ID]), key: key,
                carrier: safeStr(md[i][M.Carrier_Code]),
                method: safeStr(md[i][M.Method_Code]),
                lp: safeStr(md[i][M.Letter_Parcel]) };
    if (hl === fromHl) { srcByKey[key] = r; srcRoutes.push(r); }
    else               { tgtByKey[key] = r; tgtRoutes.push(r); }
  }

  const matched = [], missingInTarget = [], extraInTarget = [];
  srcRoutes.forEach(function (r) {
    if (tgtByKey[r.key]) matched.push({ from: r.id, to: tgtByKey[r.key].id, label: r.carrier + ' / ' + r.method + ' / ' + r.lp });
    else missingInTarget.push(r.carrier + ' / ' + r.method + ' / ' + r.lp);
  });
  tgtRoutes.forEach(function (r) {
    if (!srcByKey[r.key]) extraInTarget.push(r.carrier + ' / ' + r.method + ' / ' + r.lp);
  });

  if (!matched.length) {
    throw new Error('No routes pair up between those two segments. Copy the routes across ' +
      'first, using Clone from another under High Level & Modelling IDs.');
  }

  const fromId = {}, toId = {};
  matched.forEach(function (m) { fromId[m.from] = m.to; toId[m.to] = m.from; });

  // ---- work out what would be written ------------------------------------
  const plan = { rateBase: [], rateSurcharge: [], mixMethod: [], mixLP: [] };
  const skipped = { rateBase: 0, rateSurcharge: 0, mixMethod: 0, mixLP: 0 };

  function existingCount(sheetName, midCol, activeCol, scenarioCol) {
    const data = getAllData_(sheetName), have = {};
    for (let i = 1; i < data.length; i++) {
      if (!safeBool(data[i][activeCol])) continue;
      if (scenarioCol !== undefined && safeInt(data[i][scenarioCol]) !== scenarioId) continue;
      const mid = safeInt(data[i][midCol]);
      if (toId[mid]) have[mid] = (have[mid] || 0) + 1;
    }
    return have;
  }

  if (want.rates) {
    const C = COL.RATE_BASE;
    const have = existingCount(SHEET.RATE_BASE, C.Modelling_ID, C.Active, C.Scenario_ID);
    const data = getAllData_(SHEET.RATE_BASE);
    for (let i = 1; i < data.length; i++) {
      const mid = safeInt(data[i][C.Modelling_ID]);
      if (!fromId[mid] || !safeBool(data[i][C.Active])) continue;
      if (safeInt(data[i][C.Scenario_ID]) !== scenarioId) continue;
      if (have[fromId[mid]] && !p.replace) { skipped.rateBase++; continue; }
      plan.rateBase.push({ src: data[i].slice(), toMid: fromId[mid] });
    }
  }

  if (want.surcharges) {
    const C = COL.RATE_SURCHARGE;
    const have = existingCount(SHEET.RATE_SURCHARGE, C.Modelling_ID, C.Active, C.Scenario_ID);
    const data = getAllData_(SHEET.RATE_SURCHARGE);
    for (let i = 1; i < data.length; i++) {
      const mid = safeInt(data[i][C.Modelling_ID]);
      if (!fromId[mid] || !safeBool(data[i][C.Active])) continue;
      if (safeInt(data[i][C.Scenario_ID]) !== scenarioId) continue;
      if (have[fromId[mid]] && !p.replace) { skipped.rateSurcharge++; continue; }
      plan.rateSurcharge.push({ src: data[i].slice(), toMid: fromId[mid] });
    }
  }

  if (want.mixes) {
    const C = COL.MIX_METHOD;
    const have = existingCount(SHEET.MIX_METHOD, C.Modelling_ID, C.Active, C.Scenario_ID);
    const data = getAllData_(SHEET.MIX_METHOD);
    for (let i = 1; i < data.length; i++) {
      const mid = safeInt(data[i][C.Modelling_ID]);
      if (!fromId[mid] || !safeBool(data[i][C.Active])) continue;
      if (safeInt(data[i][C.Scenario_ID]) !== scenarioId) continue;
      if (have[fromId[mid]] && !p.replace) { skipped.mixMethod++; continue; }
      plan.mixMethod.push({ src: data[i].slice(), toMid: fromId[mid] });
    }
  }

  if (want.letterParcel) {
    const C = COL.MIX_LETTERPARCEL;
    const data = getAllData_(SHEET.MIX_LETTERPARCEL);
    let targetHas = 0;
    for (let i = 1; i < data.length; i++) {
      if (safeInt(data[i][C.High_Level_ID]) === toHl && safeBool(data[i][C.Active])) targetHas++;
    }
    for (let i = 1; i < data.length; i++) {
      if (safeInt(data[i][C.High_Level_ID]) !== fromHl || !safeBool(data[i][C.Active])) continue;
      if (safeInt(data[i][C.Scenario_ID]) !== scenarioId) continue;
      if (targetHas && !p.replace) { skipped.mixLP++; continue; }
      plan.mixLP.push({ src: data[i].slice() });
    }
  }

  // ---- would the copied mix still total 100%? ----------------------------
  // Only if every source route has a partner. A missing one takes its share with it.
  const mixWarning = (want.mixes && missingInTarget.length)
    ? ('The mix will not total 100% — ' + missingInTarget.length + ' route(s) in the source ' +
       'have no partner here, so their share is not carried across. Add those routes first, ' +
       'or redistribute the mix afterwards.')
    : '';

  const total = plan.rateBase.length + plan.rateSurcharge.length +
                plan.mixMethod.length + plan.mixLP.length;

  const summary = {
    matchedRoutes: matched.length,
    missingInTarget: missingInTarget,
    extraInTarget: extraInTarget,
    counts: { rateBase: plan.rateBase.length, rateSurcharge: plan.rateSurcharge.length,
              mixMethod: plan.mixMethod.length, mixLetterParcel: plan.mixLP.length },
    skipped: skipped,
    total: total,
    mixWarning: mixWarning
  };

  if (p.preview !== false) {
    summary.preview = true;
    return summary;
  }

  // ---- write --------------------------------------------------------------
  return withLock_(function () {
    const stamp = 'Copied from High Level ID ' + fromHl + ' on ' + fmtDate(new Date());
    let written = 0;

    function appendCopies(tableKey, rows, midCol, idCol, extras) {
      if (!rows.length) return 0;
      const t = TABLES[tableKey], C = COL[tableKey], sh = getSheet_(t.sheet);
      let nextId = getNextId_(t.sheet, idCol);
      const out = rows.map(function (r) {
        const row = r.src.slice();
        row[idCol] = nextId++;
        if (midCol !== null) row[midCol] = r.toMid;
        if (extras) extras(row, C);
        if (C.Source_Ref !== undefined) row[C.Source_Ref] = stamp;
        if (C.Notes !== undefined) {
          row[C.Notes] = (safeStr(row[C.Notes]) ? safeStr(row[C.Notes]) + '; ' : '') + stamp;
        }
        if (C.Created_TS !== undefined) { row[C.Created_TS] = new Date(); row[C.Created_By] = perms.email; }
        if (C.Updated_TS !== undefined) { row[C.Updated_TS] = new Date(); row[C.Updated_By] = perms.email; }
        row[C.Active] = true;
        return row;
      });
      sh.getRange(sh.getLastRow() + 1, 1, out.length, t.headers.length).setValues(out);
      invalidateSheetCache_(t.sheet);
      recordChangesBatch_(tableKey, out.map(function (row) {
        return { id: row[idCol], before: null, after: row, type: 'CREATE' }; }));
      return out.length;
    }

    written += appendCopies('RATE_BASE', plan.rateBase,
      COL.RATE_BASE.Modelling_ID, COL.RATE_BASE.Rate_ID);
    written += appendCopies('RATE_SURCHARGE', plan.rateSurcharge,
      COL.RATE_SURCHARGE.Modelling_ID, COL.RATE_SURCHARGE.Surcharge_Rate_ID);
    written += appendCopies('MIX_METHOD', plan.mixMethod,
      COL.MIX_METHOD.Modelling_ID, COL.MIX_METHOD.Mix_ID);
    written += appendCopies('MIX_LETTERPARCEL', plan.mixLP,
      null, COL.MIX_LETTERPARCEL.LP_Mix_ID,
      function (row, C) { row[C.High_Level_ID] = toHl; });

    logAudit_('CREATE', 'COPY_SEGMENT', toHl, '', '', written + ' rows',
              'copied from High Level ID ' + fromHl, true);

    summary.preview = false;
    summary.written = written;
    return summary;
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// DIMENSIONS
// ─────────────────────────────────────────────────────────────────────────────

function assertInReferenceList_(listName, code) {
  const data = getAllData_(SHEET.DIM_REFERENCE), C = COL.DIM_REFERENCE;
  for (let i = 1; i < data.length; i++) {
    if (safeStr(data[i][C.List_Name]).toUpperCase() !== listName) continue;
    if (normKey(data[i][C.Code]) !== normKey(code)) continue;
    if (!safeBool(data[i][C.Active])) break;
    return true;
  }
  throw new Error('"' + code + '" is not in the ' + listName.replace(/_/g, ' ').toLowerCase() +
                  ' list. Add it under Admin first.');
}

function highLevelCodeFor_(hlId) {
  const data = getAllData_(SHEET.HIGH_LEVEL_IDS), C = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][C.High_Level_ID]) !== safeInt(hlId)) continue;
    return safeStr(data[i][C.High_Level_Code]) ||
           buildHighLevelCode_(data[i][C.Brand], data[i][C.Geo],
                               data[i][C.Treatment_Type], data[i][C.WL_Split]);
  }
  throw new Error('High Level ID ' + hlId + ' does not exist.');
}

/**
 * Where a reference value is actually used.
 *
 * This decides what may be changed. A label is cosmetic and can always be
 * edited; a code is a foreign key, so renaming one that is in use would orphan
 * every row pointing at it. Renaming one nobody references yet is harmless —
 * which is the case that matters, because that is a typo made minutes ago.
 */
function referenceUsage_(listName, code) {
  const want = normKey(code);
  const list = safeStr(listName).toUpperCase();
  const where = [];

  function countIn(sheetName, colIndex, label) {
    let n = 0;
    const data = getAllData_(sheetName);
    for (let i = 1; i < data.length; i++) {
      if (normKey(data[i][colIndex]) === want) n++;
    }
    if (n) where.push({ label: label, count: n });
    return n;
  }

  const H = COL.HIGH_LEVEL_IDS;
  if (list === 'BRAND')          countIn(SHEET.HIGH_LEVEL_IDS, H.Brand, 'High Level IDs');
  if (list === 'GEO')            countIn(SHEET.HIGH_LEVEL_IDS, H.Geo, 'High Level IDs');
  if (list === 'TREATMENT_TYPE') countIn(SHEET.HIGH_LEVEL_IDS, H.Treatment_Type, 'High Level IDs');
  if (list === 'WL_SPLIT')       countIn(SHEET.HIGH_LEVEL_IDS, H.WL_Split, 'High Level IDs');

  if (list === 'CURRENCY') {
    countIn(SHEET.HIGH_LEVEL_IDS, H.Currency, 'High Level IDs');
    countIn(SHEET.DIM_CARRIER, COL.DIM_CARRIER.Default_Currency, 'carriers');
    countIn(SHEET.RATE_BASE, COL.RATE_BASE.Currency, 'base rates');
    countIn(SHEET.RATE_SURCHARGE, COL.RATE_SURCHARGE.Currency, 'surcharges');
    countIn(SHEET.ACTUALS, COL.ACTUALS.Currency, 'actuals');
    countIn(SHEET.FX_RATES, COL.FX_RATES.Currency, 'FX rates');
  }
  if (list === 'LETTER_PARCEL') countIn(SHEET.MODELLING_IDS, COL.MODELLING_IDS.Letter_Parcel, 'routes');
  if (list === 'TEMP_REGIME')   countIn(SHEET.MIX_METHOD, COL.MIX_METHOD.Temp_Regime, 'method mix rows');

  let total = 0;
  where.forEach(function (w) { total += w.count; });
  return { total: total, where: where };
}

/** Lists the calculation branches on. Their members cannot be removed. */
const LOCKED_REFERENCE_LISTS = { LETTER_PARCEL: true, TEMP_REGIME: true };


/** Usage counts for every reference value, so the interface can show them. */
function getReferenceUsage() {
  requirePermissions_();
  prewarmSheetCache_([SHEET.DIM_REFERENCE, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CARRIER,
                      SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.ACTUALS,
                      SHEET.FX_RATES, SHEET.MODELLING_IDS, SHEET.MIX_METHOD,
                      SHEET.PERMISSIONS, SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING,
                      SHEET.CONFIG]);
  const data = getAllData_(SHEET.DIM_REFERENCE), C = COL.DIM_REFERENCE;
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const list = safeStr(data[i][C.List_Name]).toUpperCase();
    const code = safeStr(data[i][C.Code]);
    if (!list || !code) continue;
    const u = referenceUsage_(list, code);
    out[list + '|' + code] = { id: safeInt(data[i][C.Ref_ID]), total: u.total,
                               where: u.where, locked: !!LOCKED_REFERENCE_LISTS[list],
                               active: safeBool(data[i][C.Active]) };
  }
  return out;
}


/** Add a value to one of the dropdown lists, or change one that already exists. */
function saveReferenceValue(p) {
  prewarmForWrite_([SHEET.DIM_REFERENCE, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CARRIER,
                    SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.ACTUALS,
                    SHEET.FX_RATES, SHEET.MODELLING_IDS, SHEET.MIX_METHOD]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const listName = safeStr(p.listName).toUpperCase();
  const code     = safeStr(p.code).toUpperCase().replace(/[^A-Z0-9*_]/g, '_');
  if (!listName || !code) throw new Error('List and code are both required.');

  const t = TABLES.DIM_REFERENCE, C = COL.DIM_REFERENCE, width = t.headers.length;
  const id = safeInt(p.id);

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const data = getAllData_(t.sheet);

    let rowIndex = -1, before = null, oldCode = '';
    if (id) {
      rowIndex = findRowById_(t.sheet, C.Ref_ID, id);
      if (rowIndex < 0) throw new Error('That value no longer exists.');
      before = readRow_(t.sheet, rowIndex, width);
      oldCode = safeStr(before[C.Code]);
    }

    // a code has to stay unique within its list
    for (let i = 1; i < data.length; i++) {
      if (safeInt(data[i][C.Ref_ID]) === id) continue;
      if (safeStr(data[i][C.List_Name]).toUpperCase() !== listName) continue;
      if (normKey(data[i][C.Code]) === normKey(code)) {
        throw new Error('"' + code + '" is already in that list.');
      }
    }

    // renaming is only safe while nothing refers to the old code
    if (id && normKey(oldCode) !== normKey(code)) {
      const u = referenceUsage_(listName, oldCode);
      if (u.total) {
        const parts = u.where.map(function (w) { return w.count + ' ' + w.label; });
        throw new Error('"' + oldCode + '" is used by ' + parts.join(', ') +
          ', so its code cannot be changed — every one of those rows points at it. ' +
          'Change the label instead, or add a new code and move the rows across.');
      }
    }

    const sh = getSheet_(t.sheet);
    let row, maxSort = 0;
    for (let i = 1; i < data.length; i++) {
      if (safeStr(data[i][C.List_Name]).toUpperCase() === listName) {
        maxSort = Math.max(maxSort, safeInt(data[i][C.Sort_Order]));
      }
    }

    if (rowIndex < 0) {
      row = blankRow_('DIM_REFERENCE');
      row[C.Ref_ID] = getNextId_(t.sheet, C.Ref_ID);
      rowIndex = sh.getLastRow() + 1;
    } else {
      row = before.slice();
    }

    row[C.List_Name]  = listName;
    row[C.Code]       = code;
    row[C.Label]      = safeStr(p.label) || code;
    row[C.Sort_Order] = safeInt(p.sortOrder) || safeInt(row[C.Sort_Order]) || (maxSort + 10);
    row[C.Active]     = (p.active === undefined) ? true : !!p.active;
    row[C.Notes]      = safeStr(p.notes);

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    logAudit_(before ? 'UPDATE' : 'CREATE', SHEET.DIM_REFERENCE, row[C.Ref_ID],
              'Code', before ? oldCode : '', code,
              listName + ': ' + safeStr(row[C.Label]), true);
    return { ok: true, id: row[C.Ref_ID], code: code, isNew: !before };
  });
}


/**
 * Remove a reference value.
 *
 * Refused if anything uses it, and refused outright for the two lists the
 * calculation branches on. Otherwise it is deactivated rather than deleted, so
 * the audit trail keeps making sense.
 */
function deleteReferenceValue(id) {
  prewarmForWrite_([SHEET.DIM_REFERENCE, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_CARRIER,
                    SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.ACTUALS,
                    SHEET.FX_RATES, SHEET.MODELLING_IDS, SHEET.MIX_METHOD]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const t = TABLES.DIM_REFERENCE, C = COL.DIM_REFERENCE, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.Ref_ID, id);
    if (rowIndex < 0) throw new Error('That value no longer exists.');

    const before = readRow_(t.sheet, rowIndex, width);
    const listName = safeStr(before[C.List_Name]).toUpperCase();
    const code = safeStr(before[C.Code]);

    if (LOCKED_REFERENCE_LISTS[listName]) {
      throw new Error(listName.replace(/_/g, ' ').toLowerCase() +
        ' values cannot be removed — the calculation branches on them by name.');
    }

    const u = referenceUsage_(listName, code);
    if (u.total) {
      const parts = u.where.map(function (w) { return w.count + ' ' + w.label; });
      throw new Error('"' + code + '" is used by ' + parts.join(', ') +
        '. Change those to something else first, then remove it.');
    }

    const row = before.slice();
    row[C.Active] = false;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    logAudit_('DELETE', SHEET.DIM_REFERENCE, id, '', code, '',
              listName + ' value removed', true);
    return { ok: true, code: code };
  });
}


function saveCarrier(p) {
  prewarmForWrite_([SHEET.DIM_CARRIER]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const code = safeStr(p.code).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!code) throw new Error('A carrier code is required.');

  const t = TABLES.DIM_CARRIER, C = COL.DIM_CARRIER, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const sh = getSheet_(t.sheet);
    let rowIndex = findRowByText_(t.sheet, C.Carrier_Code, code);
    const isNew = rowIndex < 0;
    const row = isNew ? blankRow_('DIM_CARRIER') : readRow_(t.sheet, rowIndex, width);
    const before = isNew ? null : row.slice();

    row[C.Carrier_Code]     = code;
    row[C.Carrier_Name]     = safeStr(p.name) || code;
    row[C.Default_Currency] = safeStr(p.currency).toUpperCase() || 'GBP';
    row[C.Active]           = (p.active === undefined) ? true : !!p.active;
    row[C.Notes]            = safeStr(p.notes);

    if (isNew) sh.appendRow(row);
    else sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    logAudit_(isNew ? 'CREATE' : 'UPDATE', SHEET.DIM_CARRIER, code, '',
              before ? safeStr(before[C.Carrier_Name]) : '', safeStr(row[C.Carrier_Name]), '', true);
    return { ok: true, code: code, isNew: isNew };
  });
}

function saveMethod(p) {
  prewarmForWrite_([SHEET.DIM_METHOD, SHEET.DIM_CARRIER]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const code    = safeStr(p.code).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const carrier = safeStr(p.carrier).toUpperCase();
  if (!code || !carrier) throw new Error('Method code and carrier are both required.');

  const cData = getAllData_(SHEET.DIM_CARRIER), CC = COL.DIM_CARRIER;
  let found = false;
  for (let i = 1; i < cData.length; i++) {
    if (safeStr(cData[i][CC.Carrier_Code]) === carrier) { found = true; break; }
  }
  if (!found) throw new Error('Carrier "' + carrier + '" does not exist. Add it first.');

  const t = TABLES.DIM_METHOD, C = COL.DIM_METHOD, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const sh = getSheet_(t.sheet);
    let rowIndex = findRowByText_(t.sheet, C.Method_Code, code);
    const isNew = rowIndex < 0;
    const row = isNew ? blankRow_('DIM_METHOD') : readRow_(t.sheet, rowIndex, width);

    row[C.Method_Code]   = code;
    row[C.Carrier_Code]  = carrier;
    row[C.Method_Name]   = safeStr(p.name) || code;
    row[C.Service_Level] = safeStr(p.serviceLevel);
    row[C.Is_Tracked]    = !!p.isTracked;
    row[C.Active]        = (p.active === undefined) ? true : !!p.active;
    row[C.Notes]         = safeStr(p.notes);

    if (isNew) sh.appendRow(row);
    else sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    logAudit_(isNew ? 'CREATE' : 'UPDATE', SHEET.DIM_METHOD, code, '', '',
              carrier + '/' + code, '', true);
    return { ok: true, code: code, isNew: isNew };
  });
}

/** A new surcharge type. The engine picks it up with no code change. */
function saveSurchargeType(p) {
  prewarmForWrite_([SHEET.DIM_SURCHARGE]);
  const perms = requirePermissions_();
  requireEditStructure_(perms);

  const code = safeStr(p.code).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!code) throw new Error('A surcharge code is required.');
  const valueType = safeStr(p.valueType).toUpperCase();
  if (valueType !== 'PCT' && valueType !== 'AMT') {
    throw new Error('Value type must be PCT (a percentage of base) or AMT (a fixed amount).');
  }
  const proration = safeStr(p.proration).toUpperCase() || 'DAY_WEIGHTED';
  if (proration !== 'DAY_WEIGHTED' && proration !== 'POINT_IN_TIME') {
    throw new Error('Proration must be DAY_WEIGHTED or POINT_IN_TIME.');
  }

  const t = TABLES.DIM_SURCHARGE, C = COL.DIM_SURCHARGE, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const sh = getSheet_(t.sheet);
    let rowIndex = findRowByText_(t.sheet, C.Surcharge_Code, code);
    const isNew = rowIndex < 0;
    const row = isNew ? blankRow_('DIM_SURCHARGE') : readRow_(t.sheet, rowIndex, width);

    row[C.Surcharge_Code] = code;
    row[C.Surcharge_Name] = safeStr(p.name) || code;
    row[C.Value_Type]     = valueType;
    row[C.Applies_To]     = safeStr(p.appliesTo).toUpperCase() || 'BASE';
    row[C.Apply_Order]    = safeInt(p.applyOrder) || 1;
    row[C.Proration]      = proration;
    row[C.Active]         = (p.active === undefined) ? true : !!p.active;
    row[C.Notes]          = safeStr(p.notes);

    if (isNew) sh.appendRow(row);
    else sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    logAudit_(isNew ? 'CREATE' : 'UPDATE', SHEET.DIM_SURCHARGE, code, '', '',
              code + ' ' + valueType + ' ' + proration, '', true);
    return { ok: true, code: code, isNew: isNew };
  });
}

function findRowByText_(sheetName, colIndex, value) {
  const data = getAllData_(sheetName);
  const want = normKey(value);
  for (let i = 1; i < data.length; i++) {
    if (normKey(data[i][colIndex]) === want) return i + 1;
  }
  return -1;
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC
// ─────────────────────────────────────────────────────────────────────────────

function testStructureWrite() {
  requireMaintenance_();
  Logger.log('=== STRUCTURE TEST ===');
  const perms = requirePermissions_();
  Logger.log('  acting as ' + perms.email + ' (' + perms.role + ')');

  Logger.log('');
  Logger.log('--- a duplicate High Level ID must be refused ---');
  let refusedDup = false;
  try {
    saveHighLevelId({ brand: 'MEDEXPRESS', geo: 'GB', treatmentType: 'WL',
                      wlSplit: 'MOUNJARO', currency: 'GBP' });
  } catch (e) { refusedDup = true; Logger.log('  refused, correctly: ' + e.message); }
  if (!refusedDup) Logger.log('  PROBLEM — a duplicate was accepted');

  Logger.log('');
  Logger.log('--- a brand that is not in the reference list must be refused ---');
  let refusedRef = false;
  try {
    saveHighLevelId({ brand: 'NOTAREALBRAND', geo: 'GB', treatmentType: 'WL', wlSplit: '*' });
  } catch (e) { refusedRef = true; Logger.log('  refused, correctly: ' + e.message); }
  if (!refusedRef) Logger.log('  PROBLEM — an unknown brand was accepted');

  Logger.log('');
  Logger.log('--- deleting a route that carries volume must be refused ---');
  let refusedDel = false;
  try { deleteModellingId(16); }
  catch (e) { refusedDel = true; Logger.log('  refused, correctly: ' + e.message.slice(0, 130)); }
  if (!refusedDel) Logger.log('  PROBLEM — a route carrying volume was deactivated');

  Logger.log('');
  Logger.log('--- an unknown carrier must be refused ---');
  let refusedCarrier = false;
  try {
    saveModellingId({ hlId: 1, carrier: 'NOTACARRIER', method: 'RM24', letterParcel: 'PARCEL' });
  } catch (e) { refusedCarrier = true; Logger.log('  refused, correctly: ' + e.message); }
  if (!refusedCarrier) Logger.log('  PROBLEM — an unknown carrier was accepted');

  Logger.log('');
  Logger.log('--- codes are generated consistently ---');
  Logger.log('  MEDEXPRESS / GB / WL / MOUNJARO -> ' +
             buildHighLevelCode_('MEDEXPRESS', 'GB', 'WL', 'MOUNJARO'));
  Logger.log('  DERMATICA / GB / CORE_RX / *    -> ' +
             buildHighLevelCode_('DERMATICA', 'GB', 'CORE_RX', '*'));

  const ok = refusedDup && refusedRef && refusedDel && refusedCarrier;
  Logger.log('');
  Logger.log(ok ? 'STRUCTURE GUARDS WORKING'
                : 'PROBLEM — one of the guards did not fire. Send me this log.');
  Logger.log('');
  Logger.log('Nothing was created or changed — every attempt above was meant to fail.');
  return { ok: ok };
}


// ─────────────────────────────────────────────────────────────────────────────
// FINDING DUPLICATED ROUTES — read only
//
// A duplicated Modelling ID is not a display problem, and that is the whole
// reason this reports rather than deletes. engine.js line 198:
//
//     hlTotal[d.hlId] += contribution;
//
// Every Modelling ID under a High Level ID has its contribution ADDED. Two rows
// describing the same carrier, method and class therefore contribute twice, so the
// segment's forecast rate is overstated by however many copies exist. Nothing
// flags it: each row is individually valid, the mix still sums to 100%, and
// validation has no rule against two routes being identical.
//
// Which means deduplicating a dropdown would hide an inflated forecast instead of
// fixing one. The rows have to go, not the symptom.
//
// It also has to be the RIGHT row that goes. A duplicate pair is rarely
// symmetrical: one usually carries the rates and mixes and the other is empty, and
// deactivating the populated one silently prices that route at nothing. So every
// copy is reported with what hangs off it — base rates, surcharges, method mixes —
// and the recommendation follows the data rather than the row order.
//
// Dim_Method is checked too, because a duplicate there would put a method in the
// dropdown twice for EVERY segment. If the complaint is about two High Level IDs
// specifically, that is not the cause — but ruling it out costs one pass.
// ─────────────────────────────────────────────────────────────────────────────

/** The identity of a route: same High Level ID, carrier, method and class. */
function dupKeyOf_(hlId, carrier, method, lp) {
  return safeInt(hlId) + '|' + normKey(carrier) + '|' + normKey(method) + '|' + normKey(lp);
}

/** How many rows in each rate/mix table point at this Modelling ID. */
function dupAttachedCounts_(modellingId) {
  const mid = safeInt(modellingId);
  function count(sheetName, colIdx, activeIdx) {
    const d = getAllData_(sheetName);
    let all = 0, active = 0;
    for (let i = 1; i < d.length; i++) {
      if (safeInt(d[i][colIdx]) !== mid) continue;
      all++;
      if (activeIdx === undefined || safeBool(d[i][activeIdx])) active++;
    }
    return { all: all, active: active };
  }
  return {
    base:      count(SHEET.RATE_BASE,      COL.RATE_BASE.Modelling_ID,      COL.RATE_BASE.Active),
    surcharge: count(SHEET.RATE_SURCHARGE, COL.RATE_SURCHARGE.Modelling_ID, COL.RATE_SURCHARGE.Active),
    mix:       count(SHEET.MIX_METHOD,     COL.MIX_METHOD.Modelling_ID,     COL.MIX_METHOD.Active)
  };
}

/**
 * Report duplicated routes, and what is attached to each copy.
 *
 * @param {number=} hlIdFilter  optional: look at one High Level ID only
 */
function diagnoseDuplicateRoutes(hlIdFilter) {
  requireMaintenance_();
  Logger.log('=== DUPLICATE ROUTES (read-only, nothing is written) ===');

  prewarmSheetCache_([SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_METHOD,
                      SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.MIX_METHOD,
                      SHEET.PERMISSIONS, SHEET.PORTAL_ROLES, SHEET.CONFIG]);

  Logger.log('  spreadsheet : ' + spreadsheetId_());
  const only = safeInt(hlIdFilter);
  if (only) Logger.log('  restricted to High Level ID ' + only);

  // ---- 1. Dim_Method, which feeds the method dropdown for every segment ---
  Logger.log('');
  Logger.log('════ 1. Dim_Method — the dropdown source, shared by all segments ════');
  const dm = getAllData_(SHEET.DIM_METHOD), DM = COL.DIM_METHOD;
  const byCode = {};
  for (let i = 1; i < dm.length; i++) {
    const code = normKey(dm[i][DM.Method_Code]);
    if (!code) continue;
    (byCode[code] = byCode[code] || []).push({
      row: i + 1, carrier: safeStr(dm[i][DM.Carrier_Code]),
      name: safeStr(dm[i][DM.Method_Name]), active: safeBool(dm[i][DM.Active])
    });
  }
  const dupMethods = Object.keys(byCode).filter(function (c) { return byCode[c].length > 1; });
  if (!dupMethods.length) {
    Logger.log('  No duplicate Method_Code. So the dropdown is not doubled globally,');
    Logger.log('  and a per-segment complaint is about Modelling_IDs below, not this.');
  } else {
    Logger.log('  ' + dupMethods.length + ' Method_Code(s) appear more than once:');
    dupMethods.sort().forEach(function (c) {
      Logger.log('    ' + c + ':');
      byCode[c].forEach(function (m) {
        Logger.log('      sheet row ' + pad_(m.row, 5) + '  ' + pad_(m.carrier, 12) + '  ' +
                   m.name + (m.active ? '  ACTIVE' : '  inactive'));
      });
    });
    Logger.log('  Only ACTIVE rows reach the dropdown (loadMethods_ in Code.gs).');
  }

  // ---- 2. Modelling_IDs — the ones that double the forecast --------------
  Logger.log('');
  Logger.log('════ 2. Modelling_IDs — duplicates here DOUBLE the forecast ════');

  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  const hlName = {};
  for (let i = 1; i < hl.length; i++) {
    const id = safeInt(hl[i][H.High_Level_ID]);
    if (id) hlName[id] = safeStr(hl[i][H.Brand]) + ' ' + safeStr(hl[i][H.Geo]) + ' ' +
                         safeStr(hl[i][H.Treatment_Type]) +
                         (safeStr(hl[i][H.WL_Split]) ? ' ' + safeStr(hl[i][H.WL_Split]) : '');
  }

  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const groups = {};
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id) continue;
    const hlId = safeInt(md[i][M.High_Level_ID]);
    if (only && hlId !== only) continue;
    const k = dupKeyOf_(hlId, md[i][M.Carrier_Code], md[i][M.Method_Code], md[i][M.Letter_Parcel]);
    (groups[k] = groups[k] || []).push({
      id: id, sheetRow: i + 1, hlId: hlId,
      carrier: normKey(md[i][M.Carrier_Code]),
      method: normKey(md[i][M.Method_Code]),
      lp: normKey(md[i][M.Letter_Parcel]),
      code: safeStr(md[i][M.Modelling_Code]),
      active: safeBool(md[i][M.Active])
    });
  }

  const dupKeys = Object.keys(groups).filter(function (k) { return groups[k].length > 1; });
  dupKeys.sort(function (a, b) { return safeInt(a) - safeInt(b) || (a < b ? -1 : 1); });

  const perHl = {}, recommend = [];
  let dupActivePairs = 0;

  if (!dupKeys.length) {
    Logger.log('  No duplicates. Every carrier/method/class appears at most once per');
    Logger.log('  High Level ID, so nothing here is double counting.');
  } else {
    Logger.log('  ' + dupKeys.length + ' route identity(ies) appear more than once.');
    let lastHl = null;
    dupKeys.forEach(function (k) {
      const g = groups[k];
      const actives = g.filter(function (r) { return r.active; });
      if (g[0].hlId !== lastHl) {
        Logger.log('');
        Logger.log('  ── High Level ID ' + g[0].hlId + '  ' + (hlName[g[0].hlId] || '') + ' ──');
        lastHl = g[0].hlId;
      }
      Logger.log('    ' + g[0].carrier + ' / ' + g[0].method + ' / ' + g[0].lp +
                 '   x' + g.length + (actives.length > 1
                   ? '   ** ' + actives.length + ' ACTIVE — counted ' + actives.length +
                     ' times in the forecast **'
                   : '   (' + actives.length + ' active)'));

      if (actives.length > 1) {
        dupActivePairs++;
        perHl[g[0].hlId] = (perHl[g[0].hlId] || 0) + 1;
      }

      /* What hangs off each copy decides which one may safely go. */
      const scored = g.map(function (r) {
        const a = dupAttachedCounts_(r.id);
        return { r: r, a: a, weight: a.base.active + a.surcharge.active + a.mix.active };
      });
      scored.forEach(function (s) {
        Logger.log('        MID ' + pad_(s.r.id, 4) + '  row ' + pad_(s.r.sheetRow, 5) + '  ' +
                   (s.r.active ? 'ACTIVE  ' : 'inactive') +
                   '  base ' + pad_(s.a.base.active, 3) + '/' + pad_(s.a.base.all, 3) +
                   '  surcharge ' + pad_(s.a.surcharge.active, 4) + '/' + pad_(s.a.surcharge.all, 4) +
                   '  mix ' + pad_(s.a.mix.active, 3) + '/' + pad_(s.a.mix.all, 3) +
                   '   ' + s.r.code);
      });

      if (actives.length > 1) {
        /* Keep the copy carrying the most live data; it is the one the rest of the
           sheet already refers to. An exact tie is not resolved here — picking
           between two equally-referenced routes is a judgement about which the
           business means, and guessing it silently is how the wrong one goes. */
        const live = scored.filter(function (s) { return s.r.active; })
                           .sort(function (x, y) { return y.weight - x.weight || x.r.id - y.r.id; });
        const tie = live.length > 1 && live[0].weight === live[1].weight;
        if (tie && live[0].weight === 0) {
          Logger.log('        -> all copies are empty. Any one may be kept; keeping the');
          Logger.log('           lowest ID (' + live[0].r.id + ') is the conventional choice.');
          recommend.push({ key: k, hlId: g[0].hlId, keep: live[0].r.id,
                           deactivate: live.slice(1).map(function (s) { return s.r.id; }),
                           confidence: 'all empty' });
        } else if (tie) {
          Logger.log('        -> TIE: two active copies carry the same amount of live data.');
          Logger.log('           Not recommending one. Decide which the business means.');
          recommend.push({ key: k, hlId: g[0].hlId, keep: null,
                           deactivate: [], confidence: 'tie — needs a human' });
        } else {
          Logger.log('        -> keep MID ' + live[0].r.id + ' (most live data), deactivate ' +
                     live.slice(1).map(function (s) { return s.r.id; }).join(', '));
          recommend.push({ key: k, hlId: g[0].hlId, keep: live[0].r.id,
                           deactivate: live.slice(1).map(function (s) { return s.r.id; }),
                           confidence: 'clear' });
        }
      }
    });
  }

  // ---- 3. what it does to the forecast -----------------------------------
  Logger.log('');
  Logger.log('════ 3. forecast impact ════');
  const hlIds = Object.keys(perHl).map(function (k) { return safeInt(k); })
                      .sort(function (a, b) { return a - b; });
  if (!hlIds.length) {
    Logger.log('  None. No route identity has two or more ACTIVE copies, so no High Level');
    Logger.log('  ID is double counting. Duplicates that are all inactive do not reach the');
    Logger.log('  engine: computeModel_ filters them out (engine.gs, "active !== false").');
  } else {
    hlIds.forEach(function (id) {
      Logger.log('  HL ' + pad_(id, 4) + '  ' + pad_(perHl[id], 4) + ' duplicated route(s)   ' +
                 (hlName[id] || ''));
    });
    Logger.log('');
    Logger.log('  Those segments\' published rates are OVERSTATED. Each duplicated route');
    Logger.log('  contributes once per active copy, and nothing in validation objects.');
    Logger.log('  Compare OUTPUT against previewOutput() after fixing, not before — the');
    Logger.log('  published figures still carry the duplicates.');
  }

  Logger.log('');
  Logger.log('--- what to do ---');
  if (dupActivePairs) {
    Logger.log('  Do NOT dedupe the dropdown. The dropdown is showing what is really there,');
    Logger.log('  and hiding it would leave the forecast overstated with nothing on screen');
    Logger.log('  to say so.');
    Logger.log('  The fix is deleteModellingId(<id>) on the copy to drop — a soft delete, so');
    Logger.log('  the row and its history survive and it stops reaching the engine.');
    Logger.log('  Check the "keep / deactivate" lines above first: deactivating the copy that');
    Logger.log('  holds the rates prices that route at nothing instead.');
  } else {
    Logger.log('  Nothing to remove. If the dropdown still shows something twice, it is not');
    Logger.log('  coming from a duplicated route or a duplicated method — say what the');
    Logger.log('  dropdown is and what it repeats.');
  }
  Logger.log('  Nothing was written. This function has no write path.');

  return { ok: true, spreadsheetId: spreadsheetId_(),
           duplicateMethodCodes: dupMethods,
           duplicateRouteIdentities: dupKeys.length,
           duplicatedActiveRoutes: dupActivePairs,
           affectedHighLevelIds: hlIds,
           recommendations: recommend };
}


/**
 * Dump every Modelling_IDs row for one or more High Level IDs, verbatim.
 *
 * Deliberately dumb. diagnoseDuplicateRoutes groups by carrier/method/class and
 * reported no duplicates for High Level IDs 11 and 12, while the Rates dropdown
 * shows every Modelling ID twice — so the grouping is asking the wrong question and
 * the honest next step is to stop inferring and print the rows.
 *
 * It also answers two things that grouping cannot:
 *
 *   - whether the same Modelling_ID VALUE appears on more than one row, which no
 *     earlier check looked for and which would put the same ID in the dropdown
 *     twice however different the rest of the row was
 *   - whether rows are INACTIVE, because loadModellingIdsForClient_ does not filter
 *     on Active and neither does the dropdown that reads it, so a soft-deleted
 *     route is still offered for rate editing
 *
 * @param {string=} hlIds  comma-separated, e.g. '11,12'. Default '11,12'.
 */
function dumpModellingIds(hlIds) {
  requireMaintenance_();
  Logger.log('=== Modelling_IDs, raw (read-only) ===');
  prewarmSheetCache_([SHEET.MODELLING_IDS, SHEET.PERMISSIONS, SHEET.PORTAL_ROLES,
                      SHEET.CONFIG]);
  Logger.log('  spreadsheet : ' + spreadsheetId_());

  const want = {};
  safeStr(hlIds || '11,12').split(',').forEach(function (s) {
    const n = safeInt(s); if (n) want[n] = true; });
  Logger.log('  High Level IDs: ' + Object.keys(want).join(', '));

  const data = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;

  /* Duplicate ID VALUES across the whole sheet, not just the requested segments —
     an ID reused elsewhere is worth knowing about wherever it is. */
  const idCount = {};
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][M.Modelling_ID]);
    if (id) idCount[id] = (idCount[id] || 0) + 1;
  }
  const reused = Object.keys(idCount).filter(function (k) { return idCount[k] > 1; });
  Logger.log('');
  Logger.log('--- Modelling_ID values appearing on more than one row, anywhere: ' +
             reused.length + ' ---');
  reused.sort(function (a, b) { return safeInt(a) - safeInt(b); }).forEach(function (k) {
    Logger.log('    ID ' + pad_(k, 6) + '  on ' + idCount[k] + ' rows'); });
  if (!reused.length) Logger.log('    None — every Modelling_ID is unique in the sheet.');

  Logger.log('');
  Logger.log('--- rows, in sheet order ---');
  Logger.log('  ' + pad_('row', 5) + pad_('ID', 7) + pad_('HL', 5) + '  ' +
             pad_('Carrier', 12) + pad_('Method', 32) + pad_('Class', 8) +
             pad_('Active', 8) + '  Modelling_Code');
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const hlId = safeInt(data[i][M.High_Level_ID]);
    if (!want[hlId]) continue;
    n++;
    Logger.log('  ' + pad_(i + 1, 5) + pad_(safeInt(data[i][M.Modelling_ID]), 7) +
               pad_(hlId, 5) + '  ' +
               pad_(safeStr(data[i][M.Carrier_Code]), 12) +
               pad_(safeStr(data[i][M.Method_Code]), 32) +
               pad_(safeStr(data[i][M.Letter_Parcel]), 8) +
               pad_(safeBool(data[i][M.Active]) ? 'yes' : 'NO', 8) + '  ' +
               safeStr(data[i][M.Modelling_Code]));
  }
  Logger.log('');
  Logger.log('  ' + n + ' row(s) for those High Level IDs.');
  Logger.log('  The Rates dropdown shows every row above whose High Level ID matches,');
  Logger.log('  INCLUDING ones marked Active = NO — it filters on hlId only. So a whole');
  Logger.log('  deactivated set would appear as a duplicate of the live one.');
  Logger.log('  Nothing was written.');
  return { ok: true, rows: n, reusedIds: reused.map(function (k) { return safeInt(k); }) };
}
