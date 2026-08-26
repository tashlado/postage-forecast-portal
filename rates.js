/**
 * Postage Forecast Portal — rates.gs
 *
 * Everything to do with editing rates. This is the portal's main job:
 * change a rate for a period, see the forecast move.
 *
 * Every write follows the same shape:
 *   1. prewarm the sheets this write will touch, in one call
 *   2. re-derive permissions from the spreadsheet — never trust the browser
 *   3. check the caller may touch THIS row, not just rates in general
 *   4. validate the values
 *   5. refuse overlapping periods
 *   6. take a lock, write, snapshot to history, log
 */

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/** Row index (1-based, as the sheet sees it) for a record, or -1. */
function findRowById_(sheetName, idColIndex, id) {
  const data = getAllData_(sheetName);
  const want = safeInt(id);
  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][idColIndex]) === want) return i + 1;
  }
  return -1;
}

/**
 * Reject a period that overlaps another for the same key.
 *
 * The rule lives in assertNoOverlapIn_, which takes the table contents as an
 * argument rather than reading them, because a bulk update checks against an
 * in-memory copy it is still mutating: it closes periods for many routes before
 * any of them reaches the sheet, and re-reading mid-plan would not see those
 * closes. One implementation, two callers, so the rule cannot drift between a
 * single save and a bulk one.
 */
function assertNoOverlap_(sheetName, C, keyMatch, fromKey, toKey, excludeId, idCol, label) {
  assertNoOverlapIn_(getAllData_(sheetName), C, keyMatch, fromKey, toKey,
                     excludeId, idCol, label);
}

function assertNoOverlapIn_(data, C, keyMatch, fromKey, toKey, excludeId, idCol, label) {
  for (let i = 1; i < data.length; i++) {
    if (!safeBool(data[i][C.Active])) continue;
    if (safeInt(data[i][idCol]) === safeInt(excludeId)) continue;
    if (!keyMatch(data[i])) continue;
    const f = dateKey(data[i][C.Valid_From]), t = dateKey(data[i][C.Valid_To]);
    if (fromKey <= t && toKey >= f) {
      throw new Error('That period overlaps an existing one for ' + label + ' (' +
        fmtDate(data[i][C.Valid_From]) + ' to ' + fmtDate(data[i][C.Valid_To]) +
        '). Close or shorten that period first.');
    }
  }
}

function validateDates_(fromStr, toStr) {
  const from = normaliseDate(fromStr), to = normaliseDate(toStr);
  if (!from) throw new Error('The start date is missing or not a date.');
  if (!to)   throw new Error('The end date is missing or not a date.');
  if (dateKey(from) > dateKey(to)) throw new Error('The start date is after the end date.');
  return { from: from, to: to, fromKey: dateKey(from), toKey: dateKey(to) };
}

function readRow_(sheetName, rowIndex, width) {
  return getSheet_(sheetName).getRange(rowIndex, 1, 1, width).getValues()[0];
}


// ─────────────────────────────────────────────────────────────────────────────
// BASE RATES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create or update one base rate.
 *
 * @param {Object} p { id, modellingId, validFrom, validTo, value,
 *                     currency, sourceRef, notes, scenarioId, closePrevious }
 */
function saveBaseRate(p) {
  prewarmForWrite_([SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  return withLock_(function () { return writeBaseRate_(p, perms); });
}

/**
 * One base rate written, assuming the caller already holds the script lock.
 *
 * Split out of saveBaseRate so bulkRateChange can write inside its own lock
 * instead of calling saveBaseRate and having saveBaseRate's finally-block hand
 * the lock back after the first route — which left every route after the first
 * running unprotected. Anything that needs a lock of its own calls
 * saveBaseRate; anything already holding one calls this.
 *
 * The order here matters and is the order the periods logic was written for:
 * authorise, then close the preceding period, then refuse an overlap, then
 * write. Do not reorder those three.
 */
function writeBaseRate_(p, perms) {
  const modellingId = safeInt(p.modellingId);
  assertCanEditModellingId_(perms, modellingId);

  const d = validateDates_(p.validFrom, p.validTo);
  const value = safeNum(p.value);
  if (value < 0) throw new Error('A base rate cannot be negative.');

  const t = TABLES.RATE_BASE, C = COL.RATE_BASE, width = t.headers.length;
  const isNew = !safeInt(p.id);

  invalidateSheetCache_(t.sheet);

  // The row's current owner, not just the one being asked for — before any
  // period is closed, so a refusal leaves nothing half-changed behind it.
  assertCanEditRowOwner_(perms, t.sheet, C.Rate_ID, p.id,
                         C.Modelling_ID, 'MODELLING_ID');

  if (p.closePrevious) closePrecedingPeriod_(t.sheet, C, C.Rate_ID,
    r => safeInt(r[C.Modelling_ID]) === modellingId, d.from, perms, 'RATE_BASE');

  invalidateSheetCache_(t.sheet);
  assertNoOverlap_(t.sheet, C, r => safeInt(r[C.Modelling_ID]) === modellingId,
                   d.fromKey, d.toKey, p.id, C.Rate_ID, 'Modelling ID ' + modellingId);

  const sh = getSheet_(t.sheet);
  let rowIndex, before = null, row;

  if (isNew) {
    row = blankRow_('RATE_BASE');
    row[C.Rate_ID]     = getNextId_(t.sheet, C.Rate_ID);
    row[C.Created_TS]  = new Date();
    row[C.Created_By]  = perms.email;
    rowIndex = sh.getLastRow() + 1;
  } else {
    rowIndex = findRowById_(t.sheet, C.Rate_ID, p.id);
    if (rowIndex < 0) throw new Error('That rate no longer exists — someone may have deleted it.');
    before = readRow_(t.sheet, rowIndex, width);
    row = before.slice();
  }

  row[C.Modelling_ID] = modellingId;
  row[C.Valid_From]   = d.from;
  row[C.Valid_To]     = d.to;
  row[C.Base_Rate]    = value;
  row[C.Currency]     = safeStr(p.currency) || currencyForModellingId_(modellingId);
  row[C.Scenario_ID]  = safeInt(p.scenarioId) || 1;
  row[C.Source_Ref]   = safeStr(p.sourceRef);
  row[C.Notes]        = safeStr(p.notes);
  row[C.Active]       = true;
  row[C.Updated_TS]   = new Date();
  row[C.Updated_By]   = perms.email;

  sh.getRange(rowIndex, 1, 1, width).setValues([row]);
  invalidateSheetCache_(t.sheet);

  recordChange_('RATE_BASE', row[C.Rate_ID], before, row, isNew ? 'CREATE' : 'UPDATE');
  return { ok: true, id: row[C.Rate_ID], isNew: isNew };
}


function deleteBaseRate(rateId) {
  prewarmForWrite_([SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS, SHEET.MODELLING_IDS]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const t = TABLES.RATE_BASE, C = COL.RATE_BASE, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.Rate_ID, rateId);
    if (rowIndex < 0) throw new Error('That rate no longer exists.');

    const before = readRow_(t.sheet, rowIndex, width);
    assertCanEditModellingId_(perms, safeInt(before[C.Modelling_ID]));

    // Soft delete: the row stays, so history and any audit trail stay intact.
    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('RATE_BASE', safeInt(before[C.Rate_ID]), before, row, 'DELETE');
    return { ok: true, id: safeInt(before[C.Rate_ID]) };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// SURCHARGES
// ─────────────────────────────────────────────────────────────────────────────

function saveSurchargeRate(p) {
  prewarmForWrite_([SHEET.RATE_SURCHARGE, SHEET.RATE_SURCHARGE_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.DIM_SURCHARGE]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const modellingId = safeInt(p.modellingId);
  assertCanEditModellingId_(perms, modellingId);

  const code = safeStr(p.code).toUpperCase();
  const def  = surchargeDefinition_(code);
  if (!def) throw new Error('Unknown surcharge type "' + code + '".');

  const d = validateDates_(p.validFrom, p.validTo);
  const value = safeNum(p.value);
  if (def.valueType === 'PCT' && (value < -1 || value > 2)) {
    throw new Error('A percentage surcharge should be a fraction — 0.14 for 14%.');
  }

  const t = TABLES.RATE_SURCHARGE, C = COL.RATE_SURCHARGE, width = t.headers.length;
  const isNew = !safeInt(p.id);

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);

    // The row's current owner, not just the one being asked for — before any
    // period is closed, so a refusal leaves nothing half-changed behind it.
    assertCanEditRowOwner_(perms, t.sheet, C.Surcharge_Rate_ID, p.id,
                           C.Modelling_ID, 'MODELLING_ID');

    if (p.closePrevious) closePrecedingPeriod_(t.sheet, C, C.Surcharge_Rate_ID,
      r => safeInt(r[C.Modelling_ID]) === modellingId &&
           safeStr(r[C.Surcharge_Code]).toUpperCase() === code,
      d.from, perms, 'RATE_SURCHARGE');

    invalidateSheetCache_(t.sheet);
    assertNoOverlap_(t.sheet, C,
      r => safeInt(r[C.Modelling_ID]) === modellingId &&
           safeStr(r[C.Surcharge_Code]).toUpperCase() === code,
      d.fromKey, d.toKey, p.id, C.Surcharge_Rate_ID,
      'Modelling ID ' + modellingId + ' ' + code);

    const sh = getSheet_(t.sheet);
    let rowIndex, before = null, row;

    if (isNew) {
      row = blankRow_('RATE_SURCHARGE');
      row[C.Surcharge_Rate_ID] = getNextId_(t.sheet, C.Surcharge_Rate_ID);
      row[C.Created_TS] = new Date();
      row[C.Created_By] = perms.email;
      rowIndex = sh.getLastRow() + 1;
    } else {
      rowIndex = findRowById_(t.sheet, C.Surcharge_Rate_ID, p.id);
      if (rowIndex < 0) throw new Error('That surcharge no longer exists.');
      before = readRow_(t.sheet, rowIndex, width);
      row = before.slice();
    }

    row[C.Modelling_ID]   = modellingId;
    row[C.Surcharge_Code] = code;
    row[C.Valid_From]     = d.from;
    row[C.Valid_To]       = d.to;
    row[C.Value]          = value;
    row[C.Currency]       = (def.valueType === 'AMT')
                            ? (safeStr(p.currency) || currencyForModellingId_(modellingId)) : '';
    row[C.Scenario_ID]    = safeInt(p.scenarioId) || 1;
    row[C.Source_Ref]     = safeStr(p.sourceRef);
    row[C.Notes]          = safeStr(p.notes);
    row[C.Active]         = true;
    row[C.Updated_TS]     = new Date();
    row[C.Updated_By]     = perms.email;

    sh.getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('RATE_SURCHARGE', row[C.Surcharge_Rate_ID], before, row,
                  isNew ? 'CREATE' : 'UPDATE');
    return { ok: true, id: row[C.Surcharge_Rate_ID], isNew: isNew };
  });
}


function deleteSurchargeRate(id) {
  prewarmForWrite_([SHEET.RATE_SURCHARGE, SHEET.RATE_SURCHARGE_AMENDS, SHEET.MODELLING_IDS]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const t = TABLES.RATE_SURCHARGE, C = COL.RATE_SURCHARGE, width = t.headers.length;

  return withLock_(function () {
    invalidateSheetCache_(t.sheet);
    const rowIndex = findRowById_(t.sheet, C.Surcharge_Rate_ID, id);
    if (rowIndex < 0) throw new Error('That surcharge no longer exists.');

    const before = readRow_(t.sheet, rowIndex, width);
    assertCanEditModellingId_(perms, safeInt(before[C.Modelling_ID]));

    const row = before.slice();
    row[C.Active]     = false;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    getSheet_(t.sheet).getRange(rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(t.sheet);

    recordChange_('RATE_SURCHARGE', safeInt(before[C.Surcharge_Rate_ID]), before, row, 'DELETE');
    return { ok: true, id: safeInt(before[C.Surcharge_Rate_ID]) };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// SHARED
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Shorten whatever period is running when a new one starts, so the two meet
 * exactly. This is what makes "add a rate change from 1 April" one click rather
 * than two edits — and it is the only correct behaviour now that overlapping
 * periods are a publish-blocking error.
 */
function closePrecedingPeriod_(sheetName, C, idCol, keyMatch, newFrom, perms, tableKey) {
  const data = getAllData_(sheetName);
  const width = TABLES[tableKey].headers.length;
  const sh = getSheet_(sheetName);

  precedingPeriodCloses_(data, C, keyMatch, newFrom).forEach(function (c) {
    const before = readRow_(sheetName, c.rowIndex, width);
    const row = before.slice();
    row[C.Valid_To]   = c.newTo;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    sh.getRange(c.rowIndex, 1, 1, width).setValues([row]);
    invalidateSheetCache_(sheetName);
    recordChange_(tableKey, safeInt(before[idCol]), before, row, 'UPDATE');
  });
}

/**
 * Which rows a new period starting `newFrom` should shorten, and to what date.
 *
 * The rule only — it reads nothing and writes nothing — so that the single-save
 * path above can apply it a row at a time against the sheet, and the bulk path
 * can apply it to an in-memory copy of the table for many routes before writing
 * anything. Splitting it out is what lets a bulk update be *the same rule*
 * applied many times rather than a second, similar-looking implementation.
 *
 * @param {Array}    data      the table including its header row
 * @param {Object}   C         that table's COL map
 * @param {Function} keyMatch  row -> is this row on the key we are changing
 * @param {Date}     newFrom   first day of the new period
 * @returns {Array} [{ rowIndex, newTo }] — rowIndex is 1-based, as the sheet sees it
 */
function precedingPeriodCloses_(data, C, keyMatch, newFrom) {
  const newFromKey = dateKey(newFrom);

  // The day before the new period starts. Plain local-date arithmetic —
  // constructing it from year/month/day-1 lets the Date object handle month
  // and year rollover, and never touches UTC or a timestamp.
  const newTo = new Date(newFrom.getFullYear(), newFrom.getMonth(), newFrom.getDate() - 1);
  const newToKey = dateKey(newTo);
  const out = [];

  for (let i = 1; i < data.length; i++) {
    if (!safeBool(data[i][C.Active])) continue;
    if (!keyMatch(data[i])) continue;
    const f = dateKey(data[i][C.Valid_From]), t = dateKey(data[i][C.Valid_To]);
    if (f >= newFromKey || t < newFromKey) continue;   // not the running period
    if (newToKey < f) continue;                        // would invert the period
    out.push({ rowIndex: i + 1, dataIndex: i, newTo: newTo });
  }
  return out;
}

function currencyForModellingId_(modellingId) {
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  let hlId = 0;
  for (let i = 1; i < md.length; i++) {
    if (safeInt(md[i][M.Modelling_ID]) === safeInt(modellingId)) {
      hlId = safeInt(md[i][M.High_Level_ID]); break;
    }
  }
  const hl = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < hl.length; i++) {
    if (safeInt(hl[i][H.High_Level_ID]) === hlId) return safeStr(hl[i][H.Currency]) || 'GBP';
  }
  return 'GBP';
}

function surchargeDefinition_(code) {
  const data = getAllData_(SHEET.DIM_SURCHARGE), C = COL.DIM_SURCHARGE;
  for (let i = 1; i < data.length; i++) {
    if (safeStr(data[i][C.Surcharge_Code]).toUpperCase() !== code) continue;
    return { code: code, valueType: safeStr(data[i][C.Value_Type]),
             proration: safeStr(data[i][C.Proration]) };
  }
  return null;
}


// ─────────────────────────────────────────────────────────────────────────────
// RESOLVED BY MONTH — what the engine will actually use
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Month-by-month working for one route: base, surcharges, rate per parcel and
 * the contribution it makes.
 *
 * This is the panel that makes proration visible. In the workbook, a fuel change
 * on the 4th of the month resolves to a blended figure buried inside a 2,500-row
 * SUMPRODUCT that nobody can check. Here it is a column you can read.
 */
function getResolvedByMonth(modellingId, scenarioId) {
  const perms = requirePermissions_();
  assertCanSeeModellingId_(perms, modellingId);

  const input  = loadEngineInput_(safeInt(scenarioId) || 1);
  const result = computeModel_(input);
  const want = safeInt(modellingId);

  const label = {};
  input.calendar.forEach(m => { label[m.dateId] = m.label; });

  return result.outputDetailRows
    .filter(r => r.modellingId === want)
    .sort((a, b) => a.dateId - b.dateId)
    .map(r => ({
      dateId: r.dateId, month: label[r.dateId] || r.dateId,
      monthStart: fmtDate(r.monthStart),
      baseRate: r.baseRate,
      surchargePct: r.surchargePctTotal,
      surchargeAmt: r.surchargeAmtTotal,
      ratePerParcel: r.ratePerParcel,
      methodMix: r.methodMix,
      lpMix: r.lpMix,
      contribution: r.rateContribution
    }));
}


// ─────────────────────────────────────────────────────────────────────────────
// BULK CHANGE
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Apply one uplift across many routes — "Royal Mail plus 6% from April".
 *
 * Always run with preview true first. The result lists exactly which routes
 * would change and by how much, and nothing is written.
 *
 * @param {Object} p { scope: 'CARRIER'|'HIGH_LEVEL_ID'|'MODELLING_IDS',
 *                     scopeValue, modellingIds[], changeType: 'PCT'|'AMOUNT'|'SET',
 *                     changeValue, validFrom, validTo, sourceRef,
 *                     preview: true|false, scenarioId }
 */
function bulkRateChange(p) {
  prewarmForWrite_([SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const d = validateDates_(p.validFrom, p.validTo);
  const scenarioId = safeInt(p.scenarioId) || 1;

  // ---- which routes -------------------------------------------------------
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const targets = [];
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id || !safeBool(md[i][M.Active])) continue;

    let match = false;
    if (p.scope === 'CARRIER') {
      match = normKey(md[i][M.Carrier_Code]) === normKey(p.scopeValue);
    } else if (p.scope === 'HIGH_LEVEL_ID') {
      match = safeInt(md[i][M.High_Level_ID]) === safeInt(p.scopeValue);
    } else if (p.scope === 'MODELLING_IDS') {
      match = (p.modellingIds || []).map(safeInt).indexOf(id) >= 0;
    }
    if (!match) continue;
    if (!canSeeModellingId_(perms, id, true)) continue;   // silently skip out of scope
    targets.push(id);
  }
  if (!targets.length) throw new Error('No routes matched, or none that you can edit.');

  // ---- what the new rate would be ----------------------------------------
  const rb = getAllData_(SHEET.RATE_BASE), C = COL.RATE_BASE;
  const plan = [];
  targets.forEach(mid => {
    let current = null;
    for (let i = 1; i < rb.length; i++) {
      if (safeInt(rb[i][C.Modelling_ID]) !== mid) continue;
      if (!safeBool(rb[i][C.Active])) continue;
      if (safeInt(rb[i][C.Scenario_ID]) !== scenarioId) continue;
      const f = dateKey(rb[i][C.Valid_From]), t = dateKey(rb[i][C.Valid_To]);
      if (f <= d.fromKey && t >= d.fromKey) { current = safeNum(rb[i][C.Base_Rate]); break; }
      if (f <= d.fromKey && (current === null)) current = safeNum(rb[i][C.Base_Rate]);
    }
    if (current === null) return;   // nothing to uplift

    let next;
    if (p.changeType === 'PCT')         next = current * (1 + safeNum(p.changeValue));
    else if (p.changeType === 'AMOUNT') next = current + safeNum(p.changeValue);
    else                                next = safeNum(p.changeValue);
    next = Math.round(next * 1000000) / 1000000;
    if (next < 0) throw new Error('That change would make Modelling ID ' + mid + ' negative.');

    plan.push({ modellingId: mid, from: current, to: next,
                diff: next - current, pct: current ? (next - current) / current * 100 : 0 });
  });

  if (p.preview !== false) {
    return { preview: true, count: plan.length, skipped: targets.length - plan.length,
             changes: plan };
  }

  // ---- write --------------------------------------------------------------
  // writeBaseRate_, not saveBaseRate: saveBaseRate takes the lock itself and
  // releases it in a finally-block, so calling it from inside this lock freed
  // the lock after the first route and left the rest of the batch unprotected.
  return withLock_(function () {
    let written = 0;
    plan.forEach(item => {
      writeBaseRate_({
        modellingId: item.modellingId,
        validFrom: fmtDate(d.from), validTo: fmtDate(d.to),
        value: item.to, scenarioId: scenarioId,
        sourceRef: safeStr(p.sourceRef) || 'Bulk change',
        notes: 'Bulk: ' + p.changeType + ' ' + p.changeValue,
        closePrevious: true
      }, perms);
      written++;
    });
    logAudit_('UPDATE', 'BULK_RATE_CHANGE', p.scope + ':' + (p.scopeValue || ''),
              '', '', String(written) + ' routes',
              p.changeType + ' ' + p.changeValue + ' from ' + fmtDate(d.from), true);
    return { preview: false, written: written, changes: plan };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// BULK UPDATE BY CHARGE TYPE AND DIMENSION
//
// The two selects at the top of the Rates screen edit one route at a time, which
// is right for "Royal Mail RM24 letters in GB went up". It is the wrong shape for
// "the 24-hour letter fuel surcharge went up everywhere", which is one commercial
// decision spread across dozens of routes and, done by hand, dozens of chances to
// miss one.
//
// A bulk change is therefore two questions. First WHICH CHARGE — the base rate,
// one named surcharge, or all of them. Then WHICH ROUTES — name a value for any
// dimension, or leave it as All. The change lands on every EXISTING row that
// matches. "Existing" is the operative word: this updates rows that are already
// there, it never creates a rate for a route that had none. A route with no rate
// is a structural gap that wants noticing on the Rates screen, not filling in
// silently by a bulk action.
//
// THE DIMENSIONS. Neither rate table stores any of them. Both key on
// Modelling_ID, and every dimension is reached by a two-hop join:
//
//   Rate_Base / Rate_Surcharge  ->  Modelling_IDs  ->  High_Level_IDs
//
//   High_Level_IDs : Brand, Geo, Treatment_Type, WL_Split
//   Modelling_IDs  : Method_Code, Letter_Parcel
//
// Modelling_IDs.Carrier_Code is deliberately NOT selectable. It is a genuinely
// separate column — migrated from its own source column, and free to disagree
// with the owning carrier that Dim_Method records for the method (validation only
// WARNs, as METHOD_CARRIER) — but it was removed from the selector by request, so
// a bulk change always spans every carrier. Nothing here reads it, rather than
// defaulting it, so a hand-built payload cannot narrow by carrier either.
//
// ALL IS NOT A STORED VALUE. It exists only in the selector, as the empty string.
// Nothing is written as "All" — every row written is an ordinary row against one
// concrete Modelling_ID. Note that WL_Split legitimately stores '*' meaning "not
// applicable" for segments that are not weight-loss; that is a real value to be
// matched, NOT a wildcard, and conflating the two would silently widen every
// change.
//
// TWO MODES, BECAUSE ONE VALUE CANNOT MEAN TWO THINGS. A base rate is currency
// per parcel, FUEL is a fraction and the other surcharges are currency amounts
// (Dim_Surcharge.Value_Type). So:
//
//   one charge type  ->  SET: every matching row takes the value entered
//   All charge types ->  PCT: every matching row moves by the percentage entered
//
// A single value across mixed charge types would set a base rate to £0.14 and a
// fuel surcharge to 14% from the same keystroke. A percentage uplift means the
// same thing to all of them, so that is what All offers.
//
// SAME RULES, BATCHED WRITES. Every check a single save makes is made here, per
// row, by the same functions: assertCanEditModellingId_ for scope,
// precedingPeriodCloses_ for the period-closing rule, assertNoOverlapIn_ for
// overlaps, and validateChargeValue_ for the value. What changes is only how the
// result reaches the sheet. A single save costs about seven Sheets API calls;
// Rate_Surcharge holds some 1,650 rows and a wide selection can match hundreds of
// them, which at seven calls each would pass the six-minute execution limit and
// leave a half-applied price change behind. So the plan is built in memory, one
// working copy per table, then each table is flushed as one ranged write for the
// closed periods, one append for the new rows, and one recordChangesBatch_ for
// the history — which produces the same Amends snapshots and the same per-field
// Audit_Log rows the single path does.
//
// PREVIEW FIRST, ALWAYS. preview defaults to true and returns the matched rows
// with their current values; applying requires passing back the planKey the
// preview issued. If anything about the matched set has moved in between — an
// edit by somebody else, a route deactivated — the key no longer matches and the
// write is refused rather than applied to a set the user never saw.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Does a dimension selection match a row's value?
 *
 * Nothing selected means All, and so does a list with nothing usable in it.
 * There is deliberately no magic string: a literal 'ALL' is treated as a value
 * to match, because a brand code called ALL is one data operation away, and a
 * wildcard that a data entry can impersonate is a wildcard waiting to widen
 * someone's rate change.
 *
 * A list matches the UNION of its values — tick two methods and you get the
 * routes on either, not the routes on both, which no route could satisfy. The
 * dimensions are still ANDed with each other, so ticking methods narrows methods
 * and leaves every treatment type, brand and geo carrying them in the set.
 */
function dimensionMatches_(selected, rowValue) {
  const wanted = dimensionWanted_(selected);
  if (!wanted.length) return true;
  const row = normKey(rowValue);
  for (let i = 0; i < wanted.length; i++) if (wanted[i] === row) return true;
  return false;
}

/**
 * A dimension selection normalised to the list of codes it matches.
 *
 * Accepts one code or a list of them, because the bulk form now ticks boxes.
 * Every box ticked and no box ticked both arrive as an empty list and both mean
 * All — the form refuses to submit with none ticked, so an empty list can only
 * ever have meant "everything", never "nothing". Blanks are dropped rather than
 * matched and duplicates collapse, so ['RM48', '', 'rm48'] and ['RM48'] plan the
 * same write.
 */
function dimensionWanted_(selected) {
  const raw = Array.isArray(selected) ? selected : [selected];
  const out = [];
  for (let i = 0; i < raw.length; i++) {
    const v = normKey(raw[i]);
    if (v !== '' && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

/**
 * The active, editable Modelling IDs matching a dimension selection.
 *
 * Carrier_Code is not consulted — see the note above about why it is not
 * selectable. Routes the caller cannot edit are counted and skipped rather than
 * refused, the same choice bulkRateChange makes, so a scoped user gets their own
 * slice of a company-wide change instead of an error naming segments they cannot
 * see.
 */
function resolveDimensionTargets_(dims, perms) {
  const sel = dims || {};
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

  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const out = { ids: [], routesMatched: 0, skippedNoScope: 0 };
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id || !safeBool(md[i][M.Active])) continue;
    if (!hlOk[safeInt(md[i][M.High_Level_ID])]) continue;
    if (!dimensionMatches_(sel.method,       md[i][M.Method_Code]))   continue;
    if (!dimensionMatches_(sel.letterParcel, md[i][M.Letter_Parcel])) continue;

    out.routesMatched++;
    if (!canSeeModellingId_(perms, id, true)) { out.skippedNoScope++; continue; }
    out.ids.push(id);
  }
  return out;
}

/** A human description of a selection, for the audit trail and the preview. */
function describeDimensions_(dims, chargeType) {
  const sel = dims || {};
  const parts = ['charge=' + (normKey(chargeType) || 'ALL')];
  [['brand', 'brand'], ['geo', 'geo'], ['treatmentType', 'treatment'],
   ['wlSplit', 'WL split'], ['method', 'method'],
   ['letterParcel', 'class']].forEach(function (pair) {
    const wanted = dimensionWanted_(sel[pair[0]]);
    /* Joined with + because the parts themselves are comma-joined, and a comma
       here would read as another dimension in the audit trail. Sorted so the
       same set ticked in a different order writes the same description. */
    parts.push(pair[1] + '=' + (wanted.length ? wanted.slice().sort().join('+') : 'All'));
  });
  parts.push('carrier=All');   // stated, because it is not selectable
  return parts.join(', ');
}

/** Active surcharge codes, in Dim_Surcharge order. Never a hardcoded list. */
function activeSurchargeCodes_() {
  const data = getAllData_(SHEET.DIM_SURCHARGE), C = COL.DIM_SURCHARGE, out = [];
  for (let i = 1; i < data.length; i++) {
    const code = normKey(data[i][C.Surcharge_Code]);
    if (!code || !safeBool(data[i][C.Active])) continue;
    out.push(code);
  }
  return out;
}

/**
 * The charge types a user may pick, for the picker the client shows first.
 *
 * Base rate is not in Dim_Surcharge — it is a different table — so it is prepended
 * here rather than stored anywhere. Everything else is read live, so a surcharge
 * type added next year appears with no code change.
 */
function listChargeTypes() {
  const perms = requirePermissions_();
  requireEditRates_(perms);
  const out = [{ chargeType: 'BASE', name: 'Base rate', valueType: 'RATE',
                 expressed: 'rate per parcel' }];
  const data = getAllData_(SHEET.DIM_SURCHARGE), C = COL.DIM_SURCHARGE;
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const code = normKey(data[i][C.Surcharge_Code]);
    if (!code || !safeBool(data[i][C.Active])) continue;
    const vt = normKey(data[i][C.Value_Type]);
    rows.push({ chargeType: code, name: safeStr(data[i][C.Surcharge_Name]) || code,
                valueType: vt, order: safeInt(data[i][C.Apply_Order]),
                expressed: vt === 'PCT' ? 'percentage of base' : 'fixed amount' });
  }
  rows.sort(function (a, b) { return a.order - b.order || a.chargeType.localeCompare(b.chargeType); });
  return out.concat(rows);
}

/** Which table a charge type is against, and how its rows are keyed. */
function bulkRateShape_(chargeType) {
  const ct = normKey(chargeType);
  if (ct && ct !== 'BASE') {
    const SC = COL.RATE_SURCHARGE;
    return {
      kind: 'SURCHARGE', chargeType: ct, code: ct,
      tableKey: 'RATE_SURCHARGE', table: TABLES.RATE_SURCHARGE,
      C: SC, idCol: SC.Surcharge_Rate_ID, valueCol: SC.Value,
      keyMatch: function (mid) {
        return function (r) {
          return safeInt(r[SC.Modelling_ID]) === mid &&
                 normKey(r[SC.Surcharge_Code]) === ct;
        };
      },
      label: function (mid) { return 'Modelling ID ' + mid + ' ' + ct; }
    };
  }
  const BC = COL.RATE_BASE;
  return {
    kind: 'BASE', chargeType: 'BASE', code: '',
    tableKey: 'RATE_BASE', table: TABLES.RATE_BASE,
    C: BC, idCol: BC.Rate_ID, valueCol: BC.Base_Rate,
    keyMatch: function (mid) {
      return function (r) { return safeInt(r[BC.Modelling_ID]) === mid; };
    },
    label: function (mid) { return 'Modelling ID ' + mid; }
  };
}

/**
 * The charge types one request covers: one, or every one there is.
 *
 * 'ALL' expands to the base rate plus every active surcharge, so a single request
 * can span both rate tables. There is no default — an empty charge type is a
 * refusal, because the client's first step exists precisely to make the user
 * choose one.
 */
function chargeTypeGroups_(chargeType) {
  const ct = normKey(chargeType);
  if (!ct) throw new Error('Choose which charge you are changing before going on.');
  if (ct === 'ALL') {
    return [bulkRateShape_('BASE')].concat(
      activeSurchargeCodes_().map(function (c) { return bulkRateShape_(c); }));
  }
  if (ct !== 'BASE' && !surchargeDefinition_(ct)) {
    throw new Error('Unknown charge type "' + ct + '".');
  }
  return [bulkRateShape_(ct)];
}

/**
 * The single-save value rules, applied to a bulk row.
 *
 * Identical to the checks in saveBaseRate and saveSurchargeRate, in one place so
 * both the set-to and the uplift mode are held to them — an uplift can just as
 * easily drive a rate negative as a typed value can.
 */
function validateChargeValue_(shape, value, whose) {
  const where = whose ? ' (' + whose + ')' : '';
  if (shape.kind === 'BASE') {
    if (value < 0) throw new Error('A base rate cannot be negative' + where + '.');
    return value;
  }
  const def = surchargeDefinition_(shape.code);
  if (!def) throw new Error('Unknown surcharge type "' + shape.code + '".');
  if (def.valueType === 'PCT' && (value < -1 || value > 2)) {
    throw new Error('A percentage surcharge should be a fraction — 0.14 for 14%' +
                    where + '.');
  }
  return value;
}

function roundRate_(v) { return Math.round(v * 1000000) / 1000000; }

/**
 * Work out what a bulk update would do, without touching anything.
 *
 * For each charge type and each matching route this finds the row a new period
 * would supersede: the one in force on the new start date, or failing that the
 * latest one starting before it. A route with no such row is reported as skipped,
 * not created.
 *
 * Scenario_ID is deliberately NOT part of the match, because the single-save path
 * does not filter on it either (finding C5). With one scenario the two are
 * identical; when a second exists, C5 has to be fixed across all of these paths
 * in one go rather than diverging here first.
 */
function planBulkRateUpdate_(p, perms) {
  const chargeType = normKey(p.chargeType);
  const groups = chargeTypeGroups_(chargeType);
  const mode = chargeType === 'ALL' ? 'PCT' : 'SET';
  const entered = safeNum(p.value);

  if (mode === 'PCT' && entered <= -100) {
    throw new Error('An uplift of -100% or less would take every charge to zero ' +
                    'or below. Enter a percentage above -100.');
  }
  if (mode === 'SET') validateChargeValue_(groups[0], entered);

  const d = validateDates_(p.validFrom, p.validTo);
  const targets = resolveDimensionTargets_(p.dimensions, perms);
  const factor = 1 + entered / 100;

  let count = 0, skippedNoRate = 0;
  const planned = groups.map(function (shape) {
    const C = shape.C;
    const data = getAllData_(shape.table.sheet);
    const items = [];

    targets.ids.forEach(function (mid) {
      const matches = shape.keyMatch(mid);
      let cur = null;
      for (let i = 1; i < data.length; i++) {
        if (!safeBool(data[i][C.Active])) continue;
        if (!matches(data[i])) continue;
        const f = dateKey(data[i][C.Valid_From]), t = dateKey(data[i][C.Valid_To]);
        const row = { id: safeInt(data[i][shape.idCol]), fromKey: f,
                      value: safeNum(data[i][shape.valueCol]),
                      validFrom: fmtDate(data[i][C.Valid_From]),
                      validTo: fmtDate(data[i][C.Valid_To]) };
        if (f <= d.fromKey && t >= d.fromKey) { cur = row; break; }   // in force
        if (f <= d.fromKey && (!cur || f > cur.fromKey)) cur = row;   // latest before
      }
      if (!cur) { skippedNoRate++; return; }

      const to = roundRate_(mode === 'PCT' ? cur.value * factor : entered);
      validateChargeValue_(shape, to, shape.label(mid));
      items.push({ modellingId: mid, rateId: cur.id, from: cur.value, to: to,
                   diff: roundRate_(to - cur.value),
                   pct: cur.value ? (to - cur.value) / cur.value * 100 : 0,
                   currentFrom: cur.validFrom, currentTo: cur.validTo });
    });

    count += items.length;
    return { shape: shape, items: items, distinct: distinctMoves_(items) };
  });

  return {
    chargeType: chargeType, mode: mode, entered: entered, dates: d,
    groups: planned.filter(function (g) { return g.items.length > 0; }),
    emptyGroups: planned.filter(function (g) { return !g.items.length; }).length,
    count: count, routesMatched: targets.routesMatched,
    skippedNoScope: targets.skippedNoScope, skippedNoRate: skippedNoRate,
    planKey: planKey_(planned), selection: describeDimensions_(p.dimensions, chargeType)
  };
}

/**
 * The distinct from -> to moves in a group, commonest first.
 *
 * Selection is by dimension alone, so showing the spread of what is about to be
 * overwritten is the safeguard rather than a filter on it. In set-to mode every
 * `to` is the same and this reads as "what is being flattened"; in uplift mode
 * each `from` has its own `to`.
 */
function distinctMoves_(items) {
  const seen = {};
  items.forEach(function (it) {
    const k = String(it.from) + '>' + String(it.to);
    if (!seen[k]) seen[k] = { from: it.from, to: it.to, rows: 0 };
    seen[k].rows++;
  });
  return Object.keys(seen).map(function (k) { return seen[k]; })
    .sort(function (a, b) { return b.rows - a.rows || a.from - b.from; });
}

/**
 * A short fingerprint of a plan's matched rows and their current values.
 *
 * The client previews, the user confirms, and the apply call sends this back. If
 * the matched set has moved in between — someone else edited a rate, a route was
 * deactivated — the recomputed key differs and the write is refused, so a
 * confirmation can only ever apply the set that was actually shown. Keyed on the
 * charge type too, so a plan cannot be replayed against a different charge.
 */
function planKey_(groups) {
  const parts = [];
  groups.forEach(function (g) {
    g.items.forEach(function (it) {
      parts.push(g.shape.chargeType + ':' + it.rateId + ':' +
                 Math.round(it.from * 1000000) + ':' + Math.round(it.to * 1000000));
    });
  });
  const s = parts.sort().join('|');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s || 'empty');
  return bytes.map(function (b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('').slice(0, 12);
}

/**
 * Bulk update one charge type, or every charge type, across a dimension selection.
 *
 * Previews by default. Pass preview:false together with the planKey the preview
 * returned to actually write.
 *
 * @param {Object} p { chargeType: 'BASE' | <surcharge code> | 'ALL',
 *                     dimensions: { brand, geo, treatmentType, wlSplit,
 *                                   method, letterParcel },
 *                     validFrom, validTo,
 *                     value,        // set-to value, or an uplift % when ALL
 *                     currency, sourceRef, scenarioId, preview, planKey }
 */
function bulkUpdateRates(p) {
  prewarmForWrite_([SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS,
                    SHEET.RATE_SURCHARGE, SHEET.RATE_SURCHARGE_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_SURCHARGE]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const plan = planBulkRateUpdate_(p, perms);

  if (p.preview !== false) {
    return {
      preview: true, chargeType: plan.chargeType, mode: plan.mode,
      value: plan.entered, count: plan.count,
      routesMatched: plan.routesMatched, skippedNoScope: plan.skippedNoScope,
      skippedNoRate: plan.skippedNoRate, planKey: plan.planKey,
      selection: plan.selection,
      validFrom: fmtDate(plan.dates.from), validTo: fmtDate(plan.dates.to),
      groups: plan.groups.map(function (g) {
        const def = g.shape.kind === 'SURCHARGE' ? surchargeDefinition_(g.shape.code) : null;
        return { chargeType: g.shape.chargeType, kind: g.shape.kind,
                 valueType: def ? def.valueType : 'RATE',
                 count: g.items.length, distinct: g.distinct };
      })
    };
  }

  if (!plan.count) throw new Error('Nothing matches that selection, so there is ' +
    'nothing to update. ' + plan.skippedNoRate + ' matching charge(s) have no rate ' +
    'to supersede, and ' + plan.skippedNoScope + ' route(s) are outside what you can edit.');

  if (safeStr(p.planKey) !== plan.planKey) {
    throw new Error('These rates have changed since the preview was taken, so ' +
      'nothing was applied. Run the preview again and check it still says what ' +
      'you expect.');
  }

  return withLock_(function () { return applyBulkRateUpdate_(p, plan, perms); });
}

/**
 * Write a planned bulk update. Assumes the caller holds the lock.
 *
 * Runs the same per-row rules as a single save against one working copy per
 * table, then flushes each table once. The copy is mutated as it goes — closing a
 * period for one route before planning the next — so each route's overlap check
 * sees the closes and the new rows that earlier routes in the same batch
 * produced. One copy PER TABLE rather than per charge type, because "All" puts
 * several surcharge groups on Rate_Surcharge and separate copies would each flush
 * over the last one's work.
 */
function applyBulkRateUpdate_(p, plan, perms) {
  const now = new Date();
  const batchRef = bulkBatchRef_(now);
  const scenarioId = safeInt(p.scenarioId) || 1;
  const ctx = {};

  function workingSet(shape) {
    const name = shape.table.sheet;
    if (!ctx[name]) {
      invalidateSheetCache_(name);
      const width = shape.table.headers.length;
      const source = getAllData_(name);
      ctx[name] = {
        sheetName: name, tableKey: shape.tableKey, width: width, idCol: shape.idCol,
        originalRows: source.length - 1,
        data: source.map(function (r) { return padRow_(r, width); }),
        changed: {}, creates: [], history: [],
        nextId: getNextId_(name, shape.idCol)
      };
    }
    return ctx[name];
  }

  plan.groups.forEach(function (g) {
    const shape = g.shape, C = shape.C;
    const ws = workingSet(shape);

    g.items.forEach(function (item) {
      const mid = item.modellingId;
      assertCanEditModellingId_(perms, mid);     // per row, as a single save does
      const matches = shape.keyMatch(mid);

      // 1. close whatever period is running — the same rule the single save uses
      precedingPeriodCloses_(ws.data, C, matches, plan.dates.from).forEach(function (c) {
        const before = ws.data[c.dataIndex].slice();
        const after = before.slice();
        after[C.Valid_To]   = c.newTo;
        after[C.Updated_TS] = now;
        after[C.Updated_By] = perms.email;
        ws.data[c.dataIndex] = after;
        ws.changed[c.dataIndex] = true;
        ws.history.push({ id: safeInt(before[shape.idCol]), before: before,
                          after: after, type: 'UPDATE' });
      });

      // 2. refuse an overlap — the same rule, against the copy as it now stands
      assertNoOverlapIn_(ws.data, C, matches, plan.dates.fromKey, plan.dates.toKey,
                         0, shape.idCol, shape.label(mid));

      // 3. the new period
      const row = blankRow_(shape.tableKey);
      row[shape.idCol]     = ws.nextId++;
      row[C.Modelling_ID]  = mid;
      if (shape.kind === 'SURCHARGE') row[C.Surcharge_Code] = shape.code;
      row[C.Valid_From]    = plan.dates.from;
      row[C.Valid_To]      = plan.dates.to;
      row[shape.valueCol]  = item.to;
      row[C.Currency]      = bulkRowCurrency_(shape, mid, p);
      row[C.Scenario_ID]   = scenarioId;
      row[C.Source_Ref]    = safeStr(p.sourceRef) || batchRef;
      row[C.Notes]         = truncate_('Bulk ' + batchRef + ': ' +
                                       describeMove_(plan, item) + ' (' + plan.selection + ')');
      row[C.Active]        = true;
      row[C.Created_TS]    = now;
      row[C.Created_By]    = perms.email;
      row[C.Updated_TS]    = now;
      row[C.Updated_By]    = perms.email;

      ws.creates.push(row);
      ws.data.push(row);      // so the next route's overlap check sees it
      ws.history.push({ id: row[shape.idCol], before: null, after: row, type: 'CREATE' });
    });
  });

  // ---- flush, one table at a time -----------------------------------------
  // One ranged write for the closed periods rather than one per row. Rows come
  // back from batchGet with trailing empty cells omitted, which is why the whole
  // working copy was padded to the header width up front — finding C6, which is
  // exactly this mistake made in structure.gs.
  let rowsClosed = 0, rowsCreated = 0;
  Object.keys(ctx).forEach(function (name) {
    const ws = ctx[name];
    const sh = getSheet_(name);
    if (Object.keys(ws.changed).length) {
      sh.getRange(2, 1, ws.originalRows, ws.width)
        .setValues(ws.data.slice(1, 1 + ws.originalRows));
    }
    if (ws.creates.length) {
      sh.getRange(sh.getLastRow() + 1, 1, ws.creates.length, ws.width)
        .setValues(ws.creates);
    }
    invalidateSheetCache_(name);

    // The same Amends snapshots and the same per-field Audit_Log rows a single
    // save makes, in a handful of calls instead of two per row, each stamped with
    // the batch reference so History can show them as one action.
    recordChangesBatch_(ws.tableKey, ws.history, batchRef);
    rowsClosed  += Object.keys(ws.changed).length;
    rowsCreated += ws.creates.length;
  });

  logAudit_('UPDATE', 'BULK_RATE_UPDATE', batchRef, '', '',
            plan.count + ' rows, ' + describeChange_(plan),
            'from ' + fmtDate(plan.dates.from) + ' — ' + plan.selection, true);

  return { preview: false, batchRef: batchRef, written: plan.count,
           mode: plan.mode, chargeType: plan.chargeType,
           rowsClosed: rowsClosed, rowsCreated: rowsCreated,
           skippedNoScope: plan.skippedNoScope, skippedNoRate: plan.skippedNoRate,
           groups: plan.groups.map(function (g) {
             return { chargeType: g.shape.chargeType, written: g.items.length }; }) };
}

/** "set to 0.165", or "0.14 +6% -> 0.1484", for a row's Notes. */
function describeMove_(plan, item) {
  if (plan.mode === 'SET') return 'set to ' + item.to;
  return item.from + ' ' + (plan.entered >= 0 ? '+' : '') + plan.entered + '% -> ' + item.to;
}

/** "set to 0.165" or "+6% across every charge type", for the summary audit row. */
function describeChange_(plan) {
  if (plan.mode === 'SET') return 'set to ' + plan.entered;
  return (plan.entered >= 0 ? '+' : '') + plan.entered + '% across every charge type';
}

/**
 * The currency for a bulk-written row.
 *
 * Base rates and amount surcharges carry one; percentage surcharges must not, or
 * the Rates screen shows "0.14 GBP" for what is 14%. Mirrors what saveBaseRate
 * and saveSurchargeRate each do for a single row.
 */
function bulkRowCurrency_(shape, mid, p) {
  if (shape.kind === 'SURCHARGE') {
    const def = surchargeDefinition_(shape.code);
    if (!def || def.valueType !== 'AMT') return '';
  }
  return safeStr(p.currency) || currencyForModellingId_(mid);
}

/**
 * A batch reference shared by every row of one bulk action.
 *
 * Timestamp-based rather than a counter: it needs to be unique, sortable and
 * recognisable in a Source_Ref cell months later, and it should not cost a lock
 * or a property read on a path that already has plenty to do.
 *
 * Built from the date's own components rather than Utilities.formatDate, for the
 * reason given above fmtDate in utils.gs — the runtime's local timezone already
 * is the project timezone, so the service call buys nothing.
 */
function bulkBatchRef_(when) {
  return 'BULK-' + when.getFullYear() + pad2_(when.getMonth() + 1) + pad2_(when.getDate()) +
         '-' + pad2_(when.getHours()) + pad2_(when.getMinutes()) + pad2_(when.getSeconds());
}

/** Pad a row read from batchGet out to a table's full width. See finding C6. */
function padRow_(row, width) {
  const out = (row || []).slice(0, width);
  while (out.length < width) out.push('');
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Full create, read, update, delete cycle on a real route, then puts everything
 * back. Proves the write path, the guards and the history all work together.
 */
function testRateWrite() {
  requireMaintenance_();
  Logger.log('=== RATE WRITE TEST ===');
  const perms = requirePermissions_();
  Logger.log('  acting as ' + perms.email + ' (' + perms.role + ')');

  const TEST_MID = 1;
  const before = Math.max(getSheet_(SHEET.RATE_BASE).getLastRow() - 1, 0);
  Logger.log('  Rate_Base rows before: ' + before);

  Logger.log('');
  Logger.log('--- create a rate in a period nothing else uses (2035) ---');
  const created = saveBaseRate({
    modellingId: TEST_MID, validFrom: '2035-01-01', validTo: '2035-12-31',
    value: 9.99, sourceRef: 'testRateWrite', notes: 'TEST — safe to delete'
  });
  Logger.log('  created Rate_ID ' + created.id);

  Logger.log('');
  Logger.log('--- overlap must be refused ---');
  let refused = false;
  try {
    saveBaseRate({ modellingId: TEST_MID, validFrom: '2035-06-01', validTo: '2035-08-31',
                   value: 1.11, notes: 'should be refused' });
  } catch (e) {
    refused = true;
    Logger.log('  refused, correctly: ' + e.message);
  }
  if (!refused) Logger.log('  PROBLEM — the overlap was allowed');

  Logger.log('');
  Logger.log('--- update it ---');
  saveBaseRate({ id: created.id, modellingId: TEST_MID,
                 validFrom: '2035-01-01', validTo: '2035-12-31',
                 value: 8.88, sourceRef: 'testRateWrite', notes: 'TEST — updated' });
  Logger.log('  changed 9.99 to 8.88');

  Logger.log('');
  Logger.log('--- history ---');
  invalidateSheetCache_(SHEET.RATE_BASE_AMENDS);
  readRecordHistory_('RATE_BASE', created.id).forEach(h =>
    Logger.log('  ' + h.ts + '  ' + h.type + '  Base_Rate=' + h.values.Base_Rate + '  by ' + h.by));

  Logger.log('');
  Logger.log('--- delete it ---');
  deleteBaseRate(created.id);
  Logger.log('  marked inactive (the row stays, so history survives)');

  const after = Math.max(getSheet_(SHEET.RATE_BASE).getLastRow() - 1, 0);
  Logger.log('');
  Logger.log('  Rate_Base rows after: ' + after + '  (+' + (after - before) + ', now inactive)');

  Logger.log('');
  Logger.log('--- resolved by month, Modelling ID ' + TEST_MID + ' ---');
  getResolvedByMonth(TEST_MID, 1).slice(0, 3).forEach(r =>
    Logger.log('  ' + r.month + '  base ' + r.baseRate.toFixed(4) +
               '  fuel ' + (r.surchargePct * 100).toFixed(4) + '%' +
               '  other ' + r.surchargeAmt.toFixed(4) +
               '  rate ' + r.ratePerParcel.toFixed(4)));

  Logger.log('');
  Logger.log(refused ? 'RATE WRITES WORKING' : 'PROBLEM — overlap protection did not fire');
  Logger.log('');
  Logger.log('One inactive test row was left in Rate_Base (Rate_ID ' + created.id +
             ', year 2035). It affects no forecast. Delete it whenever you like.');
  return { ok: refused, testRateId: created.id };
}

/**
 * End-to-end test of the dimension-based bulk update, which puts the table back
 * exactly as it found it.
 *
 * Unlike testRateWrite, this one cannot confine itself to a spare year: closing
 * the preceding period is the whole point of the feature, so applying it edits
 * Valid_To on live rows. So it snapshots the entire table first, and afterwards
 * restores every row it changed byte for byte and deactivates every row it
 * created. Nothing is left behind, which is what makes it safe to re-run.
 *
 * It exercises the surcharge path, because that is the harder of the two: it has
 * the extra Surcharge_Code key and the percentage-versus-amount rule.
 */
/**
 * End-to-end test of the bulk update, which puts the tables back exactly as it
 * found them.
 *
 * Unlike testRateWrite, this one cannot confine itself to a spare year: closing
 * the preceding period is the whole point of the feature, so applying it edits
 * Valid_To on live rows. So it snapshots every table it will touch first, and
 * afterwards restores each changed row byte for byte and deactivates each row it
 * created. Nothing is left behind, which is what makes it safe to re-run.
 *
 * It covers the set-to mode on one surcharge type and the uplift mode across all
 * charge types, since the second is the only path that writes to both rate tables
 * in one action.
 */
function testBulkRateUpdate() {
  requireMaintenance_();
  Logger.log('=== BULK RATE UPDATE TEST ===');
  const perms = requirePermissions_();
  Logger.log('  acting as ' + perms.email + ' (' + perms.role + ')');

  const checks = [];
  function check(name, pass, detail) {
    checks.push({ name: name, pass: !!pass });
    Logger.log('  ' + (pass ? 'ok   ' : 'FAIL ') + name + (detail ? '  — ' + detail : ''));
  }

  // ---- 1. the charge types on offer ---------------------------------------
  Logger.log('');
  Logger.log('--- charge types, read from Dim_Surcharge ---');
  const types = listChargeTypes();
  types.forEach(function (t) {
    Logger.log('  ' + t.chargeType + '  ' + t.name + '  (' + t.expressed + ')');
  });
  check('base rate is offered first', types.length > 0 && types[0].chargeType === 'BASE');
  check('at least one surcharge type is offered', types.length > 1);
  check('every charge type has a name and a value type',
        types.every(function (t) { return !!t.name && !!t.valueType; }));
  check('no charge type is called ALL', !types.some(function (t) { return t.chargeType === 'ALL'; }));

  // ---- 2. dimension resolution -------------------------------------------
  Logger.log('');
  Logger.log('--- dimension resolution ---');
  const all = resolveDimensionTargets_({}, perms);
  Logger.log('  every dimension All        : ' + all.ids.length + ' editable routes (' +
             all.routesMatched + ' matched, ' + all.skippedNoScope + ' out of scope)');
  check('All-everything matches routes', all.ids.length > 0);

  const letters = resolveDimensionTargets_({ letterParcel: 'LETTER' }, perms);
  Logger.log('  class=LETTER               : ' + letters.ids.length);
  check('naming a dimension narrows the set', letters.ids.length <= all.ids.length);

  // Carrier is no longer selectable: passing one must not narrow anything.
  const withCarrier = resolveDimensionTargets_({ carrier: 'ROYALMAIL' }, perms);
  Logger.log('  carrier=ROYALMAIL (ignored): ' + withCarrier.ids.length +
             '  (must equal the All-everything count)');
  check('a carrier in the payload is ignored, not honoured',
        withCarrier.ids.length === all.ids.length);

  const star = resolveDimensionTargets_({ wlSplit: '*' }, perms);
  Logger.log("  wlSplit='*'                : " + star.ids.length +
             '  (a value, not a wildcard)');
  check("WL_Split '*' is matched as a value", star.ids.length < all.ids.length);
  check('an unknown value matches nothing',
        resolveDimensionTargets_({ brand: 'NO_SUCH_BRAND' }, perms).ids.length === 0);

  // Ticked boxes arrive as a list. The cases that matter are the ones where a
  // list could quietly mean something other than what was ticked.
  const letterList = resolveDimensionTargets_({ letterParcel: ['LETTER'] }, perms);
  Logger.log('  class=[LETTER]             : ' + letterList.ids.length);
  check('a one-item list matches the same as naming that value',
        letterList.ids.length === letters.ids.length);
  check('an empty list means All, because that is how "every box ticked" arrives',
        resolveDimensionTargets_({ letterParcel: [] }, perms).ids.length === all.ids.length);
  check('a list of blanks means All, not nothing',
        resolveDimensionTargets_({ letterParcel: ['', '  '] }, perms).ids.length === all.ids.length);
  check('duplicates and case in a list neither widen nor narrow the set',
        resolveDimensionTargets_({ letterParcel: ['LETTER', 'letter'] }, perms).ids.length ===
        letters.ids.length);

  const bothClasses = resolveDimensionTargets_({ letterParcel: ['LETTER', 'PARCEL'] }, perms);
  Logger.log('  class=[LETTER,PARCEL]      : ' + bothClasses.ids.length +
             '  (the union, not the intersection)');
  check('a list matches the union of its values',
        bothClasses.ids.length >= letters.ids.length &&
        bothClasses.ids.length <= all.ids.length);
  check('an unknown value alongside a known one does not suppress the known one',
        resolveDimensionTargets_({ letterParcel: ['LETTER', 'NO_SUCH_CLASS'] }, perms).ids.length ===
        letters.ids.length);

  // Ticking methods must leave every other dimension wide open — the whole point
  // of the tick boxes is that the untouched dimensions still mean All.
  // Taken off a live route rather than out of Dim_Method, so the method is known
  // to carry traffic and the count below can be asserted as non-zero.
  const md = getAllData_(SHEET.MODELLING_IDS), MD = COL.MODELLING_IDS;
  let oneMethod = '';
  for (let i = 1; i < md.length && !oneMethod; i++) {
    if (safeBool(md[i][MD.Active])) oneMethod = normKey(md[i][MD.Method_Code]);
  }
  if (oneMethod) {
    const byMethod = resolveDimensionTargets_({ method: [oneMethod] }, perms);
    Logger.log('  method=[' + oneMethod + ']: ' + byMethod.ids.length);
    check('a method list narrows methods without narrowing anything else',
          byMethod.ids.length > 0 && byMethod.ids.length <= all.ids.length &&
          byMethod.ids.length ===
            resolveDimensionTargets_({ method: [oneMethod], brand: [], geo: [],
                                       treatmentType: [], wlSplit: [] }, perms).ids.length);
  }

  check('a list is described for the audit trail, sorted and +-joined',
        describeDimensions_({ letterParcel: ['PARCEL', 'letter'] }, 'BASE')
          .indexOf('class=LETTER+PARCEL') > -1);
  check('an empty list is described as All',
        describeDimensions_({ letterParcel: [] }, 'BASE').indexOf('class=All') > -1);

  // ---- 3. a charge type must be chosen -----------------------------------
  Logger.log('');
  Logger.log('--- a charge type is compulsory ---');
  let noneRefused = false, badRefused = false;
  const baseReq = { dimensions: { letterParcel: 'LETTER' },
                    validFrom: '2035-01-01', validTo: '9999-12-31',
                    sourceRef: 'testBulkRateUpdate' };
  try { bulkUpdateRates(Object.assign({}, baseReq, { chargeType: '', value: 1 })); }
  catch (e) { noneRefused = true; Logger.log('  refused, correctly: ' + e.message); }
  try { bulkUpdateRates(Object.assign({}, baseReq, { chargeType: 'NOPE', value: 1 })); }
  catch (e) { badRefused = true; }
  check('no charge type is refused', noneRefused);
  check('an unknown charge type is refused', badRefused);

  // ---- 4. set-to mode on one surcharge type ------------------------------
  const sur = types.filter(function (t) { return t.chargeType !== 'BASE'; })[0];
  const setValue = sur.valueType === 'PCT' ? 0.0777 : 0.77;
  const setReq = Object.assign({}, baseReq, { chargeType: sur.chargeType, value: setValue });

  Logger.log('');
  Logger.log('--- preview: set ' + sur.chargeType + ' to ' + setValue + ' ---');
  const pv = bulkUpdateRates(setReq);
  Logger.log('  mode                       : ' + pv.mode);
  Logger.log('  rows that would change     : ' + pv.count);
  Logger.log('  routes matched             : ' + pv.routesMatched +
             '  (no rate to supersede: ' + pv.skippedNoRate +
             ', out of scope: ' + pv.skippedNoScope + ')');
  pv.groups.forEach(function (g) {
    Logger.log('  ' + g.chargeType + ' (' + g.valueType + '): ' + g.count + ' rows');
    g.distinct.forEach(function (d) {
      Logger.log('      ' + d.from + ' -> ' + d.to + '   x ' + d.rows);
    });
  });
  check('a single charge type gives set-to mode', pv.mode === 'SET');
  check('only the chosen charge type is planned',
        pv.groups.length <= 1 &&
        (!pv.groups.length || pv.groups[0].chargeType === sur.chargeType));
  check('set-to moves every row to the same value',
        !pv.groups.length || pv.groups[0].distinct.every(function (d) { return d.to === setValue; }));
  check('the same selection gives the same planKey',
        bulkUpdateRates(setReq).planKey === pv.planKey);

  let staleRefused = false;
  try {
    bulkUpdateRates(Object.assign({}, setReq, { preview: false, planKey: 'deadbeef0000' }));
  } catch (e) { staleRefused = true; Logger.log('  stale key refused: ' + e.message); }
  check('a stale planKey is refused', staleRefused);

  if (sur.valueType === 'PCT') {
    let pctRefused = false;
    try { bulkUpdateRates(Object.assign({}, setReq, { value: 45 })); }
    catch (e) { pctRefused = true; }
    check('a percentage surcharge refuses 45 (i.e. 4500%)', pctRefused);
  }

  // ---- 5. uplift mode across every charge type ---------------------------
  Logger.log('');
  Logger.log('--- preview: +6% across ALL charge types ---');
  const allReq = Object.assign({}, baseReq, { chargeType: 'ALL', value: 6 });
  const pvAll = bulkUpdateRates(allReq);
  Logger.log('  mode                       : ' + pvAll.mode);
  Logger.log('  rows that would change     : ' + pvAll.count);
  pvAll.groups.forEach(function (g) {
    Logger.log('  ' + g.chargeType + ' (' + g.valueType + '): ' + g.count + ' rows');
    g.distinct.slice(0, 3).forEach(function (d) {
      Logger.log('      ' + d.from + ' -> ' + d.to + '   x ' + d.rows);
    });
  });
  check('ALL gives uplift mode', pvAll.mode === 'PCT');
  check('ALL spans more than one charge type', pvAll.groups.length > 1);
  check('ALL includes the base rate',
        pvAll.groups.some(function (g) { return g.chargeType === 'BASE'; }));
  check('every uplift is the entered percentage of its own current value',
        pvAll.groups.every(function (g) {
          return g.distinct.every(function (d) {
            return Math.abs(d.to - Math.round(d.from * 1.06 * 1000000) / 1000000) < 1e-9; });
        }));
  check('an uplift of -100% or worse is refused', (function () {
    try { bulkUpdateRates(Object.assign({}, allReq, { value: -100 })); return false; }
    catch (e) { return true; }
  })());

  // ---- 6. apply, verify, restore -----------------------------------------
  const CAP = 40;
  if (pvAll.count > CAP) {
    Logger.log('');
    Logger.log('SKIPPING THE WRITE — the ALL selection matches ' + pvAll.count + ' rows, more');
    Logger.log('than the ' + CAP + ' this diagnostic is willing to touch. Every preview check');
    Logger.log('above passed. Do the write half from the app instead.');
    return summariseBulkTest_(checks, null);
  }

  Logger.log('');
  Logger.log('--- snapshot, apply +6% to ALL, verify, restore ---');
  const tabs = [{ key: 'RATE_BASE', sheet: SHEET.RATE_BASE,
                  width: TABLES.RATE_BASE.headers.length, idCol: COL.RATE_BASE.Rate_ID },
                { key: 'RATE_SURCHARGE', sheet: SHEET.RATE_SURCHARGE,
                  width: TABLES.RATE_SURCHARGE.headers.length,
                  idCol: COL.RATE_SURCHARGE.Surcharge_Rate_ID }];
  tabs.forEach(function (tb) {
    invalidateSheetCache_(tb.sheet);
    tb.snapshot = getAllData_(tb.sheet).map(function (r) { return padRow_(r, tb.width); });
    tb.before = {};
    for (let i = 1; i < tb.snapshot.length; i++) {
      tb.before[safeInt(tb.snapshot[i][tb.idCol])] = tb.snapshot[i];
    }
    Logger.log('  snapshotted ' + tb.sheet + ': ' + (tb.snapshot.length - 1) + ' rows');
  });
  const auditBefore = Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0);

  const res = bulkUpdateRates(Object.assign({}, allReq,
                { preview: false, planKey: pvAll.planKey }));
  Logger.log('  batch reference            : ' + res.batchRef);
  Logger.log('  periods closed / created   : ' + res.rowsClosed + ' / ' + res.rowsCreated);
  res.groups.forEach(function (g) { Logger.log('    ' + g.chargeType + ': ' + g.written); });
  check('one new row per planned row', res.rowsCreated === pvAll.count);
  check('a batch reference was issued', /^BULK-\d{8}-\d{6}$/.test(String(res.batchRef)));
  check('both rate tables were written',
        res.groups.some(function (g) { return g.chargeType === 'BASE'; }) &&
        res.groups.some(function (g) { return g.chargeType !== 'BASE'; }));

  let fresh = 0, valuesOk = true, notesOk = true, closedOk = true;
  tabs.forEach(function (tb) {
    invalidateSheetCache_(tb.sheet);
    const C = COL[tb.key];
    const after = getAllData_(tb.sheet).map(function (r) { return padRow_(r, tb.width); });
    const valueCol = tb.key === 'RATE_BASE' ? COL.RATE_BASE.Base_Rate : COL.RATE_SURCHARGE.Value;
    tb.after = after;
    for (let i = 1; i < after.length; i++) {
      const id = safeInt(after[i][tb.idCol]);
      const was = tb.before[id];
      if (!was) {
        fresh++;
        if (fmtDate(after[i][C.Valid_From]) !== '2035-01-01') valuesOk = false;
        if (safeStr(after[i][C.Notes]).indexOf(res.batchRef) < 0) notesOk = false;
        if (safeNum(after[i][valueCol]) < 0) valuesOk = false;
        continue;
      }
      if (fmtDate(was[C.Valid_To]) === fmtDate(after[i][C.Valid_To])) continue;
      if (fmtDate(after[i][C.Valid_To]) !== '2034-12-31') closedOk = false;
    }
  });
  check('the new rows are on the sheets', fresh === pvAll.count, fresh + ' found');
  check('every new row starts on the date asked for and is non-negative', valuesOk);
  check('every new row names the batch in Notes', notesOk);
  check('closed periods end the day before the new one', closedOk);

  invalidateSheetCache_(SHEET.AUDIT_LOG);
  const auditRows = getAllData_(SHEET.AUDIT_LOG), A = COL.AUDIT_LOG;
  let stamped = 0;
  for (let i = 1; i < auditRows.length; i++) {
    if (safeStr(auditRows[i][A.Detail]).indexOf(res.batchRef) >= 0) stamped++;
  }
  Logger.log('  Audit_Log rows added       : ' +
             (Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0) - auditBefore) +
             ' (' + stamped + ' stamped with the batch reference)');
  check('every audit row is stamped with the batch', stamped >= res.rowsCreated);

  // ---- restore ------------------------------------------------------------
  tabs.forEach(function (tb) {
    const C = COL[tb.key];
    const restored = [];
    for (let i = 1; i < tb.after.length; i++) {
      const was = tb.before[safeInt(tb.after[i][tb.idCol])];
      if (was) { restored.push(was); continue; }
      const row = tb.after[i].slice();
      row[C.Active] = false;
      row[C.Notes]  = truncate_('TEST ROW from testBulkRateUpdate — deactivated, safe to delete');
      restored.push(row);
    }
    getSheet_(tb.sheet).getRange(2, 1, restored.length, tb.width).setValues(restored);
    invalidateSheetCache_(tb.sheet);
  });
  Logger.log('  restored both tables and deactivated ' + fresh + ' test row(s)');

  let identical = true;
  tabs.forEach(function (tb) {
    invalidateSheetCache_(tb.sheet);
    const final = getAllData_(tb.sheet).map(function (r) { return padRow_(r, tb.width); });
    for (let i = 1; i < tb.snapshot.length; i++) {
      for (let j = 0; j < tb.width; j++) {
        if (displayValue_(tb.snapshot[i][j]) !== displayValue_(final[i][j])) identical = false;
      }
    }
  });
  check('every pre-existing row is back exactly as it was', identical);

  return summariseBulkTest_(checks, res.batchRef);
}

function summariseBulkTest_(checks, batchRef) {
  const failed = checks.filter(function (c) { return !c.pass; });
  Logger.log('');
  Logger.log(failed.length ? 'PROBLEM — ' + failed.length + ' of ' + checks.length +
               ' checks failed: ' + failed.map(function (c) { return c.name; }).join('; ')
             : 'BULK RATE UPDATE WORKING  (' + checks.length + ' checks)');
  if (batchRef) {
    Logger.log('');
    Logger.log('Both tables were restored, so the only trace left is history: the Audit_Log');
    Logger.log('and *_Amends rows for batch ' + batchRef + ', plus the deactivated 2035');
    Logger.log('test rows. None of it can reach a forecast.');
  }
  return { ok: !failed.length, checks: checks.length, failed: failed.length,
           batchRef: batchRef };
}


// ─────────────────────────────────────────────────────────────────────────────
// ONE-OFF: ROLL HIGH LEVEL ID 1'S RM FUEL SCHEDULE ONTO MODELLING IDS 9-21
//
// A maintenance script, not a feature. bulkUpdateRates can set one value over
// one date range across a dimension selection; this needs FOUR consecutive
// periods copied onto an explicit list of routes, which the bulk screen has no
// shape for. Rather than widen that screen for a one-off, this does the one-off
// and says exactly what it did.
//
// It writes nothing directly. Every change goes through deleteSurchargeRate and
// saveSurchargeRate, the same two functions the Rates screen calls, so the
// overlap check, the scope check, Rate_Surcharge_Amends and the Audit_Log all
// behave exactly as they would for thirteen people making the edits by hand.
// The cost of that is speed — each call takes its own lock and re-reads the tab
// — and the benefit is that there is no second write path to keep correct.
//
// ── STATUS: NEVER COMMITTED. The job it was written for did not need doing. ───
//
// Previewed against the test copy on 2026-08-25 (see docs/TEST_PORTAL.md).
// runFuelScheduleCopy() has never been run, here or in production.
//
// The preview stopped at step 1 and that was the correct outcome. Two things it
// established that are worth not rediscovering:
//
//   1. High Level ID 1 has EIGHTEEN active ROYALMAIL routes, not one. A
//      Modelling ID is per delivery METHOD, so "the RM route" is RM24, RM48,
//      ROYALMAIL9AM, TRACKED24, SPECIALDELIVERY1PM and fourteen more. The
//      "expect exactly one active route" guard below therefore cannot pass for
//      this High Level ID, and will not pass for any other one either. Anyone
//      reusing this must first decide how the source route is chosen — most
//      obviously by adding Method_Code to the source key. Do not "fix" it by
//      relaxing the guard to take the first match; picking silently among
//      eighteen schedules is the failure this was built to prevent.
//
//   2. Modelling IDs 9-22 ALREADY carried the four periods this was meant to
//      copy onto 9-21, inserted as one batch (Rate_Surcharge row IDs 1650-1663).
//      Had the guard not fired, the run would have deactivated 52 correct rows
//      and rewritten them with identical values: thirteen routes of audit churn
//      and no change to any forecast.
//
// Left in place as tooling rather than deleted, because the same shape is what
// bringing Modelling IDs 5-8 into line would need — they are the outliers, on
// two other schedules. See the table in docs/TEST_PORTAL.md. If that never
// happens, this section is safe to delete whole; nothing references it.
// ─────────────────────────────────────────────────────────────────────────────

/** Where the schedule is copied FROM, and what it must look like before we trust it. */
const FUEL_COPY_SOURCE_HL_ID   = 1;
const FUEL_COPY_CODE           = 'FUEL';

/**
 * Accepted spellings of the Royal Mail carrier code, most likely first.
 *
 * A list rather than one literal because "the RM route" is what a person says
 * and `ROYALMAIL` is what Dim_Carrier seeds (setup.gs SEED_CARRIERS), and a
 * migrated sheet may carry either. Every entry is matched EXACTLY after
 * normKey — this is an allowlist of spellings, not a fuzzy match, so it cannot
 * quietly bind to some other carrier the way a substring rule could.
 *
 * If more than one of these turns up under the source High Level ID, that is
 * reported and the run stops rather than picking one.
 */
const FUEL_COPY_SOURCE_CARRIERS = ['ROYALMAIL', 'RM', 'ROYAL_MAIL', 'RMG', 'ROYAL MAIL'];

/** Inclusive range of Modelling IDs to copy ONTO. */
const FUEL_COPY_TARGET_FROM = 9;
const FUEL_COPY_TARGET_TO   = 21;

/**
 * The schedule as it was typed into the portal, restated here so the script can
 * refuse to run if the sheet no longer says this.
 *
 * This is a transcription of a manual edit, which is exactly the kind of input
 * worth not trusting: if someone corrects a date on the source route between
 * that edit and this run, a copy that read the source blindly would quietly
 * propagate a different schedule than the one that was signed off. So the source
 * is read, compared against this, and any disagreement stops the run and prints
 * both sides.
 *
 * Values are fractions because FUEL is seeded PCT in Dim_Surcharge (0.11 = 11%),
 * which is what saveSurchargeRate validates against.
 */
const FUEL_COPY_EXPECTED = [
  { from: '2026-01-01', to: '2026-05-03', value: 0.11 },
  { from: '2026-05-04', to: '2026-06-12', value: 0.16 },
  { from: '2026-06-13', to: '2026-09-30', value: 0.11 },
  { from: '2026-10-01', to: '2029-12-01', value: 0.14 }
];

const FUEL_COPY_SOURCE_REF = 'Bulk copy from HL1 RM';

/** Fractions compared with a tolerance, because 0.11 does not survive a round trip exactly. */
function fuelCopyValuesMatch_(a, b) { return Math.abs(safeNum(a) - safeNum(b)) < 1e-9; }

/** A route's active rows for one surcharge code, oldest period first. */
function fuelCopyRowsFor_(modellingId, code) {
  const data = getAllData_(SHEET.RATE_SURCHARGE), C = COL.RATE_SURCHARGE;
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][C.Modelling_ID]) !== safeInt(modellingId)) continue;
    if (normKey(data[i][C.Surcharge_Code]) !== normKey(code)) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({
      id:         safeInt(data[i][C.Surcharge_Rate_ID]),
      from:       fmtDate(data[i][C.Valid_From]),
      to:         fmtDate(data[i][C.Valid_To]),
      fromKey:    dateKey(data[i][C.Valid_From]),
      value:      safeNum(data[i][C.Value]),
      scenarioId: safeInt(data[i][C.Scenario_ID]) || 1
    });
  }
  out.sort(function (a, b) { return a.fromKey - b.fromKey; });
  return out;
}

/** "2026-01-01 to 2026-05-03    0.1100   scenario 1, row id 42" */
function fuelCopyDescribe_(r) {
  return r.from + ' to ' + r.to + '  ' + pad_(safeNum(r.value).toFixed(4), 8) +
         '   scenario ' + r.scenarioId + (r.id ? ', row id ' + r.id : '');
}

/**
 * Roll the FUEL schedule from High Level ID 1's RM route onto Modelling IDs 9-21.
 *
 * @param {boolean} commit  false (default) logs the plan and writes nothing
 */
function copyFuelScheduleToModellingIds(commit) {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log(commit ? '=== COPYING FUEL SCHEDULE (WRITING) ==='
                    : '=== FUEL SCHEDULE COPY — PREVIEW (nothing written) ===');

  prewarmSheetCache_([SHEET.RATE_SURCHARGE, SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS,
                      SHEET.DIM_SURCHARGE, SHEET.PERMISSIONS, SHEET.PORTAL_ROLES,
                      SHEET.SCOPE_MAPPING, SHEET.CONFIG]);

  const perms = requirePermissions_();
  requireEditRates_(perms);

  const def = surchargeDefinition_(FUEL_COPY_CODE);
  if (!def) throw new Error('Dim_Surcharge has no "' + FUEL_COPY_CODE + '" row, so this ' +
    'cannot run. Add the surcharge type first.');
  Logger.log('  surcharge   : ' + FUEL_COPY_CODE + '  (' + def.valueType + ', ' +
             def.proration + ')');
  Logger.log('  running as  : ' + perms.email + '  role ' + perms.role +
             (perms.allAccess ? '  (all access)' : ''));

  // ═══ 1. Find the source route, and refuse to guess ════════════════════════
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;

  /* Every route under the source High Level ID, whatever its carrier, plus a
     census of carrier codes across the whole tab. Gathered BEFORE deciding
     anything, because the useful thing to print when the expected carrier is
     absent is the list of carriers that are actually there — "check the carrier
     code in Modelling_IDs" is a failure message that makes the reader go and do
     by hand what the function already had the data to do for them. */
  const underHl = [], carrierCensus = {};
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id) continue;
    const carrier = normKey(md[i][M.Carrier_Code]);
    carrierCensus[carrier] = (carrierCensus[carrier] || 0) + 1;
    if (safeInt(md[i][M.High_Level_ID]) !== FUEL_COPY_SOURCE_HL_ID) continue;
    underHl.push({
      id: id, carrier: carrier,
      code:   safeStr(md[i][M.Modelling_Code]),
      method: safeStr(md[i][M.Method_Code]),
      lp:     safeStr(md[i][M.Letter_Parcel]),
      active: safeBool(md[i][M.Active])
    });
  }
  underHl.sort(function (a, b) { return a.id - b.id; });

  Logger.log('');
  Logger.log('--- 1. source: High Level ID ' + FUEL_COPY_SOURCE_HL_ID + ', carrier one of ' +
             FUEL_COPY_SOURCE_CARRIERS.join(' / ') + ' ---');

  /* Which accepted spellings are actually present. Exact matches only. */
  const spellingsPresent = FUEL_COPY_SOURCE_CARRIERS
    .map(function (c) { return normKey(c); })
    .filter(function (c, i, a) { return a.indexOf(c) === i; })
    .filter(function (c) {
      return underHl.some(function (r) { return r.carrier === c && r.active; });
    });

  if (!spellingsPresent.length) {
    Logger.log('  NOTHING FOUND. No ACTIVE Modelling ID under High Level ID ' +
               FUEL_COPY_SOURCE_HL_ID + ' has a Carrier_Code matching any of: ' +
               FUEL_COPY_SOURCE_CARRIERS.join(', '));
    Logger.log('');
    if (!underHl.length) {
      Logger.log('  In fact High Level ID ' + FUEL_COPY_SOURCE_HL_ID + ' has NO Modelling IDs');
      Logger.log('  at all. Check the High Level ID is the one you meant.');
    } else {
      Logger.log('  What High Level ID ' + FUEL_COPY_SOURCE_HL_ID + ' actually has:');
      underHl.forEach(function (r) {
        Logger.log('    Modelling ID ' + pad_(r.id, 4) + '  carrier ' + pad_(r.carrier, 12) +
                   '  method ' + pad_(r.method, 14) + '  ' + pad_(r.lp, 7) +
                   (r.active ? '' : '   [INACTIVE]'));
      });
    }
    Logger.log('');
    Logger.log('  Every Carrier_Code in Modelling_IDs, with how many routes use it:');
    Object.keys(carrierCensus).sort().forEach(function (c) {
      Logger.log('    ' + pad_(carrierCensus[c], 5) + '  ' + (c || '(blank)'));
    });
    Logger.log('');
    Logger.log('  If the right code is in the census but not in the accepted list, add it to');
    Logger.log('  FUEL_COPY_SOURCE_CARRIERS in rates.gs and run this again.');
    return { ok: false, reason: 'no source route',
             highLevelIdRoutes: underHl, carrierCensus: carrierCensus };
  }

  if (spellingsPresent.length > 1) {
    Logger.log('  STOPPING. High Level ID ' + FUEL_COPY_SOURCE_HL_ID + ' has active routes ' +
               'under MORE THAN ONE accepted spelling of the carrier: ' +
               spellingsPresent.join(', ') + '.');
    Logger.log('  These are probably the same carrier entered two ways, which is a data');
    Logger.log('  problem worth fixing before a schedule is copied off one of them.');
    underHl.forEach(function (r) {
      Logger.log('    Modelling ID ' + pad_(r.id, 4) + '  carrier ' + pad_(r.carrier, 12) +
                 '  method ' + pad_(r.method, 14) + (r.active ? '' : '   [INACTIVE]'));
    });
    return { ok: false, reason: 'carrier code spelled more than one way',
             spellings: spellingsPresent };
  }

  const sourceCarrier = spellingsPresent[0];
  Logger.log('  matched Carrier_Code: ' + sourceCarrier +
             (sourceCarrier === normKey(FUEL_COPY_SOURCE_CARRIERS[0]) ? ''
              : '   (an alternative spelling — ' +
                normKey(FUEL_COPY_SOURCE_CARRIERS[0]) + ' is the expected one)'));

  const candidates = underHl.filter(function (r) { return r.carrier === sourceCarrier; });

  /* Every candidate is listed with its schedule BEFORE any decision, because the
     thing worth knowing when there is more than one is not "which did it pick"
     but "do they even agree". */
  candidates.forEach(function (c) {
    const rows = fuelCopyRowsFor_(c.id, FUEL_COPY_CODE);
    Logger.log('  Modelling ID ' + c.id + '  ' + c.code + '  method ' + c.method +
               '  ' + c.lp + (c.active ? '' : '   [INACTIVE]'));
    if (!rows.length) Logger.log('      no active ' + FUEL_COPY_CODE + ' rows');
    rows.forEach(function (r) { Logger.log('      ' + fuelCopyDescribe_(r)); });
  });

  const active = candidates.filter(function (c) { return c.active; });
  if (active.length !== 1) {
    Logger.log('');
    Logger.log('  STOPPING. Expected exactly one ACTIVE ' + sourceCarrier +
               ' route under High Level ID ' +
               FUEL_COPY_SOURCE_HL_ID + ', found ' + active.length + '.');
    Logger.log('  Their schedules are printed above — compare them and tell me which one');
    Logger.log('  is the source, or deactivate the ones that are not.');
    return { ok: false, reason: 'source is ambiguous',
             candidates: candidates.map(function (c) { return c.id; }),
             activeCandidates: active.map(function (c) { return c.id; }) };
  }

  const sourceId = active[0].id;
  const sourceRows = fuelCopyRowsFor_(sourceId, FUEL_COPY_CODE);
  Logger.log('');
  Logger.log('  source is Modelling ID ' + sourceId + ' (the only active RM route).');

  // ═══ 2. Confirm the source matches the schedule that was signed off ═══════
  Logger.log('');
  Logger.log('--- 2. does the source match the four periods I was given? ---');
  const diffs = [];
  if (sourceRows.length !== FUEL_COPY_EXPECTED.length) {
    diffs.push('expected ' + FUEL_COPY_EXPECTED.length + ' active periods, found ' +
               sourceRows.length);
  }
  const n = Math.max(sourceRows.length, FUEL_COPY_EXPECTED.length);
  for (let i = 0; i < n; i++) {
    const want = FUEL_COPY_EXPECTED[i], got = sourceRows[i];
    if (!want) { diffs.push('period ' + (i + 1) + ': unexpected extra — ' +
                            fuelCopyDescribe_(got)); continue; }
    if (!got)  { diffs.push('period ' + (i + 1) + ': missing — expected ' + want.from +
                            ' to ' + want.to + '  ' + want.value.toFixed(4)); continue; }
    if (got.from !== want.from || got.to !== want.to ||
        !fuelCopyValuesMatch_(got.value, want.value)) {
      diffs.push('period ' + (i + 1) + ': expected ' + want.from + ' to ' + want.to +
                 '  ' + want.value.toFixed(4) + '   but found ' + got.from + ' to ' +
                 got.to + '  ' + got.value.toFixed(4));
    }
  }

  FUEL_COPY_EXPECTED.forEach(function (w, i) {
    Logger.log('  expected ' + (i + 1) + '  ' + w.from + ' to ' + w.to + '  ' +
               pad_(w.value.toFixed(4), 8));
  });
  if (diffs.length) {
    Logger.log('');
    Logger.log('  STOPPING. The source does not match what I was told to copy:');
    diffs.forEach(function (d) { Logger.log('    ' + d); });
    Logger.log('  Nothing has been written. Either the source was edited after those four');
    Logger.log('  periods were given to me, or the periods were transcribed wrongly — and');
    Logger.log('  I cannot tell which, so I am not guessing.');
    return { ok: false, reason: 'source does not match the expected schedule',
             sourceModellingId: sourceId, differences: diffs };
  }
  Logger.log('  MATCHES — all ' + FUEL_COPY_EXPECTED.length + ' periods agree.');

  const sourceScenarios = {};
  sourceRows.forEach(function (r) { sourceScenarios[r.scenarioId] = true; });
  Logger.log('  source Scenario_ID(s): ' + Object.keys(sourceScenarios).join(', ') +
             '   (targets keep their OWN scenario, not the source one)');

  // ═══ 3. Plan each target ═════════════════════════════════════════════════
  const midById = {};
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (id) midById[id] = { id: id, code: safeStr(md[i][M.Modelling_Code]),
                            hlId: safeInt(md[i][M.High_Level_ID]),
                            carrier: normKey(md[i][M.Carrier_Code]),
                            active: safeBool(md[i][M.Active]) };
  }

  Logger.log('');
  Logger.log('--- 3. targets: Modelling IDs ' + FUEL_COPY_TARGET_FROM + ' to ' +
             FUEL_COPY_TARGET_TO + ' ---');

  const plan = [], skipped = [];
  for (let mid = FUEL_COPY_TARGET_FROM; mid <= FUEL_COPY_TARGET_TO; mid++) {
    const meta = midById[mid];
    Logger.log('');
    Logger.log('  Modelling ID ' + mid + (meta ? '  ' + meta.code + '  HL' + meta.hlId +
               ' ' + meta.carrier : ''));

    const skip = function (why) {
      Logger.log('      SKIPPED — ' + why);
      skipped.push({ modellingId: mid, reason: why });
    };

    if (!meta)            { skip('no such Modelling ID in Modelling_IDs'); continue; }
    if (!meta.active)     { skip('the route is inactive'); continue; }
    if (mid === sourceId) { skip('this IS the source route — already correct, so it is ' +
                                'left alone rather than rewritten'); continue; }
    /* The non-throwing form, so one out-of-scope route does not end the batch.
       saveSurchargeRate re-checks this itself; this only decides whether to try. */
    if (!canSeeModellingId_(perms, mid, true)) {
      skip('outside what ' + perms.email + ' can edit'); continue;
    }

    const current = fuelCopyRowsFor_(mid, FUEL_COPY_CODE);
    if (!current.length) {
      Logger.log('      currently: no active ' + FUEL_COPY_CODE + ' rows');
    } else {
      current.forEach(function (r) { Logger.log('      currently: ' + fuelCopyDescribe_(r)); });
    }

    /* Scenario_ID is carried over, not assumed. A route whose FUEL rows sit on
       more than one scenario is skipped rather than resolved: writing four rows
       on one scenario while deactivating rows across two would collapse a
       distinction someone deliberately made, and there is nothing here that says
       which scenario was meant. */
    const scenarios = {};
    current.forEach(function (r) { scenarios[r.scenarioId] = true; });
    const scenarioKeys = Object.keys(scenarios);
    if (scenarioKeys.length > 1) {
      skip('its ' + FUEL_COPY_CODE + ' rows span Scenario_IDs ' + scenarioKeys.join(', ') +
           ' — tell me which one to write to and I will run it again');
      continue;
    }
    const scenarioId = scenarioKeys.length ? safeInt(scenarioKeys[0]) : 1;
    Logger.log('      scenario : ' + scenarioId +
               (scenarioKeys.length ? '  (carried over from the rows above)'
                                    : '  (no existing rows, so defaulting to 1)'));

    FUEL_COPY_EXPECTED.forEach(function (w, i) {
      Logger.log('      becomes  ' + (i + 1) + '  ' + w.from + ' to ' + w.to + '  ' +
                 pad_(w.value.toFixed(4), 8) + '   scenario ' + scenarioId);
    });

    plan.push({ modellingId: mid, scenarioId: scenarioId,
                deleteIds: current.map(function (r) { return r.id; }),
                current: current });
  }

  const rowsToDeactivate = plan.reduce(function (a, x) { return a + x.deleteIds.length; }, 0);

  Logger.log('');
  Logger.log('--- summary ---');
  Logger.log('  targets to change  : ' + plan.length);
  Logger.log('  rows to deactivate : ' + rowsToDeactivate);
  Logger.log('  rows to create     : ' + (plan.length * FUEL_COPY_EXPECTED.length));
  Logger.log('  skipped            : ' + skipped.length);
  skipped.forEach(function (s) {
    Logger.log('      ' + s.modellingId + ' — ' + s.reason);
  });

  if (!commit) {
    Logger.log('');
    Logger.log('  NOTHING WRITTEN. Run copyFuelScheduleToModellingIds(true) — or');
    Logger.log('  runFuelScheduleCopy() — to apply exactly the plan above.');
    return { ok: true, preview: true, sourceModellingId: sourceId,
             targets: plan.length, skipped: skipped,
             rowsToDeactivate: rowsToDeactivate,
             rowsToCreate: plan.length * FUEL_COPY_EXPECTED.length,
             plan: plan.map(function (x) {
               return { modellingId: x.modellingId, scenarioId: x.scenarioId,
                        replacing: x.deleteIds.length }; }) };
  }

  if (!plan.length) throw new Error('Nothing to do — every target was skipped. The ' +
    'reasons are in the log above.');

  // ═══ 4. Write, one target at a time ══════════════════════════════════════
  Logger.log('');
  Logger.log('--- 4. writing ---');
  const done = [], failed = [];

  plan.forEach(function (x) {
    const note = 'FUEL schedule copied from Modelling ID ' + sourceId +
                 ' (High Level ID ' + FUEL_COPY_SOURCE_HL_ID + ' ' +
                 sourceCarrier + ')';
    let deactivated = 0, created = 0;
    try {
      /* Re-read rather than trusting the plan's IDs. Every call below invalidates
         the cache, and this is a long batch with no lock held across it — so what
         is actually active for this route right now is the only safe thing to
         deactivate. A difference from the preview is worth saying out loud. */
      invalidateSheetCache_(SHEET.RATE_SURCHARGE);
      const nowRows = fuelCopyRowsFor_(x.modellingId, FUEL_COPY_CODE);
      const nowIds  = nowRows.map(function (r) { return r.id; }).sort().join(',');
      const planIds = x.deleteIds.slice().sort().join(',');
      if (nowIds !== planIds) {
        Logger.log('  Modelling ID ' + x.modellingId + ': rows changed since the preview ' +
                   '(was [' + planIds + '], now [' + nowIds + ']) — deactivating what is ' +
                   'actually there.');
      }

      nowRows.forEach(function (r) { deleteSurchargeRate(r.id); deactivated++; });

      FUEL_COPY_EXPECTED.forEach(function (w) {
        saveSurchargeRate({
          modellingId: x.modellingId,
          code:        FUEL_COPY_CODE,
          validFrom:   w.from,
          validTo:     w.to,
          value:       w.value,
          scenarioId:  x.scenarioId,
          sourceRef:   FUEL_COPY_SOURCE_REF,
          notes:       note
          /* closePrevious deliberately NOT set. The four periods are already
             exactly contiguous, so closing the preceding one would rewrite its
             Valid_To to the date it already holds — a no-change amend row per
             period, thirteen times over. assertNoOverlap_ still runs on every
             save, which is the check that actually matters here. */
        });
        created++;
      });

      Logger.log('  Modelling ID ' + pad_(x.modellingId, 3) + '  OK — ' + deactivated +
                 ' deactivated, ' + created + ' created, scenario ' + x.scenarioId);
      done.push({ modellingId: x.modellingId, deactivated: deactivated, created: created });

    } catch (e) {
      /* One target's failure must not end the batch, but a target that failed
         PART WAY is a different problem: its old rows may be off and its new ones
         incomplete, which is a route with a gap in its FUEL schedule. Said loudly,
         with the counts, because the fix depends on how far it got. */
      const partial = (deactivated > 0 || created > 0) && created < FUEL_COPY_EXPECTED.length;
      Logger.log('  Modelling ID ' + pad_(x.modellingId, 3) + '  ' +
                 (partial ? '** PART DONE **' : 'FAILED') + ' — ' + e.message);
      if (partial) {
        Logger.log('      ' + deactivated + ' row(s) deactivated and only ' + created +
                   ' of ' + FUEL_COPY_EXPECTED.length + ' created. This route now has an ' +
                   'INCOMPLETE FUEL schedule and needs fixing by hand or by re-running.');
      }
      failed.push({ modellingId: x.modellingId, error: e.message,
                    deactivated: deactivated, created: created, partial: partial });
    }
  });

  const partials = failed.filter(function (f) { return f.partial; });
  Logger.log('');
  Logger.log('--- done in ' + Math.round((Date.now() - t0) / 1000) + 's ---');
  Logger.log('  changed : ' + done.length + ' route(s)');
  Logger.log('  failed  : ' + failed.length + ' route(s)');
  Logger.log('  skipped : ' + skipped.length + ' route(s)');
  if (partials.length) {
    Logger.log('');
    Logger.log('  ** ' + partials.length + ' route(s) are PART DONE and have an incomplete');
    Logger.log('     FUEL schedule right now: ' +
               partials.map(function (f) { return f.modellingId; }).join(', '));
    Logger.log('     Re-run the preview to see where they stand before doing anything else.');
  }
  Logger.log('');
  Logger.log('  Every change went through saveSurchargeRate / deleteSurchargeRate, so it is');
  Logger.log('  all in Rate_Surcharge_Amends and the Audit_Log. Re-run the calculation and');
  Logger.log('  publish for the new schedule to reach the forecast.');

  return { ok: !failed.length, preview: false, sourceModellingId: sourceId,
           changed: done.length, failed: failed, skipped: skipped,
           partial: partials.map(function (f) { return f.modellingId; }),
           seconds: Math.round((Date.now() - t0) / 1000) };
}

/** Convenience wrappers, so both appear in the editor's function list. */
function previewFuelScheduleCopy() { return copyFuelScheduleToModellingIds(false); }
function runFuelScheduleCopy()     { return copyFuelScheduleToModellingIds(true); }


// ─────────────────────────────────────────────────────────────────────────────
// SPREAD A ROUTE'S FUEL SCHEDULE TO THE SAME ROUTE UNDER EVERY OTHER HL ID
//
// The sibling of copyFuelScheduleToModellingIds, turned ninety degrees. That one
// held the High Level ID fixed and walked a list of Modelling IDs. This holds
// the ROUTE fixed — carrier, method and letter/parcel — and walks every other
// High Level ID, because "I have updated Modelling IDs 5 and 6 for High Level ID
// 1, now do the rest" is a statement about a delivery method, not about a range
// of IDs.
//
// Two deliberate differences from the earlier one-off, both learned from it:
//
//   1. NO TRANSCRIBED SCHEDULE. copyFuelScheduleToModellingIds compared the
//      source against four periods typed into a chat message, which was right
//      for that job because a human had dictated them. Here nobody has, so
//      inventing an expectation would only create a second thing that can be
//      wrong. The source's live rows ARE the specification; the preview's job is
//      to put them on screen where a person can read them before committing.
//
//   2. METHODS ARE DERIVED, NOT NAMED. The input is the two Modelling IDs that
//      were actually edited. Their carrier, method and letter/parcel are read
//      off them. Hardcoding 'FIRSTCLASSNOSIG' would be the same mistake as
//      hardcoding 'RM' when the data says 'ROYALMAIL'.
//
// FUEL only. Base rates and every other surcharge code are untouched.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * The routes whose FUEL schedule is authoritative. Everything else is derived.
 *
 * Modelling IDs rather than a carrier/method pair because these are the rows a
 * person edited and can point at. The run reports what each one turns out to be
 * and refuses if they are not all under one High Level ID.
 */
const FUEL_SPREAD_SOURCE_MIDS = [5, 6];

const FUEL_SPREAD_SOURCE_REF = 'Bulk copy from HL1 by method';

/** Modelling_IDs keyed by ID, with the four fields that identify a route. */
function fuelSpreadRoutes_() {
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  const byId = {}, all = [];
  for (let i = 1; i < md.length; i++) {
    const id = safeInt(md[i][M.Modelling_ID]);
    if (!id) continue;
    const r = {
      id: id,
      hlId:    safeInt(md[i][M.High_Level_ID]),
      carrier: normKey(md[i][M.Carrier_Code]),
      method:  normKey(md[i][M.Method_Code]),
      lp:      normKey(md[i][M.Letter_Parcel]),
      code:    safeStr(md[i][M.Modelling_Code]),
      active:  safeBool(md[i][M.Active])
    };
    byId[id] = r;
    all.push(r);
  }
  return { byId: byId, all: all };
}

/** "ROYALMAIL / FIRSTCLASSNOSIG / PARCEL" — the identity a target must share. */
function fuelSpreadRouteKey_(r) {
  return r.carrier + ' / ' + r.method + ' / ' + r.lp;
}

/**
 * Copy each source route's FUEL schedule onto the same route under every other
 * High Level ID.
 *
 * @param {boolean} commit  false (default) logs the plan and writes nothing
 */
function copyFuelScheduleToOtherHighLevelIds(commit) {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log(commit ? '=== SPREADING FUEL SCHEDULE BY METHOD (WRITING) ==='
                    : '=== FUEL SCHEDULE SPREAD — PREVIEW (nothing written) ===');

  prewarmSheetCache_([SHEET.RATE_SURCHARGE, SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS,
                      SHEET.DIM_SURCHARGE, SHEET.PERMISSIONS, SHEET.PORTAL_ROLES,
                      SHEET.SCOPE_MAPPING, SHEET.CONFIG]);

  const perms = requirePermissions_();
  requireEditRates_(perms);

  const def = surchargeDefinition_(FUEL_COPY_CODE);
  if (!def) throw new Error('Dim_Surcharge has no "' + FUEL_COPY_CODE + '" row.');
  Logger.log('  surcharge   : ' + FUEL_COPY_CODE + '  (' + def.valueType + ', ' +
             def.proration + ')   — nothing else is touched');
  Logger.log('  running as  : ' + perms.email + '  role ' + perms.role +
             (perms.allAccess ? '  (all access)' : ''));

  const routes = fuelSpreadRoutes_();

  // ═══ 1. Resolve the source routes ════════════════════════════════════════
  Logger.log('');
  Logger.log('--- 1. sources: Modelling IDs ' + FUEL_SPREAD_SOURCE_MIDS.join(' and ') + ' ---');

  const sources = [], sourceProblems = [];
  FUEL_SPREAD_SOURCE_MIDS.forEach(function (id) {
    const r = routes.byId[id];
    if (!r) { sourceProblems.push('Modelling ID ' + id + ' does not exist'); return; }
    if (!r.active) { sourceProblems.push('Modelling ID ' + id + ' is inactive'); return; }

    const rows = fuelCopyRowsFor_(id, FUEL_COPY_CODE);
    Logger.log('  Modelling ID ' + r.id + '  HL' + r.hlId + '  ' + fuelSpreadRouteKey_(r));
    Logger.log('      ' + r.code);
    if (!rows.length) {
      Logger.log('      NO ACTIVE ' + FUEL_COPY_CODE + ' ROWS');
      /* Refuse rather than treat it as "copy nothing". Spreading an empty
         schedule means deactivating every target's FUEL rows and creating none,
         which silently prices those routes at zero fuel — the RATE_MISSING
         failure the engine has a check for. Never do that by implication. */
      sourceProblems.push('Modelling ID ' + id + ' has no active ' + FUEL_COPY_CODE +
        ' rows, so there is nothing to copy. Spreading it would delete every ' +
        'target\'s fuel schedule and replace it with nothing.');
      return;
    }
    rows.forEach(function (x) { Logger.log('      ' + fuelCopyDescribe_(x)); });
    sources.push({ route: r, rows: rows, key: fuelSpreadRouteKey_(r) });
  });

  if (sourceProblems.length) {
    Logger.log('');
    Logger.log('  STOPPING. The sources are not usable as given:');
    sourceProblems.forEach(function (p) { Logger.log('    ' + p); });
    return { ok: false, reason: 'source not usable', problems: sourceProblems };
  }

  /* All sources must sit under one High Level ID. "I have updated them for high
     level ID 1" is only meaningful if they are in fact all under one, and if two
     came from different segments the phrase describes no single intention. */
  const sourceHls = {};
  sources.forEach(function (s) { sourceHls[s.route.hlId] = true; });
  const sourceHlIds = Object.keys(sourceHls).map(function (k) { return safeInt(k); });
  if (sourceHlIds.length !== 1) {
    Logger.log('');
    Logger.log('  STOPPING. The source Modelling IDs are spread across High Level IDs ' +
               sourceHlIds.join(', ') + '. Expected all of them under one.');
    return { ok: false, reason: 'sources span several High Level IDs',
             highLevelIds: sourceHlIds };
  }
  const sourceHl = sourceHlIds[0];
  Logger.log('');
  Logger.log('  all sources are under High Level ID ' + sourceHl +
             ', which is therefore excluded from the targets.');

  /* Two sources resolving to the same route identity would make the second
     silently overwrite the first's plan. Cannot happen while a Modelling ID is
     unique per HL/carrier/method/letter-parcel, which is the point of asserting
     it rather than assuming it. */
  const seenKeys = {};
  for (let i = 0; i < sources.length; i++) {
    if (seenKeys[sources[i].key]) {
      Logger.log('');
      Logger.log('  STOPPING. Modelling IDs ' + FUEL_SPREAD_SOURCE_MIDS.join(' and ') +
                 ' describe the same route (' + sources[i].key + '), so one would ' +
                 'overwrite the other.');
      return { ok: false, reason: 'sources are the same route', key: sources[i].key };
    }
    seenKeys[sources[i].key] = true;
  }

  /* Whether the sources agree with each other is worth stating, because a person
     who edited "them" as a pair probably expects them to match — and if they do
     not, each still spreads its own, which is correct but worth seeing. */
  if (sources.length > 1) {
    const shape = function (s) {
      return s.rows.map(function (x) {
        return x.from + '>' + x.to + '@' + safeNum(x.value).toFixed(4); }).join(' | ');
    };
    const first = shape(sources[0]);
    const agree = sources.every(function (s) { return shape(s) === first; });
    Logger.log('  the ' + sources.length + ' source schedules ' +
               (agree ? 'are IDENTICAL to each other.'
                      : 'DIFFER from each other — each spreads its own, not a merge.'));
  }

  // ═══ 2. Find the targets ═════════════════════════════════════════════════
  Logger.log('');
  Logger.log('--- 2. targets: the same route under every other High Level ID ---');

  const plan = [], skipped = [], nearMisses = [];
  let totalDeletes = 0;

  sources.forEach(function (s) {
    Logger.log('');
    Logger.log('  === ' + s.key + '   (from Modelling ID ' + s.route.id + ') ===');

    /* Matched on carrier AND method AND letter/parcel — the whole route identity.
       Routes sharing carrier and method but differing on letter/parcel are
       reported separately rather than swept in: a LETTER route is a different
       thing from a PARCEL one and the engine branches on it. Listing them means
       nothing is missed silently in either direction. */
    const targets = routes.all.filter(function (r) {
      return r.hlId !== sourceHl && r.carrier === s.route.carrier &&
             r.method === s.route.method && r.lp === s.route.lp;
    }).sort(function (a, b) { return a.hlId - b.hlId || a.id - b.id; });

    routes.all.forEach(function (r) {
      if (r.hlId === sourceHl) return;
      if (r.carrier !== s.route.carrier || r.method !== s.route.method) return;
      if (r.lp === s.route.lp) return;
      nearMisses.push({ id: r.id, hlId: r.hlId, key: fuelSpreadRouteKey_(r),
                        wanted: s.key });
    });

    if (!targets.length) {
      Logger.log('    No other High Level ID has this route at all.');
      return;
    }

    targets.forEach(function (r) {
      const label = '    Modelling ID ' + pad_(r.id, 4) + '  HL' + pad_(r.hlId, 3);

      if (!r.active) {
        Logger.log(label + '  SKIPPED — the route is inactive');
        skipped.push({ modellingId: r.id, hlId: r.hlId, reason: 'route is inactive' });
        return;
      }
      if (!canSeeModellingId_(perms, r.id, true)) {
        Logger.log(label + '  SKIPPED — outside what ' + perms.email + ' can edit');
        skipped.push({ modellingId: r.id, hlId: r.hlId, reason: 'outside editable scope' });
        return;
      }

      const current = fuelCopyRowsFor_(r.id, FUEL_COPY_CODE);

      const scenarios = {};
      current.forEach(function (x) { scenarios[x.scenarioId] = true; });
      const scenarioKeys = Object.keys(scenarios);
      if (scenarioKeys.length > 1) {
        Logger.log(label + '  SKIPPED — its ' + FUEL_COPY_CODE + ' rows span Scenario_IDs ' +
                   scenarioKeys.join(', '));
        skipped.push({ modellingId: r.id, hlId: r.hlId,
                       reason: 'FUEL rows span Scenario_IDs ' + scenarioKeys.join(', ') });
        return;
      }
      const scenarioId = scenarioKeys.length ? safeInt(scenarioKeys[0]) : 1;

      /* Already correct is worth detecting, not just tolerating. The previous
         one-off would have rewritten 52 identical rows had a different guard not
         stopped it; skipping a match here costs one comparison and saves the
         audit trail from changes that change nothing. */
      const same = current.length === s.rows.length && current.every(function (c, i) {
        return c.from === s.rows[i].from && c.to === s.rows[i].to &&
               fuelCopyValuesMatch_(c.value, s.rows[i].value);
      });
      if (same) {
        Logger.log(label + '  already matches — left alone');
        skipped.push({ modellingId: r.id, hlId: r.hlId,
                       reason: 'already has this exact schedule' });
        return;
      }

      Logger.log(label + '  ' + current.length + ' row(s) -> ' + s.rows.length +
                 ', scenario ' + scenarioId +
                 (scenarioKeys.length ? '' : ' (defaulted, no existing rows)'));
      current.forEach(function (x) { Logger.log('           was: ' + fuelCopyDescribe_(x)); });
      s.rows.forEach(function (x, i) {
        Logger.log('           now: ' + (i + 1) + '  ' + x.from + ' to ' + x.to + '  ' +
                   pad_(safeNum(x.value).toFixed(4), 8) + '   scenario ' + scenarioId);
      });

      totalDeletes += current.length;
      plan.push({ modellingId: r.id, hlId: r.hlId, scenarioId: scenarioId,
                  deleteIds: current.map(function (x) { return x.id; }),
                  rows: s.rows, sourceId: s.route.id, key: s.key });
    });
  });

  // ═══ 3. Summary ══════════════════════════════════════════════════════════
  const totalCreates = plan.reduce(function (a, x) { return a + x.rows.length; }, 0);
  const calls = totalDeletes + totalCreates;

  Logger.log('');
  Logger.log('--- summary ---');
  Logger.log('  routes to change   : ' + plan.length);
  Logger.log('  rows to deactivate : ' + totalDeletes);
  Logger.log('  rows to create     : ' + totalCreates);
  Logger.log('  skipped            : ' + skipped.length);

  const bySkip = {};
  skipped.forEach(function (s) { bySkip[s.reason] = (bySkip[s.reason] || 0) + 1; });
  Object.keys(bySkip).sort().forEach(function (k) {
    Logger.log('      ' + pad_(bySkip[k], 4) + '  ' + k);
  });

  if (nearMisses.length) {
    Logger.log('');
    Logger.log('--- ' + nearMisses.length + ' route(s) share the carrier and method but NOT ' +
               'the letter/parcel side, so they are NOT targeted ---');
    nearMisses.forEach(function (n) {
      Logger.log('    Modelling ID ' + pad_(n.id, 4) + '  HL' + pad_(n.hlId, 3) + '  ' +
                 n.key + '   (wanted ' + n.wanted + ')');
    });
    Logger.log('  Say so if these should be included and I will widen the match.');
  }

  /* Every write goes through saveSurchargeRate / deleteSurchargeRate, one lock
     and one tab re-read each. On the test project, whose GCP project has the
     Sheets API disabled, reads fall through to one round-trip per tab and run
     roughly six times slower than production (FINDINGS.md P14) — so the 6-minute
     ceiling is a real constraint here in a way it is not in production. Said in
     the preview, where it can still change the decision. */
  Logger.log('');
  Logger.log('  ' + calls + ' write call(s) if committed, each taking the script lock and');
  Logger.log('  re-reading Rate_Surcharge. Against the 6-minute execution ceiling that is');
  Logger.log('  ' + (calls > 120 ? 'A REAL RISK — consider narrowing the sources and running twice.'
                                 : calls > 60 ? 'worth watching.' : 'comfortable.'));
  if (calls > 60) {
    Logger.log('  Enabling the Sheets API in the test project\'s GCP project (736650435578)');
    Logger.log('  removes most of the risk — see docs/TEST_PORTAL.md.');
  }

  if (!commit) {
    Logger.log('');
    Logger.log('  NOTHING WRITTEN. Read the "was/now" lines above, then run');
    Logger.log('  runFuelScheduleSpread() to apply exactly this plan.');
    return { ok: true, preview: true, sourceHighLevelId: sourceHl,
             sources: sources.map(function (s) {
               return { modellingId: s.route.id, key: s.key, periods: s.rows.length }; }),
             routesToChange: plan.length, rowsToDeactivate: totalDeletes,
             rowsToCreate: totalCreates, writeCalls: calls,
             skipped: skipped, nearMisses: nearMisses,
             plan: plan.map(function (x) {
               return { modellingId: x.modellingId, hlId: x.hlId,
                        scenarioId: x.scenarioId, replacing: x.deleteIds.length }; }) };
  }

  if (!plan.length) throw new Error('Nothing to do — every target was skipped or already ' +
    'matches. The reasons are in the log above.');

  // ═══ 4. Write ════════════════════════════════════════════════════════════
  Logger.log('');
  Logger.log('--- 4. writing ---');
  const done = [], failed = [];

  plan.forEach(function (x) {
    const note = FUEL_COPY_CODE + ' schedule copied from Modelling ID ' + x.sourceId +
                 ' (High Level ID ' + sourceHl + ' ' + x.key + ')';
    let deactivated = 0, created = 0;
    try {
      invalidateSheetCache_(SHEET.RATE_SURCHARGE);
      const nowRows = fuelCopyRowsFor_(x.modellingId, FUEL_COPY_CODE);
      const nowIds = nowRows.map(function (r) { return r.id; }).sort().join(',');
      const planIds = x.deleteIds.slice().sort().join(',');
      if (nowIds !== planIds) {
        Logger.log('  Modelling ID ' + x.modellingId + ': rows changed since the preview ' +
                   '(was [' + planIds + '], now [' + nowIds + ']) — deactivating what is ' +
                   'actually there.');
      }

      nowRows.forEach(function (r) { deleteSurchargeRate(r.id); deactivated++; });

      x.rows.forEach(function (w) {
        saveSurchargeRate({
          modellingId: x.modellingId,
          code:        FUEL_COPY_CODE,
          validFrom:   w.from,
          validTo:     w.to,
          value:       w.value,
          scenarioId:  x.scenarioId,
          sourceRef:   FUEL_SPREAD_SOURCE_REF,
          notes:       note
          /* closePrevious deliberately unset — the source's periods are already
             contiguous, so closing would rewrite a Valid_To to the date it holds.
             assertNoOverlap_ still runs on every save. */
        });
        created++;
      });

      Logger.log('  Modelling ID ' + pad_(x.modellingId, 4) + '  HL' + pad_(x.hlId, 3) +
                 '  OK — ' + deactivated + ' deactivated, ' + created + ' created');
      done.push({ modellingId: x.modellingId, hlId: x.hlId,
                  deactivated: deactivated, created: created });

    } catch (e) {
      const partial = (deactivated > 0 || created > 0) && created < x.rows.length;
      Logger.log('  Modelling ID ' + pad_(x.modellingId, 4) + '  HL' + pad_(x.hlId, 3) +
                 '  ' + (partial ? '** PART DONE **' : 'FAILED') + ' — ' + e.message);
      if (partial) {
        Logger.log('      ' + deactivated + ' deactivated and only ' + created + ' of ' +
                   x.rows.length + ' created. This route has an INCOMPLETE fuel schedule.');
      }
      failed.push({ modellingId: x.modellingId, hlId: x.hlId, error: e.message,
                    deactivated: deactivated, created: created, partial: partial });
    }
  });

  const partials = failed.filter(function (f) { return f.partial; });
  Logger.log('');
  Logger.log('--- done in ' + Math.round((Date.now() - t0) / 1000) + 's ---');
  Logger.log('  changed : ' + done.length + ' route(s)');
  Logger.log('  failed  : ' + failed.length + ' route(s)');
  Logger.log('  skipped : ' + skipped.length + ' route(s)');
  if (partials.length) {
    Logger.log('');
    Logger.log('  ** ' + partials.length + ' route(s) are PART DONE and have an incomplete');
    Logger.log('     fuel schedule now: ' +
               partials.map(function (f) { return f.modellingId; }).join(', '));
  }
  Logger.log('');
  Logger.log('  All of it went through saveSurchargeRate / deleteSurchargeRate, so it is in');
  Logger.log('  Rate_Surcharge_Amends and the Audit_Log. Re-run the calculation and publish');
  Logger.log('  for these to reach the forecast.');

  return { ok: !failed.length, preview: false, sourceHighLevelId: sourceHl,
           changed: done.length, failed: failed, skipped: skipped,
           partial: partials.map(function (f) { return f.modellingId; }),
           seconds: Math.round((Date.now() - t0) / 1000) };
}

/** Convenience wrappers, so both appear in the editor's function list. */
function previewFuelScheduleSpread() { return copyFuelScheduleToOtherHighLevelIds(false); }
function runFuelScheduleSpread()     { return copyFuelScheduleToOtherHighLevelIds(true); }


// ─────────────────────────────────────────────────────────────────────────────
// COMPARING RATE DATA BETWEEN TWO SPREADSHEETS
//
// clasp push moves CODE. It has never moved DATA, and the two get conflated at
// exactly the wrong moment: after a change is proven in the test portal, when the
// obvious next sentence is "now make production the same".
//
// Making them the same wholesale is the dangerous reading. The test copy is a
// SNAPSHOT — everything that happened in production since it was taken lives only
// in production. Overwriting production with the copy would delete that silently:
// somebody else's rate change, an actuals load, a published forecast. And the rows
// would arrive with no *_Amends history behind them, so the audit trail would show
// values nobody ever saved.
//
// So this reports and never writes. It answers the three questions that decide
// what is safe to do next:
//
//   ONLY IN OTHER   — changes made there that production has not got. What you
//                     probably want moved.
//   ONLY IN THIS    — rows production has and the other does not. What a wholesale
//                     copy would DESTROY. If this list is not empty, do not copy
//                     wholesale.
//   DIFFERENT       — same period, different value. Needs a decision per row,
//                     because either side could be the newer truth.
//
// Read-only, both sides. Run it from either project.
// ─────────────────────────────────────────────────────────────────────────────

/** Rate tables only. Mixes and actuals are a separate question, deliberately. */
function rateDiffTables_() {
  return [
    { key: 'RATE_BASE', label: 'Rate_Base',
      idCol: 'Rate_ID', valueCol: 'Base_Rate', codeCol: null },
    { key: 'RATE_SURCHARGE', label: 'Rate_Surcharge',
      idCol: 'Surcharge_Rate_ID', valueCol: 'Value', codeCol: 'Surcharge_Code' }
  ];
}

/**
 * One rate row reduced to what identifies it and what it says.
 *
 * Keyed on route + charge code + start date + scenario, NOT on the row ID. Row IDs
 * are allocated per spreadsheet, so the same rate created independently on both
 * sides has different IDs, and a copy taken at one moment has the same IDs for
 * rows that have since diverged. The business key is the only thing that means the
 * same on both sides.
 */
function rateDiffRows_(ss, t) {
  const sh = ss.getSheetByName(TABLES[t.key].sheet);
  if (!sh) return null;                       // tab missing — reported, not guessed at
  const data = sh.getDataRange().getValues();
  const C = COL[t.key], out = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const mid = safeInt(r[C.Modelling_ID]);
    if (!mid) continue;
    if (!safeBool(r[C.Active])) continue;      // inactive rows are history, not state
    const code = t.codeCol ? normKey(r[C[t.codeCol]]) : '';
    const key = mid + '|' + code + '|' + dateKey(r[C.Valid_From]) + '|' +
                (safeInt(r[C.Scenario_ID]) || 1);
    out[key] = {
      key: key, id: safeInt(r[C[t.idCol]]), mid: mid, code: code,
      from: fmtDate(r[C.Valid_From]), to: fmtDate(r[C.Valid_To]),
      value: safeNum(r[C[t.valueCol]]),
      currency: safeStr(r[C.Currency]),
      scenarioId: safeInt(r[C.Scenario_ID]) || 1,
      sourceRef: safeStr(r[C.Source_Ref])
    };
  }
  return out;
}

/** Same period, same money? Tolerance because a fraction does not round-trip exactly. */
function rateDiffSame_(a, b) {
  return a.to === b.to && Math.abs(a.value - b.value) < 1e-9 &&
         normKey(a.currency) === normKey(b.currency);
}

function rateDiffLine_(r) {
  return 'MID ' + pad_(r.mid, 4) + (r.code ? '  ' + pad_(r.code, 10) : '') +
         '  ' + r.from + ' to ' + r.to +
         '  ' + pad_(r.value.toFixed(4), 10) +
         (r.currency ? ' ' + r.currency : '') + '  scen ' + r.scenarioId;
}

/**
 * Compare this project's rate data against another spreadsheet's.
 *
 * @param {string=} otherId  defaults to TEST_SPREADSHEET_ID
 */
function compareRatesWithOtherSpreadsheet(otherId) {
  requireMaintenance_();
  Logger.log('=== RATE DATA COMPARISON (read-only, nothing is written) ===');

  const thisId = spreadsheetId_();
  const wanted = safeStr(otherId) || safeStr(TEST_SPREADSHEET_ID);

  if (!wanted || wanted === 'PASTE_THE_TEST_SPREADSHEET_ID_HERE') {
    throw new Error('No spreadsheet to compare against. Pass an ID, or set ' +
      'TEST_SPREADSHEET_ID in utils.gs.');
  }
  if (wanted === thisId) {
    throw new Error('That is the spreadsheet this project already uses (' + thisId +
      '), so there is nothing to compare. Pass the OTHER one.');
  }

  let thisSs, otherSs;
  try { thisSs = SpreadsheetApp.openById(thisId); }
  catch (e) { throw new Error('Cannot open this project\'s own spreadsheet ' + thisId +
    ': ' + e.message); }
  try { otherSs = SpreadsheetApp.openById(wanted); }
  catch (e) { throw new Error('Cannot open ' + wanted + ': ' + e.message +
    '  Check this account has access to it.'); }

  Logger.log('  THIS  (what a move would change) : ' + thisId);
  Logger.log('                                     ' + thisSs.getName());
  Logger.log('  OTHER (where the edits were made): ' + wanted);
  Logger.log('                                     ' + otherSs.getName());

  const summary = [];

  rateDiffTables_().forEach(function (t) {
    Logger.log('');
    Logger.log('════ ' + t.label + ' ════');

    const mine = rateDiffRows_(thisSs, t), theirs = rateDiffRows_(otherSs, t);
    if (!mine || !theirs) {
      Logger.log('  Tab missing on ' + (!mine ? 'THIS' : 'OTHER') + ' — skipped.');
      summary.push({ table: t.label, error: 'tab missing' });
      return;
    }

    const onlyOther = [], onlyThis = [], differ = [];
    Object.keys(theirs).forEach(function (k) {
      if (!mine[k]) onlyOther.push(theirs[k]);
      else if (!rateDiffSame_(mine[k], theirs[k])) differ.push({ mine: mine[k], theirs: theirs[k] });
    });
    Object.keys(mine).forEach(function (k) { if (!theirs[k]) onlyThis.push(mine[k]); });

    onlyOther.sort(function (a, b) { return a.mid - b.mid || (a.from < b.from ? -1 : 1); });
    onlyThis.sort(function (a, b) { return a.mid - b.mid || (a.from < b.from ? -1 : 1); });
    differ.sort(function (a, b) { return a.theirs.mid - b.theirs.mid; });

    Logger.log('  active rows: THIS ' + Object.keys(mine).length +
               ',  OTHER ' + Object.keys(theirs).length);

    Logger.log('');
    Logger.log('  --- ONLY IN OTHER: ' + onlyOther.length + ' — the edits not yet here ---');
    onlyOther.slice(0, 60).forEach(function (r) { Logger.log('    + ' + rateDiffLine_(r)); });
    if (onlyOther.length > 60) Logger.log('    ... and ' + (onlyOther.length - 60) + ' more');

    Logger.log('');
    Logger.log('  --- DIFFERENT: ' + differ.length + ' — same period, different value ---');
    differ.slice(0, 60).forEach(function (d) {
      Logger.log('    ~ ' + rateDiffLine_(d.theirs));
      Logger.log('      here: ' + d.mine.to + '  ' + d.mine.value.toFixed(4) +
                 (d.mine.currency ? ' ' + d.mine.currency : ''));
    });
    if (differ.length > 60) Logger.log('    ... and ' + (differ.length - 60) + ' more');

    Logger.log('');
    Logger.log('  --- ONLY IN THIS: ' + onlyThis.length +
               ' — a wholesale copy would DESTROY these ---');
    onlyThis.slice(0, 60).forEach(function (r) { Logger.log('    - ' + rateDiffLine_(r)); });
    if (onlyThis.length > 60) Logger.log('    ... and ' + (onlyThis.length - 60) + ' more');

    summary.push({ table: t.label, thisRows: Object.keys(mine).length,
                   otherRows: Object.keys(theirs).length,
                   onlyInOther: onlyOther.length, different: differ.length,
                   onlyInThis: onlyThis.length,
                   onlyInOtherRows: onlyOther.slice(0, 200),
                   differentRows: differ.slice(0, 200) });
  });

  // ---- everything else, by row count only, so other drift is visible ------
  Logger.log('');
  Logger.log('════ every other tab, row counts only ════');
  Logger.log('  ' + pad_('THIS', 8) + pad_('OTHER', 8) + '  tab');
  const drift = [];
  Object.keys(TABLES).forEach(function (k) {
    const name = TABLES[k].sheet;
    if (name === 'Rate_Base' || name === 'Rate_Surcharge') return;
    const a = thisSs.getSheetByName(name), b = otherSs.getSheetByName(name);
    const an = a ? Math.max(a.getLastRow() - 1, 0) : -1;
    const bn = b ? Math.max(b.getLastRow() - 1, 0) : -1;
    if (an !== bn) {
      drift.push({ tab: name, thisRows: an, otherRows: bn });
      Logger.log('  ' + pad_(an < 0 ? 'none' : an, 8) + pad_(bn < 0 ? 'none' : bn, 8) +
                 '  ' + name + '   <- differs');
    }
  });
  if (!drift.length) Logger.log('  Every other tab has the same number of rows.');

  const totalOnlyThis = summary.reduce(function (a, s) {
    return a + (s.onlyInThis || 0); }, 0);

  Logger.log('');
  Logger.log('--- what this means ---');
  if (totalOnlyThis) {
    Logger.log('  ' + totalOnlyThis + ' active rate row(s) exist HERE and not in the other');
    Logger.log('  spreadsheet. A wholesale copy would delete them. Move the ONLY IN OTHER');
    Logger.log('  and DIFFERENT rows deliberately instead — through saveBaseRate /');
    Logger.log('  saveSurchargeRate, so each one gets an Amends row saying who changed it.');
  } else {
    Logger.log('  Nothing exists here that is absent there, so no rate row would be lost.');
    Logger.log('  That does not make a wholesale copy right — rows written directly still');
    Logger.log('  arrive with no history behind them — but it removes the worst risk.');
  }
  Logger.log('  Nothing was written. This function has no write path.');

  return { ok: true, thisSpreadsheet: { id: thisId, name: thisSs.getName() },
           otherSpreadsheet: { id: wanted, name: otherSs.getName() },
           tables: summary, otherTabDrift: drift };
}

/** Compare against the test copy named in utils.gs. */
function compareRatesWithTestCopy() { return compareRatesWithOtherSpreadsheet(null); }


// ─────────────────────────────────────────────────────────────────────────────
// MOVING RATE DATA FROM ANOTHER SPREADSHEET INTO THIS ONE
//
// The companion to compareRatesWithOtherSpreadsheet: it does the move that
// comparison describes, and it does it through the same two functions the Rates
// screen calls, so every row lands with a Rate_Base_Amends / Rate_Surcharge_Amends
// entry and an Audit_Log line. A rate that appears with no history behind it is a
// rate nobody can later explain.
//
// THE UNIT OF TRANSFER IS A ROUTE AND A CHARGE, NOT A ROW. For each
// (Modelling_ID, charge code) whose active periods differ between the two sides,
// this side's periods are deactivated and the other side's are written in their
// place. Copying period-by-period cannot work: assertNoOverlap_ refuses a new
// period that straddles an existing one, and a schedule that was three periods
// here and four there has no row-to-row correspondence to patch. Replacing the
// whole schedule for that route is the only operation that is well defined.
//
// WHAT IT WILL NOT DO. A route-and-charge that exists HERE and not in the source
// is left completely alone, and reported. That is the asymmetry that keeps the
// move safe: the source is a snapshot, so its silence about a route means "this
// snapshot predates it", not "delete it". Making the two truly identical would
// mean honouring that silence, and there is no way to tell a rate added here last
// week from one the snapshot never had.
//
// It is IDEMPOTENT: a group whose periods already match is skipped. So a run that
// hits the 6-minute ceiling can simply be run again, and will carry on from where
// it stopped rather than redoing what it did.
// ─────────────────────────────────────────────────────────────────────────────

const RATE_MOVE_SOURCE_REF = 'Copied from test portal';

/** Group active rows by route and charge code: 'MID|CODE' -> [rows oldest first]. */
function rateMoveGroups_(rowsByKey) {
  const g = {};
  Object.keys(rowsByKey).forEach(function (k) {
    const r = rowsByKey[k];
    const gk = r.mid + '|' + r.code;
    (g[gk] = g[gk] || []).push(r);
  });
  Object.keys(g).forEach(function (gk) {
    g[gk].sort(function (a, b) { return dateKey(a.from) - dateKey(b.from); });
  });
  return g;
}

/** Two schedules for one route: same periods, same money, same order? */
function rateMoveGroupSame_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].from !== b[i].from) return false;
    if (!rateDiffSame_(a[i], b[i])) return false;
  }
  return true;
}

/**
 * Copy rate schedules from another spreadsheet into this project's.
 *
 * @param {boolean} commit   false (default) logs the plan and writes nothing
 * @param {string=} otherId  source; defaults to TEST_SPREADSHEET_ID
 */
function copyRatesFromOtherSpreadsheet(commit, otherId) {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log(commit ? '=== COPYING RATE DATA IN (WRITING) ==='
                    : '=== RATE DATA COPY — PREVIEW (nothing written) ===');

  const perms = requirePermissions_();
  requireEditRates_(perms);

  const thisId = spreadsheetId_();
  const wanted = safeStr(otherId) || safeStr(TEST_SPREADSHEET_ID);
  if (!wanted || wanted === 'PASTE_THE_TEST_SPREADSHEET_ID_HERE') {
    throw new Error('No source spreadsheet. Pass an ID, or set TEST_SPREADSHEET_ID.');
  }
  if (wanted === thisId) {
    throw new Error('Source and destination are the same spreadsheet (' + thisId + ').');
  }

  let thisSs, otherSs;
  try { thisSs = SpreadsheetApp.openById(thisId); }
  catch (e) { throw new Error('Cannot open the destination ' + thisId + ': ' + e.message); }
  try { otherSs = SpreadsheetApp.openById(wanted); }
  catch (e) { throw new Error('Cannot open the source ' + wanted + ': ' + e.message); }

  Logger.log('  FROM (source, read only) : ' + otherSs.getName() + '   ' + wanted);
  Logger.log('  INTO (written to)        : ' + thisSs.getName() + '   ' + thisId);
  Logger.log('  running as               : ' + perms.email + '  role ' + perms.role);
  Logger.log('');
  Logger.log('  Unit of transfer is one route + one charge: where the schedules differ,');
  Logger.log('  this side\'s periods are deactivated and the source\'s written in their');
  Logger.log('  place. A route-and-charge that exists only HERE is left alone.');

  /* Which Modelling IDs the destination actually has. A rate for a route that does
     not exist here cannot be written, and saying so is more useful than letting
     saveBaseRate refuse it one row at a time. */
  const destMids = {};
  (function () {
    const sh = thisSs.getSheetByName(SHEET.MODELLING_IDS);
    if (!sh) return;
    const d = sh.getDataRange().getValues(), M = COL.MODELLING_IDS;
    for (let i = 1; i < d.length; i++) {
      const id = safeInt(d[i][M.Modelling_ID]);
      if (id) destMids[id] = safeBool(d[i][M.Active]);
    }
  })();

  const plan = [], skipped = [], onlyHere = [];

  rateDiffTables_().forEach(function (t) {
    Logger.log('');
    Logger.log('════ ' + t.label + ' ════');

    const mineRows = rateDiffRows_(thisSs, t), theirRows = rateDiffRows_(otherSs, t);
    if (!mineRows || !theirRows) {
      Logger.log('  Tab missing on ' + (!mineRows ? 'the destination' : 'the source') +
                 ' — skipped entirely.');
      return;
    }
    const mine = rateMoveGroups_(mineRows), theirs = rateMoveGroups_(theirRows);

    Object.keys(theirs).sort(function (a, b) {
      return safeInt(a) - safeInt(b) || (a < b ? -1 : 1); }).forEach(function (gk) {
      const src = theirs[gk], dst = mine[gk] || [];
      const mid = src[0].mid, code = src[0].code;
      const label = '  MID ' + pad_(mid, 4) + (code ? '  ' + pad_(code, 10) : '') + '  ';

      if (rateMoveGroupSame_(dst, src)) return;            // already identical — silent

      if (destMids[mid] === undefined) {
        Logger.log(label + 'SKIPPED — no Modelling ID ' + mid + ' in the destination');
        skipped.push({ table: t.label, mid: mid, code: code,
                       reason: 'Modelling ID not in destination' });
        return;
      }
      if (!destMids[mid]) {
        Logger.log(label + 'SKIPPED — Modelling ID ' + mid + ' is inactive here');
        skipped.push({ table: t.label, mid: mid, code: code, reason: 'route inactive here' });
        return;
      }
      if (!canSeeModellingId_(perms, mid, true)) {
        Logger.log(label + 'SKIPPED — outside what ' + perms.email + ' can edit');
        skipped.push({ table: t.label, mid: mid, code: code, reason: 'outside editable scope' });
        return;
      }

      Logger.log(label + dst.length + ' period(s) here -> ' + src.length + ' from source');
      dst.forEach(function (r) { Logger.log('        was: ' + rateDiffLine_(r)); });
      src.forEach(function (r) { Logger.log('        now: ' + rateDiffLine_(r)); });

      plan.push({ table: t, mid: mid, code: code,
                  deleteIds: dst.map(function (r) { return r.id; }),
                  rows: src });
    });

    /* Groups the destination has and the source does not. Never touched. */
    Object.keys(mine).forEach(function (gk) {
      if (theirs[gk]) return;
      const r = mine[gk][0];
      onlyHere.push({ table: t.label, mid: r.mid, code: r.code, periods: mine[gk].length });
    });
  });

  const deletes = plan.reduce(function (a, x) { return a + x.deleteIds.length; }, 0);
  const creates = plan.reduce(function (a, x) { return a + x.rows.length; }, 0);

  Logger.log('');
  Logger.log('--- summary ---');
  Logger.log('  route+charge groups to change : ' + plan.length);
  Logger.log('  periods to deactivate here    : ' + deletes);
  Logger.log('  periods to write from source  : ' + creates);
  Logger.log('  groups skipped                : ' + skipped.length);
  Logger.log('  write calls if committed      : ' + (deletes + creates));

  if (onlyHere.length) {
    Logger.log('');
    Logger.log('--- ' + onlyHere.length + ' route+charge group(s) exist HERE and not in the ' +
               'source: LEFT ALONE ---');
    onlyHere.slice(0, 40).forEach(function (o) {
      Logger.log('    ' + o.table + '  MID ' + pad_(o.mid, 4) +
                 (o.code ? '  ' + o.code : '') + '  ' + o.periods + ' period(s)');
    });
    if (onlyHere.length > 40) Logger.log('    ... and ' + (onlyHere.length - 40) + ' more');
    Logger.log('  These are NOT deleted. The source is a snapshot, so its silence about a');
    Logger.log('  route means the snapshot predates it, not that it should go. If any of');
    Logger.log('  them genuinely should be removed, remove them deliberately by hand.');
  }

  if (deletes + creates > 150) {
    Logger.log('');
    Logger.log('  ** ' + (deletes + creates) + ' write calls, each taking the script lock and');
    Logger.log('     re-reading the tab. That may not finish inside the 6-minute ceiling.');
    Logger.log('     Re-running is safe and resumes: groups that already match are skipped.');
  }

  if (!commit) {
    Logger.log('');
    Logger.log('  NOTHING WRITTEN. Read the was/now lines above, then run');
    Logger.log('  runRateCopyFromTestCopy() to apply exactly this plan.');
    return { ok: true, preview: true, from: { id: wanted, name: otherSs.getName() },
             into: { id: thisId, name: thisSs.getName() },
             groups: plan.length, periodsToDeactivate: deletes, periodsToWrite: creates,
             writeCalls: deletes + creates, skipped: skipped, onlyHere: onlyHere };
  }

  if (!plan.length) throw new Error('Nothing to copy — every group already matches or was ' +
    'skipped. Reasons are in the log above.');

  // ---- write ---------------------------------------------------------------
  Logger.log('');
  Logger.log('--- writing ---');
  const done = [], failed = [];

  plan.forEach(function (x) {
    const isBase = x.table.key === 'RATE_BASE';
    const note = 'Copied from ' + otherSs.getName();
    let deactivated = 0, created = 0;
    try {
      x.deleteIds.forEach(function (id) {
        if (isBase) deleteBaseRate(id); else deleteSurchargeRate(id);
        deactivated++;
      });
      x.rows.forEach(function (r) {
        const p = { modellingId: x.mid, validFrom: r.from, validTo: r.to,
                    value: r.value, currency: r.currency,
                    scenarioId: r.scenarioId,
                    sourceRef: RATE_MOVE_SOURCE_REF, notes: note };
        if (isBase) saveBaseRate(p);
        else { p.code = x.code; saveSurchargeRate(p); }
        created++;
      });
      Logger.log('  MID ' + pad_(x.mid, 4) + (x.code ? '  ' + pad_(x.code, 10) : '') +
                 '  OK — ' + deactivated + ' deactivated, ' + created + ' written');
      done.push({ mid: x.mid, code: x.code, deactivated: deactivated, created: created });
    } catch (e) {
      /* Part way through a group is the bad case: old periods off, new ones
         incomplete, so the route has a GAP in its rate schedule and the engine
         will price it at nothing for those months. Said loudly. */
      const partial = (deactivated > 0 || created > 0) && created < x.rows.length;
      Logger.log('  MID ' + pad_(x.mid, 4) + (x.code ? '  ' + pad_(x.code, 10) : '') +
                 '  ' + (partial ? '** PART DONE **' : 'FAILED') + ' — ' + e.message);
      if (partial) {
        Logger.log('      ' + deactivated + ' deactivated, only ' + created + ' of ' +
                   x.rows.length + ' written. This route has a GAP in its schedule now.');
      }
      failed.push({ mid: x.mid, code: x.code, error: e.message, partial: partial });
    }
  });

  const partials = failed.filter(function (f) { return f.partial; });
  Logger.log('');
  Logger.log('--- done in ' + Math.round((Date.now() - t0) / 1000) + 's ---');
  Logger.log('  changed : ' + done.length + ' group(s)');
  Logger.log('  failed  : ' + failed.length);
  Logger.log('  skipped : ' + skipped.length);
  if (partials.length) {
    Logger.log('  ** PART DONE, schedule has a gap: ' +
               partials.map(function (f) { return f.mid + (f.code ? '/' + f.code : ''); })
                       .join(', '));
  }
  Logger.log('');
  Logger.log('  Every row went through saveBaseRate / saveSurchargeRate, so it is all in');
  Logger.log('  the Amends tables and the Audit_Log. Re-run the calculation and publish for');
  Logger.log('  these to reach the forecast.');

  return { ok: !failed.length, preview: false,
           from: { id: wanted, name: otherSs.getName() },
           into: { id: thisId, name: thisSs.getName() },
           changed: done.length, failed: failed, skipped: skipped,
           partial: partials, onlyHere: onlyHere,
           seconds: Math.round((Date.now() - t0) / 1000) };
}

/** Preview the copy from the test copy named in utils.gs. Writes nothing. */
function previewRateCopyFromTestCopy() { return copyRatesFromOtherSpreadsheet(false, null); }
/** Apply it. */
function runRateCopyFromTestCopy()     { return copyRatesFromOtherSpreadsheet(true, null); }
