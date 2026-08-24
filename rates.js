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
