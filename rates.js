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
// BULK UPDATE BY DIMENSION
//
// The Rates screen edits one route at a time, which is right for "Royal Mail
// RM24 letters in GB went up". It is the wrong shape for "the 24-hour letter
// surcharge went up everywhere", which is one commercial decision spread across
// dozens of routes and, done by hand, dozens of chances to miss one.
//
// So a change can name a value for each dimension or leave it as All, and the
// change lands on every EXISTING row that matches. "Existing" is the operative
// word: this updates rows that are already there, it never creates a rate for a
// route that had none. A route with no rate is a structural gap that wants
// noticing on the Rates screen, not filling in silently by a bulk action.
//
// THE DIMENSIONS. Neither rate table stores any of them. Both key on
// Modelling_ID, and every dimension is reached by a two-hop join:
//
//   Rate_Base / Rate_Surcharge  ->  Modelling_IDs  ->  High_Level_IDs
//
//   High_Level_IDs : Brand, Geo, Treatment_Type, WL_Split
//   Modelling_IDs  : Carrier_Code, Method_Code, Letter_Parcel
//   Rate_Surcharge : Surcharge_Code   (surcharges only, and never All — see below)
//
// ALL IS NOT A STORED VALUE. It exists only in the selector: an empty string
// means "do not filter on this dimension". Nothing is written as "All" — every
// row written is an ordinary row against one concrete Modelling_ID. Note that
// WL_Split legitimately stores '*' meaning "not applicable" for segments that are
// not weight-loss; that is a real value to be matched, NOT a wildcard, and
// conflating the two would silently widen every change.
//
// SURCHARGE TYPE IS ALWAYS SPECIFIC. FUEL is a percentage and GREEN, PEAK and
// SHIPSTATION are amounts (Dim_Surcharge.Value_Type), so one entered number
// cannot mean both a fraction and a currency value. Allowing All there would mean
// 0.06 landing as 6% on one row and 6 pence on the next.
//
// SAME RULES, BATCHED WRITES. Every check a single save makes is made here, per
// row, by the same functions: assertCanEditModellingId_ for scope,
// precedingPeriodCloses_ for the period-closing rule, assertNoOverlapIn_ for
// overlaps, and the identical value validation. What changes is only how the
// result reaches the sheet. A single save costs about seven Sheets API calls;
// Rate_Surcharge holds some 1,650 rows and a wide selection can match hundreds of
// them, which at seven calls each would pass the six-minute execution limit and
// leave a half-applied price change behind. So the plan is built in memory, then
// flushed as one ranged write for the closed periods, one append for the new
// rows, and one recordChangesBatch_ for the history — which produces the same
// Amends snapshots and the same per-field Audit_Log rows the single path does.
//
// PREVIEW FIRST, ALWAYS. preview defaults to true and returns the matched rows
// with their current values; applying requires passing back the planKey the
// preview issued. If anything about the matched set has moved in between — an
// edit by somebody else, a route deactivated — the key no longer matches and the
// write is refused rather than applied to a set the user never saw.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Does a selected dimension value match a row's value?
 *
 * Empty means All. There is deliberately no magic string: a literal 'ALL' is
 * treated as a value to match, because a carrier or brand code called ALL is one
 * data operation away, and a wildcard that a data entry can impersonate is a
 * wildcard waiting to widen someone's rate change.
 */
function dimensionMatches_(selected, rowValue) {
  const want = safeStr(selected);
  if (want === '') return true;
  return normKey(rowValue) === normKey(want);
}

/**
 * The active, editable Modelling IDs matching a dimension selection.
 *
 * Routes the caller cannot edit are counted and skipped rather than refused —
 * the same choice bulkRateChange makes, so a scoped user gets their own slice of
 * a company-wide change instead of an error naming segments they cannot see.
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
    if (!dimensionMatches_(sel.carrier,      md[i][M.Carrier_Code]))  continue;
    if (!dimensionMatches_(sel.method,       md[i][M.Method_Code]))   continue;
    if (!dimensionMatches_(sel.letterParcel, md[i][M.Letter_Parcel])) continue;

    out.routesMatched++;
    if (!canSeeModellingId_(perms, id, true)) { out.skippedNoScope++; continue; }
    out.ids.push(id);
  }
  return out;
}

/** A human description of a selection, for the audit trail and the preview. */
function describeDimensions_(dims, code) {
  const sel = dims || {};
  const parts = [];
  [['brand', 'brand'], ['geo', 'geo'], ['treatmentType', 'treatment'],
   ['wlSplit', 'WL split'], ['carrier', 'carrier'], ['method', 'method'],
   ['letterParcel', 'class']].forEach(function (pair) {
    const v = safeStr(sel[pair[0]]);
    parts.push(pair[1] + '=' + (v === '' ? 'All' : v.toUpperCase()));
  });
  if (code) parts.unshift('surcharge=' + code);
  return parts.join(', ');
}

/** Which table a bulk update is against, and how its rows are keyed. */
function bulkRateShape_(kind, code) {
  if (kind === 'SURCHARGE') {
    const SC = COL.RATE_SURCHARGE;
    return {
      kind: 'SURCHARGE', tableKey: 'RATE_SURCHARGE', table: TABLES.RATE_SURCHARGE,
      C: SC, idCol: SC.Surcharge_Rate_ID, valueCol: SC.Value,
      keyMatch: function (mid) {
        return function (r) {
          return safeInt(r[SC.Modelling_ID]) === mid &&
                 normKey(r[SC.Surcharge_Code]) === code;
        };
      },
      label: function (mid) { return 'Modelling ID ' + mid + ' ' + code; }
    };
  }
  const BC = COL.RATE_BASE;
  return {
    kind: 'BASE', tableKey: 'RATE_BASE', table: TABLES.RATE_BASE,
    C: BC, idCol: BC.Rate_ID, valueCol: BC.Base_Rate,
    keyMatch: function (mid) {
      return function (r) { return safeInt(r[BC.Modelling_ID]) === mid; };
    },
    label: function (mid) { return 'Modelling ID ' + mid; }
  };
}

/**
 * Work out what a bulk update would do, without touching anything.
 *
 * For each matching route this finds the row a new period would supersede: the
 * one in force on the new start date, or failing that the latest one starting
 * before it. A route with no such row is reported as skipped, not created.
 *
 * Scenario_ID is deliberately NOT part of the match, because the single-save path
 * does not filter on it either (finding C5). With one scenario the two are
 * identical; when a second exists, C5 has to be fixed across all of these paths
 * in one go rather than diverging here first.
 */
function planBulkRateUpdate_(p, perms) {
  const kind  = normKey(p.kind) === 'SURCHARGE' ? 'SURCHARGE' : 'BASE';
  const code  = kind === 'SURCHARGE' ? normKey(p.surchargeCode) : '';
  const value = safeNum(p.value);

  if (kind === 'SURCHARGE') {
    if (!code) throw new Error('Choose which surcharge type to change. "All" is ' +
      'not available there, because percentage and fixed-amount surcharges ' +
      'cannot share one value.');
    const def = surchargeDefinition_(code);
    if (!def) throw new Error('Unknown surcharge type "' + code + '".');
    if (def.valueType === 'PCT' && (value < -1 || value > 2)) {
      throw new Error('A percentage surcharge should be a fraction — 0.14 for 14%.');
    }
  } else if (value < 0) {
    throw new Error('A base rate cannot be negative.');
  }

  const d = validateDates_(p.validFrom, p.validTo);
  const shape = bulkRateShape_(kind, code);
  const C = shape.C;
  const targets = resolveDimensionTargets_(p.dimensions, perms);
  const data = getAllData_(shape.table.sheet);

  const items = [];
  let skippedNoRate = 0;

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
    items.push({ modellingId: mid, rateId: cur.id, from: cur.value, to: value,
                 diff: value - cur.value,
                 pct: cur.value ? (value - cur.value) / cur.value * 100 : 0,
                 currentFrom: cur.validFrom, currentTo: cur.validTo });
  });

  // Distinct current values, so the preview can say what is about to be
  // flattened. Selection is by dimension alone, which makes showing the spread of
  // what is being overwritten the safeguard rather than a filter on it.
  const byValue = {};
  items.forEach(function (it) {
    const k = String(Math.round(it.from * 1000000) / 1000000);
    byValue[k] = (byValue[k] || 0) + 1;
  });
  const distinct = Object.keys(byValue)
    .map(function (k) { return { value: safeNum(k), rows: byValue[k] }; })
    .sort(function (a, b) { return b.rows - a.rows || a.value - b.value; });

  return {
    kind: kind, code: code, value: value, dates: d, shape: shape, items: items,
    count: items.length, routesMatched: targets.routesMatched,
    skippedNoScope: targets.skippedNoScope, skippedNoRate: skippedNoRate,
    distinct: distinct, planKey: planKey_(items),
    selection: describeDimensions_(p.dimensions, code)
  };
}

/**
 * A short fingerprint of a plan's matched rows and their current values.
 *
 * The client previews, the user confirms, and the apply call sends this back. If
 * the matched set has moved in between — someone else edited a rate, a route was
 * deactivated — the recomputed key differs and the write is refused, so a
 * confirmation can only ever apply the set that was actually shown.
 */
function planKey_(items) {
  const s = items.map(function (it) {
    return it.rateId + ':' + Math.round(it.from * 1000000);
  }).sort().join('|');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s || 'empty');
  return bytes.map(function (b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('').slice(0, 12);
}

/**
 * Bulk update base rates or surcharges across a dimension selection.
 *
 * Previews by default. Pass preview:false together with the planKey the preview
 * returned to actually write.
 *
 * @param {Object} p { kind: 'BASE'|'SURCHARGE', surchargeCode,
 *                     dimensions: { brand, geo, treatmentType, wlSplit,
 *                                   carrier, method, letterParcel },
 *                     validFrom, validTo, value, currency, sourceRef,
 *                     scenarioId, preview, planKey }
 */
function bulkUpdateRates(p) {
  prewarmForWrite_([SHEET.RATE_BASE, SHEET.RATE_BASE_AMENDS,
                    SHEET.RATE_SURCHARGE, SHEET.RATE_SURCHARGE_AMENDS,
                    SHEET.MODELLING_IDS, SHEET.HIGH_LEVEL_IDS, SHEET.DIM_SURCHARGE]);
  const perms = requirePermissions_();
  requireEditRates_(perms);

  const plan = planBulkRateUpdate_(p, perms);

  if (p.preview !== false) {
    return { preview: true, kind: plan.kind, code: plan.code,
             count: plan.count, routesMatched: plan.routesMatched,
             skippedNoScope: plan.skippedNoScope, skippedNoRate: plan.skippedNoRate,
             distinct: plan.distinct, planKey: plan.planKey,
             selection: plan.selection, value: plan.value,
             validFrom: fmtDate(plan.dates.from), validTo: fmtDate(plan.dates.to),
             changes: plan.items };
  }

  if (!plan.count) throw new Error('Nothing matches that selection, so there is ' +
    'nothing to update. ' + plan.skippedNoRate + ' matching route(s) have no rate ' +
    'to supersede, and ' + plan.skippedNoScope + ' are outside what you can edit.');

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
 * Runs the same per-row rules as a single save against one in-memory copy of the
 * table, then flushes. The copy is mutated as it goes — closing a period for one
 * route before planning the next — so each route's overlap check sees the closes
 * and the new rows that earlier routes in the same batch produced.
 */
function applyBulkRateUpdate_(p, plan, perms) {
  const shape = plan.shape, C = shape.C, width = shape.table.headers.length;
  const sheetName = shape.table.sheet;
  invalidateSheetCache_(sheetName);

  // A padded copy, not the cached array itself. Padding once up front means the
  // whole function works on full-width rows, which is what makes the ranged write
  // at the end safe (finding C6); copying means a refusal part-way through the
  // batch cannot leave half-applied values behind in the execution cache for
  // whatever reads next.
  const source = getAllData_(sheetName);
  const originalRows = source.length - 1;   // before any new row is appended
  const data = source.map(function (r) { return padRow_(r, width); });
  const now = new Date();
  const batchRef = bulkBatchRef_(now);
  const scenarioId = safeInt(p.scenarioId) || 1;

  const changed = {};        // dataIndex -> true, for rows whose period was closed
  const creates = [];
  const history = [];
  let nextId = getNextId_(sheetName, shape.idCol);

  plan.items.forEach(function (item) {
    const mid = item.modellingId;
    assertCanEditModellingId_(perms, mid);     // per row, as a single save does
    const matches = shape.keyMatch(mid);

    // 1. close whatever period is running — the same rule the single save uses
    precedingPeriodCloses_(data, C, matches, plan.dates.from).forEach(function (c) {
      const before = data[c.dataIndex].slice();
      const after = before.slice();
      after[C.Valid_To]   = c.newTo;
      after[C.Updated_TS] = now;
      after[C.Updated_By] = perms.email;
      data[c.dataIndex] = after;
      changed[c.dataIndex] = true;
      history.push({ id: safeInt(before[shape.idCol]), before: before,
                     after: after, type: 'UPDATE' });
    });

    // 2. refuse an overlap — the same rule, against the copy as it now stands
    assertNoOverlapIn_(data, C, matches, plan.dates.fromKey, plan.dates.toKey,
                       0, shape.idCol, shape.label(mid));

    // 3. the new period
    const row = blankRow_(shape.tableKey);
    row[shape.idCol]     = nextId++;
    row[C.Modelling_ID]  = mid;
    if (shape.kind === 'SURCHARGE') row[C.Surcharge_Code] = plan.code;
    row[C.Valid_From]    = plan.dates.from;
    row[C.Valid_To]      = plan.dates.to;
    row[shape.valueCol]  = plan.value;
    row[C.Currency]      = bulkRowCurrency_(shape, plan, mid, p);
    row[C.Scenario_ID]   = scenarioId;
    row[C.Source_Ref]    = safeStr(p.sourceRef) || batchRef;
    row[C.Notes]         = truncate_('Bulk ' + batchRef + ': set to ' + plan.value +
                                     ' (' + plan.selection + ')');
    row[C.Active]        = true;
    row[C.Created_TS]    = now;
    row[C.Created_By]    = perms.email;
    row[C.Updated_TS]    = now;
    row[C.Updated_By]    = perms.email;

    creates.push(row);
    data.push(row);      // so the next route's overlap check sees it
    history.push({ id: row[shape.idCol], before: null, after: row, type: 'CREATE' });
  });

  // ---- flush ---------------------------------------------------------------
  // One ranged write for the closed periods rather than one per row. Rows come
  // back from batchGet with trailing empty cells omitted, so every row is padded
  // to the header width before being written back — finding C6, which is exactly
  // this mistake made in structure.gs.
  const sh = getSheet_(sheetName);
  if (Object.keys(changed).length) {
    sh.getRange(2, 1, originalRows, width).setValues(data.slice(1, 1 + originalRows));
  }
  if (creates.length) {
    sh.getRange(sh.getLastRow() + 1, 1, creates.length, width).setValues(creates);
  }
  invalidateSheetCache_(sheetName);

  // The same Amends snapshots and the same per-field Audit_Log rows a single save
  // makes, in a handful of calls instead of two per row, each stamped with the
  // batch reference so History can show them as one action.
  recordChangesBatch_(shape.tableKey, history, batchRef);
  logAudit_('UPDATE', 'BULK_RATE_UPDATE', batchRef, '', '',
            plan.count + ' rows set to ' + plan.value,
            plan.kind + ' ' + (plan.code || '') + ' from ' +
            fmtDate(plan.dates.from) + ' — ' + plan.selection, true);

  return { preview: false, batchRef: batchRef, written: plan.count,
           rowsClosed: Object.keys(changed).length, rowsCreated: creates.length,
           skippedNoScope: plan.skippedNoScope, skippedNoRate: plan.skippedNoRate,
           changes: plan.items };
}

/**
 * The currency for a bulk-written row.
 *
 * Base rates and amount surcharges carry one; percentage surcharges must not, or
 * the Rates screen shows "0.14 GBP" for what is 14%. Mirrors what saveBaseRate
 * and saveSurchargeRate each do for a single row.
 */
function bulkRowCurrency_(shape, plan, mid, p) {
  if (shape.kind === 'SURCHARGE') {
    const def = surchargeDefinition_(plan.code);
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
function testBulkRateUpdate() {
  requireMaintenance_();
  Logger.log('=== BULK RATE UPDATE TEST ===');
  const perms = requirePermissions_();
  Logger.log('  acting as ' + perms.email + ' (' + perms.role + ')');

  const t = TABLES.RATE_SURCHARGE, C = COL.RATE_SURCHARGE, width = t.headers.length;
  const checks = [];
  function check(name, pass, detail) {
    checks.push({ name: name, pass: !!pass });
    Logger.log('  ' + (pass ? 'ok   ' : 'FAIL ') + name + (detail ? '  — ' + detail : ''));
  }

  // ---- 1. dimension resolution -------------------------------------------
  Logger.log('');
  Logger.log('--- dimension resolution ---');
  const all = resolveDimensionTargets_({}, perms);
  Logger.log('  every dimension All        : ' + all.ids.length + ' editable routes (' +
             all.routesMatched + ' matched, ' + all.skippedNoScope + ' out of scope)');
  check('All-everything matches routes', all.ids.length > 0);

  const rm = resolveDimensionTargets_({ carrier: 'ROYALMAIL' }, perms);
  const rmLetter = resolveDimensionTargets_({ carrier: 'ROYALMAIL', letterParcel: 'LETTER' }, perms);
  Logger.log('  carrier=ROYALMAIL          : ' + rm.ids.length);
  Logger.log('  + class=LETTER             : ' + rmLetter.ids.length);
  check('naming a dimension narrows the set',
        rm.ids.length <= all.ids.length && rmLetter.ids.length <= rm.ids.length);

  // WL_Split '*' means "not applicable" and is a value to match, not a wildcard.
  const star = resolveDimensionTargets_({ wlSplit: '*' }, perms);
  Logger.log("  wlSplit='*'                : " + star.ids.length +
             '  (must be fewer than All — it is a value, not a wildcard)');
  check("WL_Split '*' is matched as a value", star.ids.length < all.ids.length);

  const nobody = resolveDimensionTargets_({ brand: 'NO_SUCH_BRAND' }, perms);
  check('an unknown value matches nothing', nobody.ids.length === 0);

  // ---- 2. pick a real, narrow selection ----------------------------------
  // Built from an existing surcharge row, so the test always has something to
  // supersede: All brands / geos / treatments / splits, with this row's carrier,
  // method, class and surcharge type named. That is example 1's exact shape.
  invalidateSheetCache_(t.sheet);
  const rows = getAllData_(t.sheet);
  const mdById = {};
  const md = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < md.length; i++) {
    mdById[safeInt(md[i][M.Modelling_ID])] = {
      carrier: safeStr(md[i][M.Carrier_Code]), method: safeStr(md[i][M.Method_Code]),
      lp: safeStr(md[i][M.Letter_Parcel])
    };
  }

  let seed = null;
  for (let i = 1; i < rows.length && !seed; i++) {
    if (!safeBool(rows[i][C.Active])) continue;
    const route = mdById[safeInt(rows[i][C.Modelling_ID])];
    if (!route || !route.carrier || !route.method) continue;
    seed = { code: normKey(rows[i][C.Surcharge_Code]), route: route };
  }
  if (!seed) {
    Logger.log('');
    Logger.log('NO ACTIVE SURCHARGE ROWS — nothing to test against. Load some rates first.');
    return { ok: false, reason: 'no surcharge rows' };
  }

  const dimensions = { brand: '', geo: '', treatmentType: '', wlSplit: '',
                       carrier: seed.route.carrier, method: seed.route.method,
                       letterParcel: seed.route.lp };
  const req = { kind: 'SURCHARGE', surchargeCode: seed.code, dimensions: dimensions,
                validFrom: '2035-01-01', validTo: '9999-12-31', value: 0.0777,
                sourceRef: 'testBulkRateUpdate' };

  Logger.log('');
  Logger.log('--- preview ---');
  Logger.log('  ' + describeDimensions_(dimensions, seed.code));
  const pv = bulkUpdateRates(req);
  Logger.log('  rows that would change     : ' + pv.count);
  Logger.log('  routes matched             : ' + pv.routesMatched +
             '  (no rate to supersede: ' + pv.skippedNoRate +
             ', out of scope: ' + pv.skippedNoScope + ')');
  Logger.log('  current values found       :');
  pv.distinct.forEach(function (d) {
    Logger.log('    ' + d.value + '  x ' + d.rows + ' row(s)');
  });
  Logger.log('  planKey                    : ' + pv.planKey);
  check('preview matched at least one row', pv.count > 0);
  check('preview wrote nothing',
        getAllData_(t.sheet).length === rows.length, 'row count unchanged');

  const pv2 = bulkUpdateRates(req);
  check('the same selection gives the same planKey', pv2.planKey === pv.planKey);

  // ---- 3. a stale confirmation must be refused ---------------------------
  Logger.log('');
  Logger.log('--- a wrong planKey must be refused ---');
  let staleRefused = false;
  try {
    bulkUpdateRates(Object.assign({}, req, { preview: false, planKey: 'deadbeef0000' }));
  } catch (e) {
    staleRefused = true;
    Logger.log('  refused, correctly: ' + e.message);
  }
  check('a stale planKey is refused', staleRefused);

  // ---- 4. a percentage surcharge cannot take an absurd value -------------
  const def = surchargeDefinition_(seed.code);
  if (def && def.valueType === 'PCT') {
    let pctRefused = false;
    try {
      bulkUpdateRates(Object.assign({}, req, { value: 45 }));
    } catch (e) { pctRefused = true; }
    check('a percentage surcharge refuses 45 (i.e. 4500%)', pctRefused);
  } else {
    Logger.log('  (skipped the percentage check — ' + seed.code + ' is an amount)');
  }

  let missingCodeRefused = false;
  try {
    bulkUpdateRates(Object.assign({}, req, { surchargeCode: '' }));
  } catch (e) { missingCodeRefused = true; }
  check('a surcharge bulk update refuses All for the type', missingCodeRefused);

  // ---- 5. apply, verify, restore -----------------------------------------
  const CAP = 25;
  if (pv.count > CAP) {
    Logger.log('');
    Logger.log('SKIPPING THE WRITE — this selection matches ' + pv.count + ' rows, more');
    Logger.log('than the ' + CAP + ' this diagnostic is willing to touch. The preview half');
    Logger.log('passed. Narrow the seed data, or do the write half from the app.');
    return summariseBulkTest_(checks, null);
  }

  Logger.log('');
  Logger.log('--- snapshot, apply, verify, restore ---');
  invalidateSheetCache_(t.sheet);
  const snapshot = getAllData_(t.sheet).map(function (r) { return padRow_(r, width); });
  const beforeById = {};
  for (let i = 1; i < snapshot.length; i++) {
    beforeById[safeInt(snapshot[i][C.Surcharge_Rate_ID])] = snapshot[i];
  }
  const auditBefore = Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0);
  Logger.log('  snapshotted ' + (snapshot.length - 1) + ' rows');

  const res = bulkUpdateRates(Object.assign({}, req,
                { preview: false, planKey: pv.planKey }));
  Logger.log('  batch reference            : ' + res.batchRef);
  Logger.log('  periods closed             : ' + res.rowsClosed);
  Logger.log('  rows created               : ' + res.rowsCreated);
  check('one new row per matched row', res.rowsCreated === pv.count);
  check('a batch reference was issued', /^BULK-\d{8}-\d{6}$/.test(String(res.batchRef)));

  invalidateSheetCache_(t.sheet);
  const after = getAllData_(t.sheet).map(function (r) { return padRow_(r, width); });
  const fresh = [];
  for (let i = 1; i < after.length; i++) {
    if (!beforeById[safeInt(after[i][C.Surcharge_Rate_ID])]) fresh.push(after[i]);
  }
  check('the new rows are on the sheet', fresh.length === pv.count);
  check('every new row carries the value asked for',
        fresh.length > 0 && fresh.every(function (r) { return safeNum(r[C.Value]) === 0.0777; }));
  check('every new row starts on the date asked for',
        fresh.length > 0 && fresh.every(function (r) { return fmtDate(r[C.Valid_From]) === '2035-01-01'; }));
  check('every new row names the batch in Notes',
        fresh.length > 0 && fresh.every(function (r) {
          return safeStr(r[C.Notes]).indexOf(res.batchRef) >= 0; }));
  check('every new row is on the surcharge type asked for',
        fresh.length > 0 && fresh.every(function (r) {
          return normKey(r[C.Surcharge_Code]) === seed.code; }));
  check('percentage surcharges carry no currency',
        !def || def.valueType !== 'PCT' ||
        fresh.every(function (r) { return safeStr(r[C.Currency]) === ''; }));

  // Every closed period must end the day before the new one starts.
  let closedOk = true;
  for (let i = 1; i < after.length; i++) {
    const id = safeInt(after[i][C.Surcharge_Rate_ID]);
    const was = beforeById[id];
    if (!was) continue;
    if (fmtDate(was[C.Valid_To]) === fmtDate(after[i][C.Valid_To])) continue;
    if (fmtDate(after[i][C.Valid_To]) !== '2034-12-31') closedOk = false;
  }
  check('closed periods end the day before the new one', closedOk);

  invalidateSheetCache_(SHEET.AUDIT_LOG);
  const auditRows = getAllData_(SHEET.AUDIT_LOG), A = COL.AUDIT_LOG;
  let stamped = 0;
  for (let i = 1; i < auditRows.length; i++) {
    if (safeStr(auditRows[i][A.Detail]).indexOf(res.batchRef) >= 0) stamped++;
  }
  const auditAdded = Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0) - auditBefore;
  Logger.log('  Audit_Log rows added       : ' + auditAdded +
             ' (' + stamped + ' stamped with the batch reference)');
  check('every audit row is stamped with the batch', stamped >= res.rowsCreated);

  // ---- restore ------------------------------------------------------------
  const restored = [];
  for (let i = 1; i < after.length; i++) {
    const id = safeInt(after[i][C.Surcharge_Rate_ID]);
    const was = beforeById[id];
    if (was) { restored.push(was); continue; }
    const row = after[i].slice();
    row[C.Active] = false;
    row[C.Notes]  = truncate_('TEST ROW from testBulkRateUpdate — deactivated, safe to delete');
    restored.push(row);
  }
  getSheet_(t.sheet).getRange(2, 1, restored.length, width).setValues(restored);
  invalidateSheetCache_(t.sheet);
  Logger.log('  restored ' + (snapshot.length - 1) + ' rows and deactivated ' +
             fresh.length + ' test row(s)');

  invalidateSheetCache_(t.sheet);
  const final = getAllData_(t.sheet).map(function (r) { return padRow_(r, width); });
  let identical = true;
  for (let i = 1; i < snapshot.length; i++) {
    for (let j = 0; j < width; j++) {
      if (displayValue_(snapshot[i][j]) !== displayValue_(final[i][j])) identical = false;
    }
  }
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
    Logger.log('The table was restored, so the only trace left is history: the Audit_Log');
    Logger.log('and Rate_Surcharge_Amends rows for batch ' + batchRef + ', plus the');
    Logger.log('deactivated 2035 test rows. None of it can reach a forecast.');
  }
  return { ok: !failed.length, checks: checks.length, failed: failed.length,
           batchRef: batchRef };
}
