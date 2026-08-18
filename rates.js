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

/** Reject a period that overlaps another for the same key. */
function assertNoOverlap_(sheetName, C, keyMatch, fromKey, toKey, excludeId, idCol, label) {
  const data = getAllData_(sheetName);
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
  const newFromKey = dateKey(newFrom);

  // The day before the new period starts. Plain local-date arithmetic —
  // constructing it from year/month/day-1 lets the Date object handle month
  // and year rollover, and never touches UTC or a timestamp.
  const newTo = new Date(newFrom.getFullYear(), newFrom.getMonth(), newFrom.getDate() - 1);

  for (let i = 1; i < data.length; i++) {
    if (!safeBool(data[i][C.Active])) continue;
    if (!keyMatch(data[i])) continue;
    const f = dateKey(data[i][C.Valid_From]), t = dateKey(data[i][C.Valid_To]);
    if (f >= newFromKey || t < newFromKey) continue;   // not the running period
    if (dateKey(newTo) < f) continue;                  // would invert the period

    const before = readRow_(sheetName, i + 1, width);
    const row = before.slice();
    row[C.Valid_To]   = newTo;
    row[C.Updated_TS] = new Date();
    row[C.Updated_By] = perms.email;
    sh.getRange(i + 1, 1, 1, width).setValues([row]);
    invalidateSheetCache_(sheetName);
    recordChange_(tableKey, safeInt(before[idCol]), before, row, 'UPDATE');
  }
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