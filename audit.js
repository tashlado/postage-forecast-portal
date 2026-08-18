/**
 * Postage Forecast Portal — audit.gs
 *
 * Two records of every change, because they answer different questions.
 *
 *   Audit_Log     one row per field changed. Cheap. Answers "what happened,
 *                 who did it, when" and is what the History screen reads.
 *
 *   *_Amends      a full snapshot of the row as it was written. Answers
 *                 "show me the rate card exactly as it stood on 3 March" —
 *                 take every amend for that record up to that timestamp and
 *                 read the last one.
 *
 * Amends are written BY COLUMN NAME, not position, so adding a column to a
 * table later needs no change here.
 */

// ─────────────────────────────────────────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────────────────────────────────────────

// ---- Log_ID allocation ------------------------------------------------------
//
// Log_ID used to be max(existing) + 1 read from the sheet and then appended,
// which races: the DENIED path in auth.gs logs outside any lock, so two
// executions refused at the same moment read the same maximum and both wrote
// it. Taking the script lock here would fix that and cause worse: logAudit_ is
// called from inside write paths that already hold it, and the inner
// releaseLock() would hand the lock back mid-write — the same defect that made
// bulkRateChange unsafe. So the counter lives in a Script Property instead,
// which narrows the window to the microseconds between one small read and one
// small write, and never queues a refused call behind a real one.
//
// Seeded from the sheet's current maximum the first time it is used, so IDs
// carry on from where the sheet left off rather than restarting at 1, and keyed
// by spreadsheet ID because the target spreadsheet is switchable (see
// spreadsheetId_) while Script Properties belong to the script project — one
// shared counter would hand a copy's numbering to the original.
//
// IDs are reserved a block at a time so a grid write costs one property
// round-trip rather than one per row. A block left part-used leaves a gap in the
// numbering, which is harmless: nothing keys on Log_ID, it only has to be
// unique and ascending.

let _auditIdNext_ = 0;
let _auditIdLeft_ = 0;

/** Reserve `count` Log_IDs and return the first. */
function reserveAuditIds_(count) {
  const n = Math.max(1, Math.floor(count) || 1);
  const key = 'AUDIT_LOG_NEXT_ID_' + spreadsheetId_();
  const props = PropertiesService.getScriptProperties();

  let next = parseInt(props.getProperty(key), 10);
  if (!(next >= 1)) next = getNextId_(SHEET.AUDIT_LOG, COL.AUDIT_LOG.Log_ID);

  props.setProperty(key, String(next + n));
  return next;
}

/** The next Log_ID, taken from the block this execution has reserved. */
function nextAuditId_() {
  if (_auditIdLeft_ <= 0) {
    _auditIdNext_ = reserveAuditIds_(8);
    _auditIdLeft_ = 8;
  }
  _auditIdLeft_--;
  return _auditIdNext_++;
}


/**
 * Record one action. Never throws — a failure to log must not roll back the
 * work itself, and definitely must not mask the real error.
 */
function logAudit_(action, entity, entityId, field, oldValue, newValue, detail, success) {
  try {
    const t = TABLES.AUDIT_LOG, C = COL.AUDIT_LOG;
    const row = blankRow_('AUDIT_LOG');
    row[C.Log_ID]    = nextAuditId_();
    row[C.TS]        = new Date();
    row[C.Email]     = getActiveEmail_() || 'system';
    row[C.Action]    = action;
    row[C.Entity]    = entity || '';
    row[C.Entity_ID] = (entityId === undefined || entityId === null) ? '' : String(entityId);
    row[C.Field]     = field || '';
    row[C.Old_Value] = truncate_(oldValue);
    row[C.New_Value] = truncate_(newValue);
    row[C.Detail]    = truncate_(detail);
    row[C.Success]   = (success === undefined) ? true : !!success;
    getSheet_(t.sheet).appendRow(row);
    invalidateSheetCache_(t.sheet);
  } catch (err) {
    Logger.log('audit write failed (continuing): ' + err.message);
  }
}

/** One audit row per field that actually changed. */
function logFieldChanges_(entity, entityId, before, after, headers) {
  let n = 0;
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (/^(Created_TS|Created_By|Updated_TS|Updated_By)$/.test(h)) continue;
    const a = displayValue_(before ? before[i] : ''), b = displayValue_(after[i]);
    if (a === b) continue;
    logAudit_('UPDATE', entity, entityId, h, a, b, '', true);
    n++;
  }
  return n;
}

function displayValue_(v) {
  if (v === null || v === undefined) return '';
  if (isDate_(v)) return fmtDate(v);
  if (typeof v === 'number') return String(Math.round(v * 1000000) / 1000000);
  return String(v);
}

function truncate_(v) {
  const s = displayValue_(v);
  return s.length > 480 ? s.slice(0, 477) + '...' : s;
}


// ─────────────────────────────────────────────────────────────────────────────
// AMENDS — full-row history
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Snapshot a row into its history table.
 *
 * @param {string} tableKey   e.g. 'RATE_BASE'
 * @param {Array}  rowValues  the row as written, in source-table column order
 * @param {string} amendType  CREATE, UPDATE or DELETE
 */
function writeAmend_(tableKey, rowValues, amendType) {
  const amendKey = tableKey + '_AMENDS';
  const amendTable = TABLES[amendKey];
  if (!amendTable) return false;   // table has no history configured

  try {
    const srcHeaders = TABLES[tableKey].headers;
    const amendCols  = COL[amendKey];
    const out = blankRow_(amendKey);

    out[amendCols.Amend_ID]   = getNextId_(amendTable.sheet, amendCols.Amend_ID);
    out[amendCols.Amend_TS]   = new Date();
    out[amendCols.Amend_By]   = getActiveEmail_() || 'system';
    out[amendCols.Amend_Type] = amendType;

    // by name, so a new column in the source table lands correctly here too
    for (let i = 0; i < srcHeaders.length; i++) {
      const target = amendCols[srcHeaders[i]];
      if (target !== undefined) out[target] = rowValues[i];
    }

    getSheet_(amendTable.sheet).appendRow(out);
    invalidateSheetCache_(amendTable.sheet);
    return true;
  } catch (err) {
    Logger.log('amend write failed for ' + tableKey + ' (continuing): ' + err.message);
    return false;
  }
}


/**
 * Record a change completely: the snapshot, plus one audit row per field.
 * This is the single call every write path should make.
 */
function recordChange_(tableKey, entityId, beforeRow, afterRow, amendType) {
  writeAmend_(tableKey, afterRow || beforeRow, amendType);
  const headers = TABLES[tableKey].headers;
  const sheet = TABLES[tableKey].sheet;

  if (amendType === 'CREATE') {
    logAudit_('CREATE', sheet, entityId, '', '', summariseRow_(afterRow, headers), '', true);
    return 1;
  }
  if (amendType === 'DELETE') {
    logAudit_('DELETE', sheet, entityId, '', summariseRow_(beforeRow, headers), '', '', true);
    return 1;
  }
  return logFieldChanges_(sheet, entityId, beforeRow, afterRow, headers);
}

/** A short human description of a row, for create and delete entries. */
function summariseRow_(row, headers) {
  if (!row) return '';
  const parts = [];
  for (let i = 0; i < headers.length && parts.length < 6; i++) {
    const h = headers[i];
    if (/^(Created_|Updated_|Notes$|Source_Ref$)/.test(h)) continue;
    const v = displayValue_(row[i]);
    if (v !== '') parts.push(h + '=' + v);
  }
  return parts.join(', ');
}


/**
 * Record many changes to one table in a handful of writes instead of hundreds.
 *
 * recordChange_ appends one row to the amends table and one to Audit_Log per
 * record. That is fine for a single edit, but a 22-route mix grid became 44
 * separate appendRow calls and took 16 seconds. Collecting the rows and writing
 * each sheet once takes it to about one.
 *
 * @param {string} tableKey  e.g. 'MIX_METHOD'
 * @param {Array}  items     [{ id, before, after, type }]
 */
function recordChangesBatch_(tableKey, items) {
  if (!items || !items.length) return 0;

  const srcHeaders = TABLES[tableKey].headers;
  const sheetName  = TABLES[tableKey].sheet;
  const amendKey   = tableKey + '_AMENDS';
  const email      = getActiveEmail_() || 'system';
  const now        = new Date();

  // ---- amends -------------------------------------------------------------
  if (TABLES[amendKey]) {
    try {
      const amendCols = COL[amendKey];
      const amendSheet = getSheet_(TABLES[amendKey].sheet);
      let nextId = getNextId_(TABLES[amendKey].sheet, amendCols.Amend_ID);
      const rows = items.map(it => {
        const out = blankRow_(amendKey);
        out[amendCols.Amend_ID]   = nextId++;
        out[amendCols.Amend_TS]   = now;
        out[amendCols.Amend_By]   = email;
        out[amendCols.Amend_Type] = it.type;
        const src = it.after || it.before;
        for (let i = 0; i < srcHeaders.length; i++) {
          const target = amendCols[srcHeaders[i]];
          if (target !== undefined) out[target] = src[i];
        }
        return out;
      });
      amendSheet.getRange(amendSheet.getLastRow() + 1, 1, rows.length,
                          TABLES[amendKey].headers.length).setValues(rows);
      invalidateSheetCache_(TABLES[amendKey].sheet);
    } catch (err) {
      Logger.log('batch amend write failed (continuing): ' + err.message);
    }
  }

  // ---- audit --------------------------------------------------------------
  try {
    const C = COL.AUDIT_LOG;
    const auditSheet = getSheet_(SHEET.AUDIT_LOG);
    const rows = [];

    items.forEach(it => {
      if (it.type === 'UPDATE' && it.before) {
        for (let i = 0; i < srcHeaders.length; i++) {
          const h = srcHeaders[i];
          if (/^(Created_TS|Created_By|Updated_TS|Updated_By)$/.test(h)) continue;
          const a = displayValue_(it.before[i]), b = displayValue_(it.after[i]);
          if (a === b) continue;
          const row = blankRow_('AUDIT_LOG');
          row[C.TS] = now; row[C.Email] = email;
          row[C.Action] = 'UPDATE'; row[C.Entity] = sheetName;
          row[C.Entity_ID] = String(it.id); row[C.Field] = h;
          row[C.Old_Value] = truncate_(a); row[C.New_Value] = truncate_(b);
          row[C.Success] = true;
          rows.push(row);
        }
      } else {
        const row = blankRow_('AUDIT_LOG');
        row[C.TS] = now; row[C.Email] = email;
        row[C.Action] = it.type; row[C.Entity] = sheetName;
        row[C.Entity_ID] = String(it.id);
        if (it.type === 'DELETE') row[C.Old_Value] = summariseRow_(it.before, srcHeaders);
        else                      row[C.New_Value] = summariseRow_(it.after, srcHeaders);
        row[C.Success] = true;
        rows.push(row);
      }
    });

    if (rows.length) {
      // One reservation for the whole batch, numbered once the row count is
      // known, so a grid write costs a single property round-trip.
      let nextLog = reserveAuditIds_(rows.length);
      rows.forEach(r => { r[C.Log_ID] = nextLog++; });

      auditSheet.getRange(auditSheet.getLastRow() + 1, 1, rows.length,
                          TABLES.AUDIT_LOG.headers.length).setValues(rows);
      invalidateSheetCache_(SHEET.AUDIT_LOG);
    }
    return rows.length;
  } catch (err) {
    Logger.log('batch audit write failed (continuing): ' + err.message);
    return 0;
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// READING HISTORY
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Every change to one record, newest first.
 * @param {string} tableKey  e.g. 'RATE_BASE'
 * @param {number} recordId  the primary key value
 */
function getRecordHistory(tableKey, recordId) {
  const perms = requirePermissions_();
  requireViewAudit_(perms);
  return readRecordHistory_(tableKey, recordId);
}

/**
 * The same read, without the permission gate.
 *
 * Setup diagnostics need this before anyone is in the Permissions table —
 * otherwise you cannot test the audit trail until permissions work, and you
 * cannot confidently set up permissions until the audit trail is tested.
 * Never expose this to the client; getRecordHistory is the public door.
 */
function readRecordHistory_(tableKey, recordId) {
  const amendKey = tableKey + '_AMENDS';
  if (!TABLES[amendKey]) return [];

  const data = getAllData_(TABLES[amendKey].sheet);
  const C = COL[amendKey];
  const pkHeader = TABLES[tableKey].headers[0];
  const pkCol = C[pkHeader];
  const out = [];

  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][pkCol]) !== safeInt(recordId)) continue;
    const entry = {
      amendId: safeInt(data[i][C.Amend_ID]),
      ts: fmtTs(data[i][C.Amend_TS]),
      by: safeStr(data[i][C.Amend_By]),
      type: safeStr(data[i][C.Amend_Type]),
      values: {}
    };
    TABLES[tableKey].headers.forEach(h => {
      if (C[h] !== undefined) entry.values[h] = displayValue_(data[i][C[h]]);
    });
    out.push(entry);
  }
  return out.sort((a, b) => b.amendId - a.amendId);
}


/** The most recent activity across everything, newest first. */
function getRecentAudit(limit) {
  const perms = requirePermissions_();
  requireViewAudit_(perms);
  return readRecentAudit_(limit);
}

/** Ungated read, for setup diagnostics only. */
function readRecentAudit_(limit) {
  limit = limit || 100;
  const data = getAllData_(SHEET.AUDIT_LOG), C = COL.AUDIT_LOG;
  const out = [];
  for (let i = data.length - 1; i >= 1 && out.length < limit; i--) {
    if (!safeStr(data[i][C.Action])) continue;
    out.push({
      logId:    safeInt(data[i][C.Log_ID]),
      ts:       fmtTs(data[i][C.TS]),
      email:    safeStr(data[i][C.Email]),
      action:   safeStr(data[i][C.Action]),
      entity:   safeStr(data[i][C.Entity]),
      entityId: safeStr(data[i][C.Entity_ID]),
      field:    safeStr(data[i][C.Field]),
      oldValue: safeStr(data[i][C.Old_Value]),
      newValue: safeStr(data[i][C.New_Value]),
      detail:   safeStr(data[i][C.Detail]),
      success:  safeBool(data[i][C.Success])
    });
  }
  return out;
}


/**
 * Every Log_ID in the tab, in sheet order — so a diagnostic can say whether the
 * allocation actually is unique and ascending rather than assuming it.
 */
function auditLogIds_() {
  const data = getAllData_(SHEET.AUDIT_LOG), C = COL.AUDIT_LOG;
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = safeInt(data[i][C.Log_ID]);
    if (v) out.push(v);
  }
  return out;
}


/**
 * How many changes have been made since the forecast was last published.
 * Drives the "12 changes since last run" badge in the portal.
 */
function getPendingChangeCount_() {
  const runs = getAllData_(SHEET.CALC_RUNS), R = COL.CALC_RUNS;
  let lastRun = 0;
  for (let i = 1; i < runs.length; i++) {
    const ts = normaliseDate(runs[i][R.Run_TS]);
    const raw = runs[i][R.Run_TS];
    const t = isDate_(raw) ? raw.getTime()
            : (typeof raw === 'number') ? (raw - 25569) * 86400000
            : (ts ? ts.getTime() : 0);
    if (t > lastRun) lastRun = t;
  }

  const audit = getAllData_(SHEET.AUDIT_LOG), C = COL.AUDIT_LOG;
  let count = 0;
  for (let i = 1; i < audit.length; i++) {
    const action = safeStr(audit[i][C.Action]);
    if (action !== 'CREATE' && action !== 'UPDATE' && action !== 'DELETE') continue;
    if (!safeBool(audit[i][C.Success])) continue;
    const raw = audit[i][C.TS];
    const t = isDate_(raw) ? raw.getTime()
            : (typeof raw === 'number') ? (raw - 25569) * 86400000 : 0;
    if (t > lastRun) count++;
  }
  return count;
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — run from the editor
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Writes one harmless test entry to Audit_Log and one to Rate_Base_Amends,
 * reads them back, then leaves them in place as proof the trail works.
 */
function testAuditTrail() {
  requireMaintenance_();
  Logger.log('=== AUDIT TRAIL TEST ===');
  const email = getActiveEmail_();
  Logger.log('  acting as: ' + (email || '(unknown)'));

  const beforeAudit  = Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0);
  const beforeAmends = Math.max(getSheet_(SHEET.RATE_BASE_AMENDS).getLastRow() - 1, 0);
  Logger.log('  Audit_Log rows before        : ' + beforeAudit);
  Logger.log('  Rate_Base_Amends rows before : ' + beforeAmends);

  // A realistic-looking but clearly marked snapshot. Nothing real is touched.
  const C = COL.RATE_BASE;
  const row = blankRow_('RATE_BASE');
  row[C.Rate_ID]      = 999999;
  row[C.Modelling_ID] = 0;
  row[C.Valid_From]   = normaliseDate('2026-01-01');
  row[C.Valid_To]     = normaliseDate('2026-12-31');
  row[C.Base_Rate]    = 1.23;
  row[C.Currency]     = 'GBP';
  row[C.Scenario_ID]  = 1;
  row[C.Notes]        = 'TEST ROW from testAuditTrail — safe to delete';
  row[C.Active]       = false;

  invalidateSheetCache_(SHEET.AUDIT_LOG);
  invalidateSheetCache_(SHEET.RATE_BASE_AMENDS);
  recordChange_('RATE_BASE', 999999, null, row, 'CREATE');

  const afterAudit  = Math.max(getSheet_(SHEET.AUDIT_LOG).getLastRow() - 1, 0);
  const afterAmends = Math.max(getSheet_(SHEET.RATE_BASE_AMENDS).getLastRow() - 1, 0);
  Logger.log('  Audit_Log rows after         : ' + afterAudit  + '  (+' + (afterAudit  - beforeAudit)  + ')');
  Logger.log('  Rate_Base_Amends rows after  : ' + afterAmends + '  (+' + (afterAmends - beforeAmends) + ')');

  Logger.log('');
  Logger.log('--- reading it back ---');
  invalidateSheetCache_(SHEET.RATE_BASE_AMENDS);
  const hist = readRecordHistory_('RATE_BASE', 999999);
  if (!hist.length) {
    Logger.log('  NOTHING FOUND — the amends write did not land');
  } else {
    hist.slice(0, 3).forEach(h => {
      Logger.log('  ' + h.ts + '  ' + h.type + '  by ' + h.by);
      Logger.log('    Base_Rate=' + h.values.Base_Rate + '  Valid_From=' + h.values.Valid_From +
                 '  Valid_To=' + h.values.Valid_To);
    });
  }

  Logger.log('');
  invalidateSheetCache_(SHEET.AUDIT_LOG);
  const recent = readRecentAudit_(3);
  Logger.log('--- three most recent actions ---');
  recent.forEach(a => Logger.log('  ' + a.ts + '  ' + a.email + '  ' + a.action + ' ' +
                                a.entity + ' ' + a.entityId + '  ' + (a.newValue || a.detail)));

  Logger.log('');
  Logger.log('--- Log_ID allocation ---');
  // Judged on the rows this run just wrote, not on the whole tab: any duplicate
  // that predates this run came from the old max + 1 allocator, and failing the
  // test for it would send you looking for a bug that has already been fixed.
  const ids = auditLogIds_();
  const added = Math.max(afterAudit - beforeAudit, 0);
  const priorIds = {}, dupes = {};
  ids.slice(0, ids.length - added).forEach(v => { priorIds[v] = true; });

  const newIds = ids.slice(ids.length - added);
  let ascending = true;
  newIds.forEach((v, i) => {
    if (priorIds[v]) dupes[v] = true;
    priorIds[v] = true;
    if (i > 0 && v <= newIds[i - 1]) ascending = false;
  });
  const dupeList = Object.keys(dupes);

  Logger.log('  rows with an ID   : ' + ids.length);
  Logger.log('  written this run  : ' + added + (newIds.length ? ' (' + newIds.join(', ') + ')' : ''));
  Logger.log('  counter now at    : ' +
             (PropertiesService.getScriptProperties()
                .getProperty('AUDIT_LOG_NEXT_ID_' + spreadsheetId_()) ||
              '(unset — seeds from the sheet on first write)'));
  Logger.log('  reused an old ID  : ' + (dupeList.length ? 'YES — ' + dupeList.join(', ') : 'no'));
  Logger.log('  ascending         : ' + (ascending ? 'yes' : 'NO — IDs go backwards'));
  const idsOk = added > 0 && !dupeList.length && ascending;

  Logger.log('');
  const perms = getUserPermissions_();
  if (!perms.found) {
    Logger.log('--- note ---');
    Logger.log('  You are not yet in the Permissions tab, so the portal itself would');
    Logger.log('  refuse you. That does not affect this test, which checks the writing');
    Logger.log('  mechanism. Add your row before Step 7 — see Part D.');
    Logger.log('');
  }
  const writesOk = (afterAudit > beforeAudit) && (afterAmends > beforeAmends) && hist.length > 0;
  const ok = writesOk && idsOk;
  Logger.log(ok ? 'AUDIT TRAIL WORKING'
                : (writesOk
                    ? 'PROBLEM — Log_ID allocation is not unique and ascending. Send me this log.'
                    : 'PROBLEM — one of the writes did not land. Send me this log.'));
  Logger.log('');
  Logger.log('Two test rows were left behind on purpose so you can look at them:');
  Logger.log('  Audit_Log        — a CREATE entry for Rate_Base 999999');
  Logger.log('  Rate_Base_Amends — the matching snapshot');
  Logger.log('Delete them whenever you like. Nothing real was changed.');
  return { ok: ok };
}