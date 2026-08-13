/**
 * Postage Forecast Portal — snapshots.gs
 *
 * A record of what the forecast said, week by week, so a change can be traced
 * back to the rate that caused it.
 *
 * THE DESIGN CHOICE WORTH KNOWING
 *
 * The job runs weekly, but a snapshot is only STORED when the forecast actually
 * differs from the last one. Storing 52 a year regardless costs almost nothing
 * in space — about 159,000 cells against a 10,000,000 limit — but it makes the
 * list useless: 32 identical entries between two real ones is noise, and
 * answering "when did this move?" gets harder rather than easier.
 *
 * Every run is logged either way, so you can always see the check happened.
 *
 * The Audit_Log already records that a rate changed. Snapshots record what that
 * did to the numbers, which is a different question.
 */

const SNAPSHOT_MIN_GAP_MINUTES = 55;   // a light rate limit, see takeSnapshotInternal_


// ─────────────────────────────────────────────────────────────────────────────
// TAKING A SNAPSHOT
// ─────────────────────────────────────────────────────────────────────────────

/** Manual snapshot, from the portal or the editor. */
function takeForecastSnapshot(label) {
  const perms = requirePermissions_();
  requireRunCalc_(perms);
  return takeSnapshotInternal_('MANUAL', perms.email, safeStr(label));
}

/**
 * The weekly job. This is what the time-driven trigger calls.
 *
 * Deliberately not permission-gated. In a trigger there is no signed-in user in
 * the usual sense, and the worst a stray call can do is store a copy of data
 * that is already published — so a rate limit is the proportionate guard rather
 * than an auth check that would break the schedule.
 */
function weeklySnapshotJob() {
  try {
    const r = takeSnapshotInternal_('WEEKLY', 'scheduled', '');
    Logger.log(r.stored ? ('Snapshot ' + r.snapshotId + ' stored, ' + r.changedRows + ' rows changed.')
                        : ('No change since snapshot ' + r.comparedTo + ' — nothing stored.'));
    return r;
  } catch (e) {
    Logger.log('Weekly snapshot failed: ' + e.message);
    logAudit_('CALC', 'SNAPSHOT', '', '', '', '', 'weekly job failed: ' + e.message, false);
    return { ok: false, error: e.message };
  }
}


function takeSnapshotInternal_(trigger, actor, label) {
  prewarmSheetCache_([SHEET.OUTPUT, SHEET.SNAPSHOTS, SHEET.SNAPSHOT_VALUES,
                      SHEET.CALC_RUNS, SHEET.CONFIG]);

  // ---- what the forecast says now ---------------------------------------
  const cur = {};
  const rows = [];
  const out = getAllData_(SHEET.OUTPUT), O = COL.OUTPUT;
  let calcRunId = 0;
  for (let i = 1; i < out.length; i++) {
    const hl = safeInt(out[i][O.High_Level_ID]), did = safeInt(out[i][O.Date_ID]);
    if (!hl || !did) continue;
    const rate = safeNum(out[i][O.Forecast_Rate_Per_Order]);
    cur[hl + '|' + did] = rate;
    rows.push({ hlId: hl, dateId: did,
                monthStart: normaliseDate(out[i][O.Month_Start]), rate: rate });
    calcRunId = Math.max(calcRunId, safeInt(out[i][O.Calc_Run_ID]));
  }
  if (!rows.length) {
    return { ok: false, stored: false, reason: 'OUTPUT is empty — publish the forecast first.' };
  }

  // ---- the most recent snapshot ------------------------------------------
  const snaps = getAllData_(SHEET.SNAPSHOTS), S = COL.SNAPSHOTS;
  let lastId = 0, lastTs = 0;
  for (let i = 1; i < snaps.length; i++) {
    const id = safeInt(snaps[i][S.Snapshot_ID]);
    if (!id) continue;
    if (id > lastId) {
      lastId = id;
      const raw = snaps[i][S.Taken_TS];
      lastTs = isDate_(raw) ? raw.getTime()
             : (typeof raw === 'number' ? Math.round((raw - SERIAL_EPOCH_OFFSET) * MS_PER_DAY) : 0);
    }
  }

  // A rate limit rather than an auth check: repeated calls cannot bloat the file.
  if (lastTs && (Date.now() - lastTs) < SNAPSHOT_MIN_GAP_MINUTES * 60000) {
    return { ok: true, stored: false, comparedTo: lastId,
             reason: 'A snapshot was taken less than an hour ago.' };
  }

  // ---- compare -----------------------------------------------------------
  let changed = 0, prevCount = 0;
  if (lastId) {
    const vals = getAllData_(SHEET.SNAPSHOT_VALUES), V = COL.SNAPSHOT_VALUES;
    const prev = {};
    for (let i = 1; i < vals.length; i++) {
      if (safeInt(vals[i][V.Snapshot_ID]) !== lastId) continue;
      prev[safeInt(vals[i][V.High_Level_ID]) + '|' + safeInt(vals[i][V.Date_ID])] =
        safeNum(vals[i][V.Forecast_Rate_Per_Order]);
      prevCount++;
    }
    if (prevCount) {
      for (const k in cur) {
        if (prev[k] === undefined || Math.abs(prev[k] - cur[k]) > 0.0000005) changed++;
      }
      for (const k in prev) if (cur[k] === undefined) changed++;
      if (!changed) {
        logAudit_('CALC', 'SNAPSHOT', '', '', '', 'no change',
                  trigger + ': identical to snapshot ' + lastId, true);
        return { ok: true, stored: false, comparedTo: lastId, changedRows: 0,
                 reason: 'The forecast has not moved since snapshot ' + lastId + '.' };
      }
    } else {
      changed = rows.length;   // previous snapshot has no values recorded
    }
  } else {
    changed = rows.length;     // the first one
  }

  // ---- store -------------------------------------------------------------
  return withLock_(function () {
    const st = TABLES.SNAPSHOTS, SC = COL.SNAPSHOTS;
    const snapshotId = getNextId_(st.sheet, SC.Snapshot_ID);

    const vt = TABLES.SNAPSHOT_VALUES, VC = COL.SNAPSHOT_VALUES;
    const vSheet = getSheet_(vt.sheet);
    const vRows = rows.map(function (r) {
      const row = blankRow_('SNAPSHOT_VALUES');
      row[VC.Snapshot_ID]              = snapshotId;
      row[VC.High_Level_ID]            = r.hlId;
      row[VC.Date_ID]                  = r.dateId;
      row[VC.Month_Start]              = r.monthStart;
      row[VC.Forecast_Rate_Per_Order]  = r.rate;
      return row;
    });
    vSheet.getRange(vSheet.getLastRow() + 1, 1, vRows.length,
                    vt.headers.length).setValues(vRows);
    invalidateSheetCache_(vt.sheet);

    const sSheet = getSheet_(st.sheet);
    const srow = blankRow_('SNAPSHOTS');
    srow[SC.Snapshot_ID]  = snapshotId;
    srow[SC.Taken_TS]     = new Date();
    srow[SC.Taken_By]     = actor;
    srow[SC.Trigger]      = trigger;
    srow[SC.Scenario_ID]  = 1;
    srow[SC.Label]        = label || (trigger === 'WEEKLY' ? 'Weekly' : 'Manual');
    srow[SC.Rows_Stored]  = vRows.length;
    srow[SC.Changed_Rows] = changed;
    srow[SC.Calc_Run_ID]  = calcRunId;
    srow[SC.Notes]        = lastId ? ('compared against snapshot ' + lastId)
                                   : 'first snapshot';
    sSheet.appendRow(srow);
    invalidateSheetCache_(st.sheet);

    logAudit_('CALC', 'SNAPSHOT', snapshotId, '', '', vRows.length + ' rows',
              trigger + ': ' + changed + ' changed vs snapshot ' + (lastId || 'none'), true);

    return { ok: true, stored: true, snapshotId: snapshotId, rows: vRows.length,
             changedRows: changed, comparedTo: lastId };
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// READING
// ─────────────────────────────────────────────────────────────────────────────

function listSnapshots() {
  const perms = requirePermissions_();
  prewarmSheetCache_([SHEET.SNAPSHOTS, SHEET.PERMISSIONS, SHEET.PORTAL_ROLES,
                      SHEET.SCOPE_MAPPING, SHEET.CONFIG]);
  const data = getAllData_(SHEET.SNAPSHOTS), C = COL.SNAPSHOTS, out = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.Snapshot_ID]);
    if (!id) continue;
    out.push({ id: id, ts: fmtTs(data[i][C.Taken_TS]),
               by: safeStr(data[i][C.Taken_By]),
               trigger: safeStr(data[i][C.Trigger]),
               label: safeStr(data[i][C.Label]),
               rows: safeInt(data[i][C.Rows_Stored]),
               changed: safeInt(data[i][C.Changed_Rows]),
               calcRunId: safeInt(data[i][C.Calc_Run_ID]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out.sort(function (a, b) { return b.id - a.id; });
}

function loadSnapshotsForClient_() {
  const data = getAllData_(SHEET.SNAPSHOTS), C = COL.SNAPSHOTS, out = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.Snapshot_ID]);
    if (!id) continue;
    out.push({ id: id, ts: fmtTs(data[i][C.Taken_TS]), by: safeStr(data[i][C.Taken_By]),
               trigger: safeStr(data[i][C.Trigger]), label: safeStr(data[i][C.Label]),
               rows: safeInt(data[i][C.Rows_Stored]),
               changed: safeInt(data[i][C.Changed_Rows]) });
  }
  return out.sort(function (a, b) { return b.id - a.id; });
}


/**
 * What moved between two snapshots.
 *
 * @param {number} fromId  the earlier one
 * @param {number} toId    the later one, or 0 for the live OUTPUT
 */
function compareSnapshots(fromId, toId) {
  const perms = requirePermissions_();
  prewarmSheetCache_([SHEET.SNAPSHOT_VALUES, SHEET.SNAPSHOTS, SHEET.OUTPUT,
                      SHEET.DIM_CALENDAR, SHEET.HIGH_LEVEL_IDS, SHEET.PERMISSIONS,
                      SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING, SHEET.CONFIG]);
  const visible = visibleHighLevelIds_(perms);

  const label = {};
  const cal = getAllData_(SHEET.DIM_CALENDAR), K = COL.DIM_CALENDAR;
  for (let i = 1; i < cal.length; i++) {
    const ms = normaliseDate(cal[i][K.Month_Start]);
    if (ms) label[safeInt(cal[i][K.Date_ID])] = monthLabel_(cal[i][K.Month_Label], ms);
  }

  function valuesOf(id) {
    const map = {};
    if (!safeInt(id)) {
      const out = getAllData_(SHEET.OUTPUT), O = COL.OUTPUT;
      for (let i = 1; i < out.length; i++) {
        const hl = safeInt(out[i][O.High_Level_ID]), d = safeInt(out[i][O.Date_ID]);
        if (!hl || !d || !visible[hl]) continue;
        map[hl + '|' + d] = safeNum(out[i][O.Forecast_Rate_Per_Order]);
      }
      return map;
    }
    const vals = getAllData_(SHEET.SNAPSHOT_VALUES), V = COL.SNAPSHOT_VALUES;
    for (let i = 1; i < vals.length; i++) {
      if (safeInt(vals[i][V.Snapshot_ID]) !== safeInt(id)) continue;
      const hl = safeInt(vals[i][V.High_Level_ID]);
      if (!visible[hl]) continue;
      map[hl + '|' + safeInt(vals[i][V.Date_ID])] = safeNum(vals[i][V.Forecast_Rate_Per_Order]);
    }
    return map;
  }

  const a = valuesOf(fromId), b = valuesOf(toId);
  const diffs = [];
  let same = 0;

  for (const k in b) {
    const parts = k.split('|');
    const before = a[k], after = b[k];
    if (before === undefined) {
      diffs.push({ hlId: safeInt(parts[0]), dateId: safeInt(parts[1]),
                   month: label[safeInt(parts[1])] || parts[1],
                   before: null, after: after, diff: null, pct: null });
      continue;
    }
    if (Math.abs(before - after) <= 0.0000005) { same++; continue; }
    diffs.push({ hlId: safeInt(parts[0]), dateId: safeInt(parts[1]),
                 month: label[safeInt(parts[1])] || parts[1],
                 before: before, after: after, diff: after - before,
                 pct: before ? (after - before) / before * 100 : null });
  }
  for (const k in a) {
    if (b[k] !== undefined) continue;
    const parts = k.split('|');
    diffs.push({ hlId: safeInt(parts[0]), dateId: safeInt(parts[1]),
                 month: label[safeInt(parts[1])] || parts[1],
                 before: a[k], after: null, diff: null, pct: null });
  }

  diffs.sort(function (x, y) {
    return Math.abs(y.diff || 0) - Math.abs(x.diff || 0) || (x.hlId - y.hlId);
  });

  // which segments moved, and by how much on average
  const bySeg = {};
  diffs.forEach(function (d) {
    if (d.diff === null) return;
    const s = bySeg[d.hlId] || (bySeg[d.hlId] = { months: 0, total: 0 });
    s.months++; s.total += d.diff;
  });
  const segments = Object.keys(bySeg).map(function (id) {
    return { hlId: safeInt(id), months: bySeg[id].months,
             meanDiff: bySeg[id].total / bySeg[id].months };
  }).sort(function (x, y) { return Math.abs(y.meanDiff) - Math.abs(x.meanDiff); });

  return { fromId: safeInt(fromId), toId: safeInt(toId),
           unchanged: same, changed: diffs.length,
           diffs: diffs.slice(0, 300), segments: segments };
}


// ─────────────────────────────────────────────────────────────────────────────
// THE WEEKLY TRIGGER
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Install the weekly job. Run once, from the editor.
 *
 * Monday morning is chosen so a snapshot exists before anyone starts changing
 * rates for the week.
 */
function installWeeklySnapshot() {
  requireMaintenance_();
  Logger.log('=== WEEKLY SNAPSHOT TRIGGER ===');

  let existing = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklySnapshotJob') {
      ScriptApp.deleteTrigger(t);
      existing++;
    }
  });
  if (existing) Logger.log('  removed ' + existing + ' existing trigger(s) first');

  ScriptApp.newTrigger('weeklySnapshotJob')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  Logger.log('  installed: every Monday between 6am and 7am, ' + DEFAULT_TZ);
  Logger.log('');
  Logger.log('  It compares the live OUTPUT against the last snapshot and only stores');
  Logger.log('  one when something has actually moved, so the list stays readable.');
  Logger.log('');
  Logger.log('  To check it later: showSnapshotTriggers(). To stop it: removeWeeklySnapshot().');
  logAudit_('CREATE', 'TRIGGER', 'weeklySnapshotJob', '', '', 'Monday 6am', '', true);
  return { ok: true, replaced: existing };
}

function removeWeeklySnapshot() {
  requireMaintenance_();
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklySnapshotJob') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log(n ? ('Removed ' + n + ' trigger(s).') : 'There was no weekly trigger installed.');
  if (n) logAudit_('DELETE', 'TRIGGER', 'weeklySnapshotJob', '', '', '', '', true);
  return { removed: n };
}

function showSnapshotTriggers() {
  requireMaintenance_();
  Logger.log('=== TRIGGERS ON THIS PROJECT ===');
  const list = ScriptApp.getProjectTriggers();
  if (!list.length) {
    Logger.log('  none installed');
    Logger.log('  Run installWeeklySnapshot() to schedule the weekly snapshot.');
    return { count: 0 };
  }
  list.forEach(function (t) {
    Logger.log('  ' + t.getHandlerFunction() + '  (' + t.getEventType() + ')');
  });
  const weekly = list.filter(function (t) {
    return t.getHandlerFunction() === 'weeklySnapshotJob'; }).length;
  Logger.log('');
  Logger.log(weekly ? '  The weekly snapshot is installed.'
                    : '  The weekly snapshot is NOT installed.');
  return { count: list.length, weekly: weekly };
}


/**
 * Thin out old snapshots: keep everything from the last 90 days, then one per
 * month before that.
 *
 * Not scheduled — run it if the file ever feels heavy. At the rate snapshots are
 * actually stored, that is unlikely to be soon.
 */
function pruneSnapshots(keepDays) {
  requireMaintenance_();
  keepDays = safeInt(keepDays) || 90;
  Logger.log('=== PRUNING SNAPSHOTS ===');
  Logger.log('  keeping everything from the last ' + keepDays + ' days, then one per month');

  const data = getAllData_(SHEET.SNAPSHOTS), C = COL.SNAPSHOTS;
  const cutoff = Date.now() - keepDays * 86400000;
  const monthKeep = {}, drop = [];

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.Snapshot_ID]);
    if (!id) continue;
    const raw = data[i][C.Taken_TS];
    const ts = isDate_(raw) ? raw.getTime()
             : (typeof raw === 'number' ? Math.round((raw - SERIAL_EPOCH_OFFSET) * MS_PER_DAY) : 0);
    list.push({ id: id, ts: ts, row: i + 1 });
  }
  list.sort(function (a, b) { return b.ts - a.ts; });

  list.forEach(function (s) {
    if (s.ts >= cutoff) return;                       // recent, keep
    const d = new Date(s.ts);
    const key = d.getFullYear() + '-' + (d.getMonth() + 1);
    if (!monthKeep[key]) { monthKeep[key] = s.id; return; }   // newest in that month
    drop.push(s.id);
  });

  if (!drop.length) { Logger.log('  nothing to prune'); return { dropped: 0 }; }

  const vt = TABLES.SNAPSHOT_VALUES, VC = COL.SNAPSHOT_VALUES;
  const sh = getSheet_(vt.sheet);
  const last = sh.getLastRow();
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, vt.headers.length).getValues();
    const dropSet = {};
    drop.forEach(function (id) { dropSet[id] = true; });
    const keep = vals.filter(function (r) { return !dropSet[safeInt(r[VC.Snapshot_ID])]; });
    clearDataRows_(vt.sheet);
    if (keep.length) sh.getRange(2, 1, keep.length, vt.headers.length).setValues(keep);
    invalidateSheetCache_(vt.sheet);
    Logger.log('  Snapshot_Values: ' + (vals.length - keep.length) + ' rows removed');
  }

  // mark the header rows rather than deleting, so the history of runs survives
  const st = TABLES.SNAPSHOTS, sSheet = getSheet_(st.sheet);
  drop.forEach(function (id) {
    const idx = findRowById_(st.sheet, C.Snapshot_ID, id);
    if (idx > 0) sSheet.getRange(idx, C.Notes + 1).setValue('values pruned');
  });
  invalidateSheetCache_(st.sheet);

  Logger.log('  ' + drop.length + ' snapshot(s) pruned: ' + drop.join(', '));
  logAudit_('DELETE', 'SNAPSHOT', '', '', '', drop.length + ' pruned', '', true);
  return { dropped: drop.length };
}