/**
 * Postage Forecast Portal — enginetest.gs
 *
 * The go-live gate for Step 3.
 *
 *   runParityTest()      compares the engine against the Output tab of the
 *                        original workbook, row by row
 *   runEngineUnitTests() checks the resolvers in isolation — leap years,
 *                        open-ended dates, mid-month changes, boundaries
 *
 * Both are read only.
 */

/** How close two numbers must be to count as equal. */
const PARITY_TOLERANCE = 0.0001;


// ─────────────────────────────────────────────────────────────────────────────
// PARITY — does the engine reproduce the workbook?
// ─────────────────────────────────────────────────────────────────────────────

function runParityTest() {
  requireMaintenance_();
  const t0 = Date.now();
  Logger.log('=== PARITY TEST ===');
  Logger.log('Comparing the engine against the Output tab of the original workbook.');
  Logger.log('');

  // ---- what the engine says ----------------------------------------------
  const input  = loadEngineInput_(1);
  const result = computeModel_(input);
  const mine = {};
  result.outputRows.forEach(r => { mine[r.hlId + '|' + r.dateId] = r.forecastRatePerOrder; });
  Logger.log('Engine produced ' + result.outputRows.length + ' rows in ' + result.stats.ms + 'ms');

  // ---- what the workbook says --------------------------------------------
  //  Output tab layout: 0 High Level ID, 1 Date ID, 2 Brand, 3 Geo,
  //                     4 Treatment Type, 5 WL Split, 6 Date, 7 Total Postage
  let src;
  try {
    src = readSourceTab_('Output');
  } catch (e) {
    Logger.log('Could not read the Output tab from the source workbook: ' + e.message);
    Logger.log('Check SOURCE_SPREADSHEET_ID in migrate.gs.');
    return { ok: false };
  }

  const theirs = {};
  let srcRows = 0;
  for (let i = 1; i < src.length; i++) {
    const r = src[i] || [];
    const hlId = safeInt(r[0]), dateId = safeInt(r[1]);
    if (!hlId || !dateId) continue;
    theirs[hlId + '|' + dateId] = safeNum(r[7]);
    srcRows++;
  }
  Logger.log('Workbook Output tab has ' + srcRows + ' rows');
  Logger.log('');

  // ---- compare ------------------------------------------------------------
  const diffs = [], missingFromEngine = [], missingFromWorkbook = [];
  let matched = 0, maxDiff = 0, maxDiffKey = '';

  for (const key in theirs) {
    if (!(key in mine)) { missingFromEngine.push(key); continue; }
    const a = theirs[key], b = mine[key], d = Math.abs(a - b);
    if (d > maxDiff) { maxDiff = d; maxDiffKey = key; }
    if (d <= PARITY_TOLERANCE) matched++;
    else diffs.push({ key: key, workbook: a, engine: b, diff: b - a,
                      pct: a ? ((b - a) / a * 100) : 0 });
  }
  for (const key in mine) if (!(key in theirs)) missingFromWorkbook.push(key);

  // ---- report -------------------------------------------------------------
  Logger.log('--- result ---');
  Logger.log('  matched within ' + PARITY_TOLERANCE + ' : ' + matched + ' of ' + srcRows);
  Logger.log('  differences                : ' + diffs.length);
  Logger.log('  in workbook but not engine : ' + missingFromEngine.length);
  Logger.log('  in engine but not workbook : ' + missingFromWorkbook.length);
  Logger.log('  largest difference         : ' + maxDiff.toExponential(3) +
             (maxDiffKey ? '  at High Level ID ' + maxDiffKey.split('|')[0] +
                           ', Date ID ' + maxDiffKey.split('|')[1] : ''));

  if (diffs.length) {
    Logger.log('');
    Logger.log('--- rows that differ (first 20) ---');
    Logger.log('  HLID  DateID    workbook       engine         diff       %');
    diffs.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 20).forEach(d => {
      const p = d.key.split('|');
      Logger.log('  ' + pad_(p[0], 4) + '  ' + pad_(p[1], 6) + '  ' +
                 pad_(d.workbook.toFixed(6), 12) + ' ' + pad_(d.engine.toFixed(6), 12) + ' ' +
                 pad_(d.diff.toFixed(6), 11) + ' ' + d.pct.toFixed(2) + '%');
    });
    const byHl = {};
    diffs.forEach(d => { const h = d.key.split('|')[0]; byHl[h] = (byHl[h] || 0) + 1; });
    Logger.log('');
    Logger.log('  differences by High Level ID:');
    for (const h in byHl) Logger.log('    HLID ' + h + ': ' + byHl[h] + ' month(s)');
  }

  if (missingFromEngine.length) {
    Logger.log('');
    Logger.log('--- in the workbook but not produced by the engine (first 10) ---');
    missingFromEngine.slice(0, 10).forEach(k =>
      Logger.log('    High Level ID ' + k.split('|')[0] + ', Date ID ' + k.split('|')[1]));
  }

  // ---- engine defects -----------------------------------------------------
  if (result.defects.length) {
    Logger.log('');
    Logger.log('--- engine defects (these do not affect parity, but read them) ---');
    const byRule = {};
    result.defects.forEach(d => { byRule[d.rule] = (byRule[d.rule] || 0) + 1; });
    for (const rule in byRule) Logger.log('  ' + rule + ': ' + byRule[rule]);
    result.defects.slice(0, 5).forEach(d => Logger.log('    ' + d.label));
  }

  // ---- verdict ------------------------------------------------------------
  const ok = (diffs.length === 0 && missingFromEngine.length === 0 && srcRows > 0);
  Logger.log('');
  if (ok) {
    Logger.log('PARITY PASS — ' + matched + ' of ' + srcRows + ' rows match.');
    Logger.log('The engine reproduces the workbook exactly. Ready for step 4.');
  } else {
    Logger.log('PARITY FAIL — ' + diffs.length + ' row(s) differ.');
    Logger.log('Do not proceed to step 4. Send me this log.');
  }
  Logger.log('');
  Logger.log('(' + (Date.now() - t0) + 'ms)');
  return { ok: ok, matched: matched, total: srcRows, diffs: diffs.slice(0, 50), maxDiff: maxDiff };
}



// ─────────────────────────────────────────────────────────────────────────────
// UNIT TESTS — the resolvers on their own
// ─────────────────────────────────────────────────────────────────────────────

function runEngineUnitTests() {
  requireMaintenance_();
  Logger.log('=== ENGINE UNIT TESTS ===');
  let pass = 0, fail = 0;

  function check(name, actual, expected, tol) {
    tol = (tol === undefined) ? 1e-9 : tol;
    const ok = (typeof expected === 'number')
      ? Math.abs(actual - expected) <= tol
      : actual === expected;
    Logger.log('  ' + (ok ? 'PASS  ' : 'FAIL  ') + name +
               (ok ? '' : '   got ' + actual + ', expected ' + expected));
    ok ? pass++ : fail++;
  }

  const D = s => dayNum(s);

  // ---- pointInTime --------------------------------------------------------
  Logger.log('-- pointInTime (base rates and mixes) --');
  const pit = [{ f: D('2026-01-01'), t: D('2026-03-31'), v: 3.60 },
               { f: D('2026-04-01'), t: D('9999-12-31'), v: 3.78 }];
  check('March uses the first period',  pointInTime_(pit, D('2026-03-01'), null, null), 3.60);
  check('April uses the second',        pointInTime_(pit, D('2026-04-01'), null, null), 3.78);
  check('a change on the 15th does not apply until next month',
        pointInTime_([{ f: D('2026-05-15'), t: D('9999-12-31'), v: 9 }], D('2026-05-01'), null, null), 0);
  check('open-ended still applies years later',
        pointInTime_(pit, D('2028-12-01'), null, null), 3.78);
  check('no matching row returns zero',
        pointInTime_(pit, D('2025-01-01'), null, null), 0);

  const defects = [];
  pointInTime_([{ f: D('2026-01-01'), t: D('2026-12-31'), v: 1 },
                { f: D('2026-06-01'), t: D('2026-08-31'), v: 1 }], D('2026-07-01'), defects, 'test');
  check('overlapping rows are recorded as a defect', defects.length, 1);

  // ---- dayWeighted --------------------------------------------------------
  Logger.log('-- dayWeighted (surcharges) --');
  const may = { ms: D('2026-05-01'), me: D('2026-05-31'), dim: 31 };
  const fuel = [{ f: D('2026-01-01'), t: D('2026-05-03'), v: 0.11 },
                { f: D('2026-05-04'), t: D('2026-09-30'), v: 0.14 }];
  check('May blends 3 days at 11% and 28 at 14%',
        dayWeighted_(fuel, may.ms, may.me, may.dim), 0.137097, 0.000001);
  check('April is wholly the first rate',
        dayWeighted_(fuel, D('2026-04-01'), D('2026-04-30'), 30), 0.11);
  check('June is wholly the second',
        dayWeighted_(fuel, D('2026-06-01'), D('2026-06-30'), 30), 0.14);
  check('a period ending on the 1st contributes only 1 day of 31',
        dayWeighted_([{ f: D('2026-01-01'), t: D('2026-12-01'), v: 0.31 }],
                     D('2026-12-01'), D('2026-12-31'), 31), 0.01);
  check('a period ending at month end contributes all of it',
        dayWeighted_([{ f: D('2026-01-01'), t: D('2026-12-31'), v: 0.31 }],
                     D('2026-12-01'), D('2026-12-31'), 31), 0.31);
  check('no overlap returns zero',
        dayWeighted_([{ f: D('2027-01-01'), t: D('2027-12-31'), v: 5 }],
                     may.ms, may.me, may.dim), 0);

  // ---- leap years ---------------------------------------------------------
  Logger.log('-- leap years --');
  check('Feb 2028 has 29 days', D('2028-03-01') - D('2028-02-01'), 29);
  check('Feb 2026 has 28 days', D('2026-03-01') - D('2026-02-01'), 28);
  check('a full Feb 2028 at 10% averages 10%',
        dayWeighted_([{ f: D('2028-02-01'), t: D('2028-02-29'), v: 0.10 }],
                     D('2028-02-01'), D('2028-02-29'), 29), 0.10);

  // ---- boundaries ---------------------------------------------------------
  Logger.log('-- boundaries --');
  check('a single-day period on the 1st', dayWeighted_(
        [{ f: D('2026-01-01'), t: D('2026-01-01'), v: 31 }], D('2026-01-01'), D('2026-01-31'), 31), 1);
  check('a period starting before the month is clipped', dayWeighted_(
        [{ f: D('2025-06-01'), t: D('2026-01-10'), v: 31 }], D('2026-01-01'), D('2026-01-31'), 31), 10);
  check('contiguous periods total the whole month', dayWeighted_(
        [{ f: D('2026-01-01'), t: D('2026-01-10'), v: 31 },
         { f: D('2026-01-11'), t: D('2026-01-31'), v: 31 }],
        D('2026-01-01'), D('2026-01-31'), 31), 31);

  // ---- the rate formula ---------------------------------------------------
  Logger.log('-- rate assembly --');
  check('base x (1 + fuel) + fixed', 2.53 * (1 + 0.137097) + 0.10, 2.976855, 0.000001);
  check('fixed amounts are not inflated by fuel', 2.53 * (1 + 0.137097) + 0.10,
        2.53 + 2.53 * 0.137097 + 0.10, 1e-9);

  Logger.log('');
  Logger.log(fail ? (fail + ' FAILURES, ' + pass + ' passed') : ('ALL ' + pass + ' TESTS PASSED'));
  return { pass: pass, fail: fail };
}