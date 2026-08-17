# Postage Forecast Portal — measured baseline

**Status: partially captured.** `diagnoseBatchGet()` is done for both environments and is what
established that they are not comparable. `testApiPayload()` and the browser timings are still
blank — they need the Apps Script editor and a browser, so they are a step for a human, not
something Phase 0's code change could do for itself.

A blank cell means nobody has measured it. Every figure here is either **measured** or
explicitly tagged `DERIVED — not measured` with the arithmetic shown; there is exactly one
derived figure in the file and it says so. Nothing is silently estimated or inferred.

The performance severities in [FINDINGS.md](FINDINGS.md) — P1, P2, P5 in particular — were
reasoned from API call counts, never timed, and this file is what turns them into facts or
kills them.

Fill it in, commit it, and treat it as the reference every later phase compares against.

---

## Where each measurement is taken, and why

**The two environments are not interchangeable.** Production batches its reads; the test project
falls through to one round-trip per tab and is roughly **6 × slower on reads** (see
`diagnoseBatchGet()` below, and **P14** in [FINDINGS.md](FINDINGS.md)). So each measurement is
taken where its number is actually meaningful, and each table below says which.

| Measurement | Run in | Why |
|---|---|---|
| Page load, dropdown timings, `testApiPayload()`, `previewOutput()` | **production** | this is the read cost model real users experience. All read-only, bar one throttled `Last_Login_TS` cell — see below |
| Save round-trip **count**, and which functions fire | **test** | client-side behaviour in `index.html`, identical in both environments |
| Payload sizes and row counts | **test** is fine | test currently resolves to the *production* spreadsheet, so same data, same bytes, same counts |
| Production save wall-clock | **derived** | computed from the call list plus production's `testApiPayload()` figures, so that no measurement has to write to production |
| `testRateWrite()`, `testMixWrite()`, `testStructureWrite()`, `testAuditTrail()` | **test only** | they write rows. Never run these against production |

**The only write any measurement in this file causes** is `recordLogin_` (`auth.js:310`) stamping
a single `Last_Login_TS` cell, throttled to once per 15 minutes. That is exactly what happens when
anyone opens the portal, so running `testApiPayload()` in production is no more intrusive than a
page load. Nothing here edits a rate, a mix or an actual in production.

> ⚠️ **Sequencing constraint.** Capture the **test** column figures *before* step 14 of
> [PHASE_0_VERIFICATION.md](PHASE_0_VERIFICATION.md) sets the `SPREADSHEET_ID` Script Property.
> After that, test reads the throwaway copy and its row counts stop matching production's.

---

## Step 0 — record both environments

Run `showEnvironment()` in the **test** project to fill its row. Production's row comes from its
own editor (its deployed code predates Phase 0, so it has no `showEnvironment()` — read the
spreadsheet ID from `utils.gs` there, and the read path from `diagnoseBatchGet()`).

| Field | Production `1JHYSult…` | Test `1Br6gJdV…` |
|---|---|---|
| Date captured | 2026-08-14 | 2026-08-13 |
| Script project ID | `1JHYSultHlRsjZkWcJ9JJ8CNEsBFCUYJVe_n1tENX14_g48ohbe80rGWw` | `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` |
| App spreadsheet ID | | `1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE` |
| Resolved from | n/a — pre-Phase 0 code, literal only | **`fallback`** |
| Spreadsheet file name | | |
| Read path (`diagnoseBatchGet`) | **`rest`** | **`fallback`** |

**Note what the `App spreadsheet ID` row means:** the test project currently points at the **production** spreadsheet,
because no `SPREADSHEET_ID` property is set. That is why its payload sizes and row counts are
valid, and why its write diagnostics must not be run until step 14 points it at a copy.

---

## `testApiPayload()` — both environments

Run it in **production** for the reference figures, and once in **test** as well. Both columns
are needed and they answer different questions:

- **Production** is what real users experience, and what Phases 3, 4 and 5 are re-scored against.
- **Test** is what every phase's own verification runs against — the plan's per-phase step is
  "run `testApiPayload()` and compare against `BASELINE.md`". Without a test column that
  comparison spans the 6 × gap and is meaningless.

Apps Script editor → `Code.gs` → function dropdown → `testApiPayload` → Run. Copy the figures out
of the execution log.

| Measure | Production | Test |
|---|---|---|
| `initApp()` | ____ ms | ____ ms |
| `initApp()` payload | ~____ KB | ~____ KB |
| `loadAllAppData()` | ____ ms | ____ ms |
| `loadAllAppData()` payload | ~____ KB | ~____ KB |
| **Total** | ____ ms | ____ ms |

**Run it at least three times in each environment.** Every run is a **cold** execution — there is
no warm run to compare against, because every cache in this codebase is module-level and resets
between executions (`utils.js` §3; there is no `CacheService` anywhere — finding P6). Within a
single run the caching does work, which is why `loadAllAppData` comes out barely slower than
`initApp` despite reading far more: `initApp`'s prewarm fills `_sheetDataCache_` and
`loadAllAppData` reuses it.

So these samples measure **variance, not warm-up**. Record the median and the range, and put the
median in the table above. Do not quote a single run as the baseline.

| Run | Production `initMs` | Production `bulkMs` | Production total | Test `initMs` | Test `bulkMs` |
|---|---|---|---|---|---|
| 1 | 281 | 399 | **680** | | |
| 2 | 781 | 735 | **1,516** | | |
| 3 | | | | | |
| 4 (optional) | | | | | |
| **Median** | | | | | |
| **Range so far** | 281–781 | 399–735 | 680–1,516 | | |

**Both production runs above are cold and both are legitimate.** The 2.2 × spread is noise, not a
second-run penalty. Corroboration: production's `diagnoseBatchGet()` batched read of 8 tabs
measured **559 ms**, which sits between run 1's entire `initApp` (281 ms — faster than that single
8-tab read) and run 2's (781 ms — slower). The central value is therefore likely ~500–600 ms per
phase, with both runs landing as opposite-direction outliers. Note also that `recordLogin_` is
throttled to once per 15 minutes, so if it fired at all it fired on run 1 — run 1 carried an extra
sheet write and was *still* the faster of the two.

Treat the ±25 % figure from `diagnoseBatchGet()` as a floor on the noise here, not a ceiling.

**Row counts** — one column is enough, because test currently resolves to the production
spreadsheet, so both read identical data. If the two columns above show different payload sizes,
something is wrong with that assumption and it should be investigated before trusting anything
here.

| Table | Rows |
|---|---|
| High Level IDs | |
| Modelling IDs | |
| Base rates | **286** |
| Surcharges | **1,649** |
| Method mixes | |
| Letter/parcel | |
| Cold chain | |
| Output rows | |
| Actuals | |
| Snapshots | |
| Validation findings | |
| Recent audit | |
| `Audit_Log` total rows (read the tab directly — this is the one that grows forever) | |

Identical across both production runs, as expected.

**Publish status at capture:** `last published: run 6`, **0 changes since**.

**Note:** the two counts recorded so far already exceed the figures quoted in
[FINDINGS.md](FINDINGS.md), which were taken at commit `e299923` — base rates 264 → 286,
surcharges 1,418 → 1,649. Ordinary data growth, but it means FINDINGS' row counts are stale and
**this file is now the source of truth** for volumes. It matters most for P5/Phase 4, whose whole
premise is the size of `Audit_Log` — fill that row in from the tab itself, not from FINDINGS.

---

## `diagnoseBatchGet()`

Which bulk-read path is live matters more than any single timing: on the `fallback` path every
later phase's savings are measured against a different cost model.

**Captured 2026-08-13, test script project `1Br6gJdV…`, pointed at the production spreadsheet
(`resolved from: fallback`).**

| Measure | Value |
|---|---|
| Path used | **`fallback`** — neither batch path available |
| Advanced Sheets Service present | **no** |
| REST batchGet result | **HTTP 403** — "Google Sheets API has not been used in project 736650435578 before or it is disabled" |
| Batch attempt, 8 tabs | 219 ms — **not a read, see below** |
| Per-sheet reads, 8 tabs | **3,468 ms** (~434 ms per tab) |
| Speed-up | **not measurable** while on `fallback` |

**Read the 219 ms correctly.** Path 3 of `prewarmSheetCache_` reads nothing — it logs and
returns `'fallback'`, leaving `getAllData_` to read each tab on demand later. So 219 ms is the
cost of *failing* both batch attempts (mostly the REST round-trip to a 403), not the cost of
fetching 8 tabs. The only real read figure here is **3,468 ms for 8 tabs**. `diagnoseBatchGet`
prints no `Speed-up` line on this path, by design.

### The same diagnostic in production

**Captured 2026-08-14, production script project `1JHYSult…`.** Services panel is empty there
too, so no Advanced Sheets Service — but the REST path *works*, which the Services panel could
never have told us.

| Measure | Test (`1Br6gJdV…`) | Production (`1JHYSult…`) |
|---|---|---|
| Path used | **`fallback`** | **`rest`** |
| Advanced Sheets Service | no | no |
| Batch read, 8 tabs | 219 ms *(not a read)* | **559 ms** — a real batched read |
| Per-sheet reads, 8 tabs | 3,468 ms | 2,767 ms |
| Speed-up | not measurable | **4.9 ×** |

### What this actually means — the test environment is the outlier

**Production is already getting the batching.** It reads 8 tabs in **559 ms** where test takes
**3,468 ms** for the same 8 tabs of the same spreadsheet — test is roughly **6 × slower on
reads than production**. The difference is entirely which path `prewarmSheetCache_` lands on,
and the cause is that the test script project's GCP project (`736650435578`) does not have the
Sheets API enabled, while production's does.

Three corrections to what this file said before that check:

1. **The read cost is not a production problem.** Production users are not paying 434 ms per
   tab. The earlier note claiming otherwise was wrong, and is corrected here rather than
   deleted.
2. **Enabling the Advanced Sheets Service is a small win, not a large one.** Paths 1 and 2 hit
   the identical endpoint and the code comment calls them "same speed". Production is already on
   path 2. There is no 5–10 × win sitting there; that number was an artefact of measuring in
   test.
3. **Phases 3, 4 and 5 must be re-scored against production's cost model, not test's.** Every
   read-heavy figure captured in test is inflated about sixfold, which would make Phase 4 in
   particular look far more valuable than it is.

### How this file handles the gap

**Decision: measure in production, keep writes in test.** Rather than blocking the baseline on a
GCP configuration change, each measurement is taken where its number means something — the rule is
in §Where each measurement is taken. Reads and timings come from production; the save's call count
comes from test, where it is equally valid; the production save wall-clock is derived rather than
measured, so nothing has to write to production.

**P14 is still worth fixing, but it is a follow-up, not a blocker.** The cheapest fix is to enable
the Sheets API in the test project's own GCP project, so its REST path behaves exactly as
production's does — no code change, no manifest change, and it makes test *match* production
rather than making both different from each other. Adding the Advanced Service to the repo's
`appsscript.json` would also work, but it moves *both* environments onto a third path and drags
production's behaviour along with it, so it is the worse choice.

Until P14 is fixed, the per-phase verification loop has a trap worth knowing about: every phase is
verified in test, so its `testApiPayload()` numbers must be compared against this file's **test**
column. Compared against the production column, every phase would appear to have caused a ~6 ×
regression.

### One calibration figure worth keeping

Both diagnostics read the same 8 tabs of the same spreadsheet, and the per-sheet timings came
out 2,767 ms and 3,468 ms — a **25 % spread on identical work**. Treat any single Apps Script
timing in this file as ±25 %, and take three samples before believing a difference. That is why
step 12 asks for the middle of three.

---

## Browser wall-clock — the PRODUCTION portal

**Measured against production**, because these are the numbers real users live with and the
figures Phase 3's decision rule turns on.

Get production's address by running `showPortalUrl()` in the **production** Apps Script editor
(`Code.gs` → function dropdown → `showPortalUrl` → Run). Do **not** use `clasp open-web-app` —
clasp points at the test project and would hand you the wrong portal.

Open that `/exec` URL, press F12, go to **Network**, click the **Fetch/XHR** filter and tick
**Disable cache**. Take each measurement three times and record the middle value — Apps Script
round-trips are noisy enough that a single sample means little (see the ±25 % note above).

Everything in this section is read-only: opening the portal and changing dropdowns issues reads.
The one save measurement is handled separately, below, and does not run here.

### Page load

| Measure | Value |
|---|---|
| First paint / frame visible | ____ s |
| `initApp` round-trip | ____ s |
| `loadAllAppData` round-trip | ____ s |
| Fully interactive | ____ s |

### Rates tab — one dropdown change

This is the P1 measurement, and the one that decides whether Phase 3 happens at all. Rates
tab → change the segment dropdown, then the route dropdown. The server call behind it is
`getResolvedByMonth`, which runs the entire forecast engine to populate one table.

| Measure | Value |
|---|---|
| Segment dropdown change, click to redraw | ____ s |
| Route dropdown change, click to redraw | ____ s |
| `getResolvedByMonth` round-trip | ____ s |
| Number of `google.script.run` calls fired | |

**Phase 3 decision rule, from the plan: if a Rates dropdown change in PRODUCTION is already under
1.5s, skip Phase 3.** The environment matters — the same interaction in test would be roughly 6 ×
slower and would keep a phase alive that production does not need. Record the decision here once
measured:

**Phase 3 verdict:** ____________________

Worth timing the same interaction on the Mixes tab while you are here, since
`getMethodMixGrid` and `getLetterParcelGrid` do the same thing and both swallow engine
failures in a bare `catch`:

| Measure | Value |
|---|---|
| Mixes → method mix, Regime dropdown change | ____ s |
| Mixes → letter/parcel, segment change | ____ s |

### One rate save — measured in TEST, production figure derived

This is the P2 measurement and it decides Phase 5. **It is the only measurement that needs a
write, so it is deliberately not done in production** — a rate save there would mutate real
forecast data purely to take a timing.

That costs nothing that matters, because Phase 5's decision rule is the **call count**, and the
call count is client-side behaviour in `index.html` — identical in both environments.

**In the test portal:** Rates tab → edit one rate → Save, with the Network panel open as above.

| Measure | Value | Valid for production? |
|---|---|---|
| Number of `google.script.run` calls fired | | **yes** — client-side, environment-independent |
| Which functions, in order | | **yes** — same |
| Do any two overlap? | | **yes** — same |
| Largest single response | ____ KB | **yes** — test reads the production spreadsheet |
| Save click to "Saved" indicator | ____ s | **no** — upper bound only, inflated on its read portion by ~6 × |

P2 predicts three round-trips, one of which re-reads every tab. Record what actually happened.

**Phase 5 verdict** (from the call count, which is the sound figure): ____________________

#### Derived production save wall-clock

`DERIVED — not measured.` Computed, not observed. Do not quote it as a measurement.

Method: from the call list above, sum production's cost for whichever calls the save fires, using
this file's production `testApiPayload()` figures:

```
derived production save ≈ (production initMs   if the save triggers reloadInit)
                        + (production bulkMs   if the save triggers reloadData)
                        + the saveBaseRate write itself
```

| Term | Value |
|---|---|
| Calls the save fires | |
| Production `initMs` contribution | ____ ms |
| Production `bulkMs` contribution | ____ ms |
| `saveBaseRate` write | ____ ms |
| **Derived production save** | **____ ms** |

**Known weakness, stated so nobody over-trusts it:** `saveBaseRate` has a read component of its
own — `prewarmForWrite_` fetches 5 tabs — which is itself ~6 × slower in test. So taking the
write term from a test measurement makes this figure **understate** the real production save
rather than being exact. It is good enough to decide whether Phase 5 is worth doing; it is not
good enough to quote as a before-and-after.

---

## What this baseline is for

| Phase | Re-scored against | Which column | Drop it if |
|---|---|---|---|
| 3 — engine scope | Rates dropdown change | **production** | already under 1.5s |
| 4 — audit reads | `initMs`, `bulkMs`, `Audit_Log` row count | **production** | `Audit_Log` is small and flat |
| 5 — client fetch | calls per save (sound), derived save wall-clock (indicative) | **test** for the count | already one call, already fast |

Phases 1, 2, 6, 7 and 8 are correctness or safety work and do not depend on these numbers.
They should not get slower, though — after each phase, run `testApiPayload()` in test and compare
against this file's **test** column, not the production one. Comparing a test run against the
production column would show a spurious ~6 × regression every time.

Once **P14** is fixed and test's read path matches production's, the two columns should converge
and this distinction stops mattering. Until then it matters on every phase.
