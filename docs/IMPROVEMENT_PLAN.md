# Postage Forecast Portal — sequenced improvement plan

Derived from [FINDINGS.md](FINDINGS.md) and [ARCHITECTURE.md](ARCHITECTURE.md).

**Nothing here has been implemented.** Phases 0–8 are ready to execute. Everything after
Phase 8 is deliberately left unplanned because it depends on a decision only you can make —
those questions are in [§Decisions required](#decisions-required) and I have stopped rather
than guessed.

**Revision 2.** D3 and D8 answered, D1 answered in part. Net effect on the plan: S6 is
unblocked and joins Phase 7; `include()` is now **kept and constrained** rather than deleted;
M6's mechanism is settled even though its scope is not. No build step, no TypeScript, no
bundler — the project stays plain hand-editable JavaScript, so phase numbering, effort and
sequencing are all unchanged from revision 1.

---

## Environments

Recorded per D1. **This is documentation, not configuration** — no clasp config for
production is being created, and nothing is pushed to production until the promotion step
late in this plan.

| Role | Script ID | Status |
|---|---|---|
| Production | `1JHYSultHlRsjZkWcJ9JJ8CNEsBFCUYJVe_n1tENX14_g48ohbe80rGWw` | Recorded only. Do not push. No `.clasp.prod.json` until the promotion step. |
| Test / dev — committed in `.clasp.json` | `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` | **Confirmed test target.** What `clasp push` sends to, and where every phase is verified. |

The two are different projects, deliberately. The repo was cloned from the test copy, so a
plain `clasp push` reaches test and cannot reach production by accident — that is the safe
default and should stay that way for all of Phases 0–8. Promotion to production is a
separate, later runbook step with its own configuration, deliberately not built yet.

**Pre-work, owned outside this plan:** the test copy's `Permissions` tab is to be anonymised
directly in the sheet before Phase 0 begins, so that test audit rows do not accumulate
against real staff email addresses. This is a data change, not a code change — it is not part
of any phase and needs no branch. Worth doing first because `recordLogin_` stamps
`Last_Login_TS` and every diagnostic write lands an `Audit_Log` row carrying the acting
user's address; the longer testing runs, the more there is to clean up afterwards.

---

## Ground rules applied

- **Behaviour-preserving first.** Phases 0–8 must produce byte-identical forecast output.
  Anything that changes what a user sees is deferred past Phase 8.
- **One phase, one branch, one push, one verification.** Each phase below names its branch
  and its revert.
- **No phase mixes backend and UI.** Only one phase (5) touches `index.html`, and it touches
  nothing else. Where a fix would naturally span both, I've split it and said so.

### The regression oracle — read this once, use it every phase

The codebase already contains an exact, whole-model regression check. No new tooling needed:

1. **Before starting a backend phase**, from the Apps Script editor run `publishOutput()`.
   `OUTPUT` is now current and a `Calc_Runs` row records the baseline.
2. **After the change**, run `previewOutput()`. It writes nothing and diffs the recomputed
   model against the published `OUTPUT`. It must log:
   `Nothing would change. OUTPUT already matches the inputs.`
   Any other result means the phase changed the forecast — stop and revert.
3. Also run `runEngineUnitTests()` (must stay all-pass) and `testApiPayload()` (for timings).

That covers the engine and the publish path exactly. It does **not** cover the write paths;
those are covered by the existing self-cleaning diagnostics `testRateWrite()`,
`testMixWrite()`, `testStructureWrite()` and `testAuditTrail()`, which are named per phase
below where relevant.

---

## Ordering, and where I traded the formula off

Scoring is (severity × confidence) ÷ effort, where confidence is how sure I am that the
finding is real *and* that the fix is safe. Severity High=3 / Med=2 / Low=1. Effort S=1 /
M=3 / L=8.

| Finding | Sev | Conf | Effort | Score | Phase |
|---|---|---|---|---|---|
| C1 authorise-the-payload | 3 | 0.95 | 1 | **2.85** | 1 |
| S1 view-scope grants edit | 3 | 0.95 | 1 | **2.85** | blocked (D2) |
| P4 Audit_Log prewarmed per write | 2 | 0.90 | 1 | 1.80 | 2 |
| P8 live read per row | 2 | 0.90 | 1 | 1.80 | 2 |
| S5 requireMaintenance_ fails open | 2 | 0.90 | 1 | 1.80 | 7 |
| S6 ALLOWALL framing | 2 | 0.90 | 1 | 1.80 | 7 |
| S4 ungated snapshot job | 2 | 0.85 | 1 | 1.70 | 7 |
| C2 nested lock | 2 | 0.80 | 1 | 1.60 | 1 |
| S7 unbounded DENIED logging | 2 | 0.80 | 1 | 1.60 | 7 |
| C6 ragged-row write | 2 | 0.70 | 1 | 1.40 | 6 |
| U1 features with no UI | 3 | 1.00 | 3 | 1.00 | blocked (D9) |
| P3 per-row actuals import | 3 | 0.90 | 3 | 0.90 | 6 |
| P1 whole engine per screen | 3 | **0.70** | 3 | 0.70 | 3 |
| P2 full reload per save | 3 | **0.70** | 3 | 0.70 | 5 |
| P5 audit read cost | 2 | 0.85 | 3 | 0.57 | 4 |
| M3 no tests | 3 | 1.00 | 8 | 0.38 | see below |
| M6 monolithic client | 3 | 1.00 | 8 | 0.38 | blocked (D9) |

**Four places I deviated from that order, deliberately:**

1. **Phase 0 ranks last and I put it first.** `S9` (hardcoded spreadsheet IDs) is Low
   severity, but your own workflow — "one push to the test copy" — is *impossible* while
   `SPREADSHEET_ID` is a compile-time constant. Every other phase's verification depends on
   it. Sequencing constraint beats the score.

2. **P1 and P2 score mid because my confidence is only 0.70, not because they're small.**
   Both severities are reasoned from API call counts; I have not measured either. Phase 0
   captures the baseline that either raises those to ~1.0 or kills them. **Phase 3 and
   Phase 5 should be re-scored after Phase 0 and may be dropped or promoted.** I've kept
   them where a 0.7 confidence puts them rather than where my instinct does.

3. **C1 tops the table but is dormant.** With `Scope_Mapping` empty and
   `SCOPE_DEFAULT_ALLOW = TRUE`, nobody is scoped, so C1 cannot currently be exploited. I
   kept it at Phase 1 as cheap insurance — it's five lines in five files and it removes a
   landmine before anyone switches scoping on and assumes it works. **If D2 comes back "we
   will never use scoping", demote Phase 1 to the Phase 8 sweep and drop S1 entirely.**

4. **M3 (no tests, effort L) is not a phase.** Instead Phase 3 carries its own in-run
   equivalence check — a maintenance function that computes the model both ways in a single
   execution and asserts they match for all 232 routes. That buys the regression safety M3
   was for, at effort S instead of L. A proper Node harness is worth doing before the Phase 8
   refactor; it is listed there, not here.

---

## Phase 0 — Make the environment switchable, and measure

**Branch** `phase-0-environment` · **Effort** S · **Backend only**

Everything else depends on this. Right now there is one spreadsheet ID compiled into
`utils.js`, so there is no way to point the code at a test copy without editing source you
then have to remember not to commit.

**Out of scope, per D1:** this phase does **not** create `.clasp.prod.json`, does not change
the committed `scriptId`, and does not push anything to the production project. Promoting to
production is a late runbook step, not part of the workbench setup. The only clasp-related
change is documentation — recording both IDs, which §Environments has already done.

**What changes**

- `SPREADSHEET_ID` and `SOURCE_SPREADSHEET_ID` read from Script Properties
  (`SPREADSHEET_ID`, `SOURCE_SPREADSHEET_ID`), **falling back to the current literal** when
  the property is absent. Production keeps working with no configuration change.
- A `showEnvironment()` maintenance function logging which spreadsheet is in use and whether
  it came from a property or the fallback — so you can never be unsure which file you are
  about to write to.
- Capture the baseline: run `testApiPayload()` and record `initMs`, `bulkMs` and payload
  sizes; open the portal with the browser network panel and record the wall-clock for a
  Rates-tab dropdown change and for one rate save. Paste both into a new
  `docs/BASELINE.md`.

**Files** `utils.js`, `migrate.js`, `docs/BASELINE.md` (new)

**What could break** If the fallback is wrong, the app writes to the wrong spreadsheet —
the highest-consequence mistake available in this codebase. The fallback must be the
existing literal, and `showEnvironment()` must be run before anything else.

**How you verify**
1. Editor → run `showEnvironment()`. Logs the production ID and "from fallback".
2. Editor → `diagnoseBatchGet()` and `showPortalUrl()` still pass.
3. Open the portal. Dashboard renders, "Last published" is unchanged.
4. Set the Script Property to a **copy** of the sheet, run `showEnvironment()` again, confirm
   it now says "from property", and confirm the portal loads the copy. Then unset it.

**Revert** `git revert` the commit; the literals are unchanged so production is unaffected.

---

## Phase 1 — Write-path correctness

**Branch** `phase-1-write-correctness` · **Effort** S · **Backend only**

Fixes C1, C2, C8. No user-visible change whatsoever.

**What changes**

- **C1** — on update, authorise **both** the row's existing owner and the proposed new
  owner. The correct pattern already exists in the same files (the delete paths read the row
  first, then assert). Five call sites: `saveBaseRate`, `saveSurchargeRate`,
  `saveLetterParcelMix`, `saveActual`, `saveModellingId`.
- **C2** — `bulkRateChange` stops calling `saveBaseRate` inside its own `withLock_`. Extract
  the unlocked body of `saveBaseRate` into a private `writeBaseRate_` that both the locked
  single-save and the already-locked bulk path call.
- **C8** — `logAudit_` takes the lock, or allocates `Log_ID` from a Script Property counter.
  Prefer the property: it avoids serialising the DENIED path behind real writes.

**Files** `rates.js`, `mixes.js`, `actuals.js`, `structure.js`, `audit.js`

**What could break** The C2 extraction is the risk: `saveBaseRate` currently invalidates the
sheet cache at three points inside its lock, and the ordering of
`closePrecedingPeriod_` → `assertNoOverlap_` → write matters. Preserve that order exactly in
the extracted function. If `Log_ID` moves to a property, the first run must seed it from the
sheet's current maximum or IDs restart at 1.

**How you verify**
1. Editor → `testRateWrite()`. Must log `RATE WRITES WORKING` — it exercises create, refused
   overlap, update, history and delete on a real route.
2. Editor → `testMixWrite()` → `MIX WRITES WORKING`; `testStructureWrite()` →
   `STRUCTURE GUARDS WORKING`; `testAuditTrail()` → `AUDIT TRAIL WORKING`.
3. In the app: Rates → any segment → any route → **Add rate change** from a date inside an
   existing period. Confirm the previous period's "To" moves to the day before, and both rows
   appear correctly.
4. Rates → **Delete** that test rate. Confirm it disappears from the list.
5. History → Change log. Confirm the create, the auto-close and the delete each appear with
   your email, and that `Log_ID`s are unique and ascending.
6. Editor → `previewOutput()` → `Nothing would change` (after cleaning up the test rate).

**Revert** Single commit revert. No data migration, no schema change.

---

## Phase 2 — Cheap backend wins

**Branch** `phase-2-cheap-wins` · **Effort** S · **Backend only**

Six small fixes that score highly and share one verification. Bundled because they are all
backend, all invisible, and a single revert of the lot is coherent.

**What changes**

- **P4** — `prewarmForWrite_` stops fetching `Audit_Log`. `getNextId_` reads only the ID
  column via `getRange(2, col, lastRow-1, 1)` when the sheet isn't already cached (it already
  has this branch), or takes the Phase 1 property counter.
- **P8** — `readRow_` returns from `_sheetDataCache_` when the sheet is cached, falling back
  to the live read. Removes one round-trip per save and per row in `closePrecedingPeriod_`.
- **P11** — `clearMixWindowBatch_` and `clearLetterParcelWindow_` write only the contiguous
  span of changed rows rather than the whole data range.
- **P12** — `PERF_LOG_ENABLED` reads from Config (default `FALSE`); the log-building blocks
  in `previewOutput` and `reportFindings_` are skipped when nothing will read them.
- **C10** — `invalidateSheetCache_` also clears `_hlMetaCache_` / `_midMetaCache_` when the
  structure tabs are invalidated, and `_configCache_` for Config.
- **M7** — `currentUserEmail_` delegates to `getActiveEmail_`; `logAction_` delegates to
  `logAudit_`; the two literal `25569`s become `SERIAL_EPOCH_OFFSET`.

**Files** `utils.js`, `rates.js`, `mixes.js`, `audit.js`, `output.js`, `migrate.js`

**What could break** P8 is the one to watch: `readRow_` currently always returns a full-width
row from `getRange`, whereas the cache can hold a **ragged** row (ARCHITECTURE §5.4). The
cached path must pad to `TABLES[key].headers.length` before returning, or every `row.slice()`
downstream writes a short array and `setValues` throws. P4 must not break `getNextId_` for
tabs that have never been written.

**How you verify**
1. Editor → `testRateWrite()`, `testMixWrite()`, `testAuditTrail()` — all pass as Phase 1.
2. Editor → `testApiPayload()`. Compare `initMs`/`bulkMs` against `docs/BASELINE.md`.
3. In the app: save a rate, then a mix grid, then a letter/parcel split. Each must save
   cleanly and the History → Change log entries must be identical in shape to before.
4. Mixes → method mix → save a grid that overlaps two existing periods. Confirm the trimmed
   periods are correct (this is the P11 blast radius).
5. `previewOutput()` → `Nothing would change`.

**Revert** Single commit revert. If only one sub-fix misbehaves, each is a separate hunk and
can be reverted individually.

---

## Phase 3 — Stop running the whole engine to render a screen

**Branch** `phase-3-engine-scope` · **Effort** M · **Backend only**

The largest performance item (P1), and the one my confidence is lowest on. **Re-score this
against `docs/BASELINE.md` before starting — if a Rates dropdown change is already under
1.5s, skip this phase.**

**What changes**

- `computeModel_` accepts an optional `{ onlyHlId }` / `{ onlyModellingId }` filter that
  narrows `validMids` and `hls` before the month loop. It stays pure.
- When a filter is present, `checkMixInvariant_` is skipped — the three read screens ignore
  `defects` entirely, and a partial invariant check would be misleading.
- `loadEngineInput_` gains the same optional filter so it stops materialising rows it will
  not use.
- `getResolvedByMonth`, `getMethodMixGrid` and `getLetterParcelGrid` pass the filter.
- The two bare `catch (e) { /* rates are a nicety */ }` blocks in `mixes.js` log the error
  instead of swallowing it silently.

**Explicitly not in this phase:** reading published `OUTPUT_Detail` instead of recomputing.
That would be faster still, but the panels currently show **unpublished** state, and reading
the published tab would change what the user sees. That is a Phase 9+ decision.

**Files** `engine.js`, `rates.js`, `mixes.js`

**What could break** Filtering must not change any number. The risk is `methodMix`, which
depends on `ccShare` and `letterMix` resolved **per High Level ID** — filtering to one
Modelling ID must still resolve its parent segment's cold-chain and letter shares, or every
rate silently changes.

**The equivalence check (build this as part of the phase).** Add a maintenance function
`verifyFilteredEngine()` that, in one execution, runs `computeModel_` unfiltered, then runs
it once per segment filtered, and asserts every `outputDetailRow` matches to 1e-9. It logs
`FILTERED ENGINE EQUIVALENT — 8352 of 8352 rows match` or names the first divergence. This
is the whole regression safety net for the phase and costs an hour.

**How you verify**
1. Editor → `verifyFilteredEngine()` → all rows match.
2. Editor → `runEngineUnitTests()` → all pass.
3. Editor → `previewOutput()` → `Nothing would change`.
4. In the app: Rates → MedExpress GB WL Mounjaro → any route. **Screenshot the "Resolved by
   month" table before the change and after.** All 36 rows identical.
5. Mixes → method mix → change the Regime dropdown from AMBIENT to CC and back. Time it
   against `docs/BASELINE.md`. Rate/parcel and Contribution columns unchanged.
6. Mixes → letter/parcel → confirm Blended rate figures unchanged.

**Revert** Single commit revert. `computeModel_`'s signature change is additive (optional
argument), so reverting cannot strand a caller.

---

## Phase 4 — Bounded audit reads

**Branch** `phase-4-audit-reads` · **Effort** M · **Backend only**

P5 and the read half of P7. `Audit_Log` is at 767 rows and read in full three times per page
load; this is the cost that grows forever.

**What changes**

- `readRecentAudit_` reads only the last ~4× `limit` rows via
  `getRange(max(2, lastRow - n), 1, …)` instead of the whole tab.
- `getPendingChangeCount_` scans **backwards** from the last row and stops at the first entry
  older than the newest `Calc_Runs.Run_TS`. Same answer, bounded work.
- `loadStatus_` stops being called twice per page load: `initApp` keeps it; `loadAllAppData`
  returns the same shape but computes `pendingChanges` from the audit rows it has already
  read for `recentAudit`.
- `Audit_Log` and the eight `*_Amends` tabs come out of the `prewarmSheetCache_` lists where
  the full contents are not needed.

**Files** `audit.js`, `Code.js`, `utils.js`

**What could break** The backwards scan assumes `Audit_Log` is append-ordered by timestamp.
It is, because every writer appends — but `recordChangesBatch_` writes a block with one
shared `now`, and the migration wrote rows out of band. Scan far enough past the boundary
(say 50 rows) before stopping, rather than breaking at the first older row.

**How you verify**
1. Note the "N changes unpublished" badge in the header. Make one rate edit. Confirm the
   badge increments by the same amount it did before this phase (it over-counts per field —
   that is C11 and is *not* being fixed here, so the number must stay wrong in the same way).
2. History → Change log. Confirm the same 50 entries appear, newest first, identical.
3. Editor → `testApiPayload()` — `initMs` and `bulkMs` against baseline.
4. Editor → `publishOutput()`, then confirm the badge returns to zero.
5. `previewOutput()` → `Nothing would change`.

**Revert** Single commit revert.

---

## Phase 5 — Client fetch discipline

**Branch** `phase-5-client-fetch` · **Effort** M · **UI only — the only phase that touches `index.html`**

P2, and the fetch half of U3. Strictly invisible except that it is faster: same data, same
rendering, fewer round-trips. **Re-score against `docs/BASELINE.md` first.**

**What changes**

- `call()` de-duplicates identical in-flight requests, so the handlers that fire
  `loadMixGrid()` and `reloadData()` together stop issuing overlapping calls.
- `reloadInit()` stops chaining into `reloadData()` unconditionally; the two are called
  explicitly where each is actually needed.
- Handlers that currently call both a grid reload and a full reload call only what changed.

**Explicitly not in this phase:** having write endpoints return the saved row so the client
can patch `S` locally. That is the bigger win, but it changes server return shapes, which
would make this a backend-and-UI phase. It is Phase 9+.

**Files** `index.html` only

**What could break** Over-aggressive de-duplication could suppress a legitimate refetch and
leave the screen stale after a save — which looks like a lost edit. Key the de-duplication on
function name **plus** serialised arguments, and only for in-flight requests, never as a
result cache.

**How you verify** With the browser network panel open:
1. Rates → edit a rate → Save. Count the `google.script.run` calls: was 3, should be fewer.
   The new value appears in the table and in Resolved by month.
2. Mixes → method mix → Save grid. Confirm the grid reloads with the saved values **and** the
   dashboard's forecast figures update. This is the regression most likely to bite.
3. Admin → add a carrier. Confirm it appears in the Methods form's carrier dropdown
   immediately — that path depends on `reloadInit` refreshing reference data.
4. Actuals → add a month → confirm the comparison table and the dashboard chart both update.
5. Leave the tab idle 30s, then save. Confirm the "Saved / Saving… / Not saved" indicator
   still transitions correctly.

**Revert** Single commit revert of `index.html`.

---

## Phase 6 — Import and publish robustness

**Branch** `phase-6-robustness` · **Effort** M · **Backend only**

P3, C6, C7.

**What changes**

- **P3** — `importActuals` builds all rows in memory, writes them with one ranged
  `setValues`, and records history with the existing `recordChangesBatch_`. The
  `{ written, skipped, errors }` return contract is preserved, with per-row validation still
  done before the batch so a bad row is skipped rather than failing the import.
- **C6** — `appendCopies` pads every sliced row to `TABLES[key].headers.length` before
  writing. Same fix pattern as Phase 2's `readRow_`; consider one shared `padRow_` helper.
- **C7** — `publishOutput` writes `OUTPUT`/`OUTPUT_Detail` without leaving a window where
  they are empty: build the full value array first, then `clearDataRows_` and `setValues` as
  the last two operations, and write the `Calc_Runs` row only after both succeed.

**Files** `actuals.js`, `structure.js`, `output.js`, `utils.js`

**What could break** P3 changes failure granularity — today one bad row throws inside its own
`saveActual` and is skipped; batched, a single bad row could reject the block. Validate every
row up front, collect failures, batch only what passed. C7 does not make publish atomic
(Apps Script has no transactions); it narrows the window. Say so in the code comment rather
than implying more.

**How you verify**
1. Editor → `previewActualsImport()`. Note "actual rows to write". Then `runActualsImport()`.
   Compare `written` and `skipped` against the previous run's log — must match exactly.
2. Editor → `diagnoseActuals(1)`. Confirm no new zero-rate rows appeared (that is the symptom
   of a parse or write regression).
3. In the app: Actuals tab → confirm the same months, same blended rates, same variances as
   before the import.
4. Admin → High Level & Modelling IDs → **Copy rates & mixes** into a segment. Preview, then
   copy. Confirm the reported row count matches what lands, and no "number of columns" error.
5. Editor → `publishOutput()` → succeeds; `previewOutput()` → `Nothing would change`.
6. Time `publishOutput()`. Confirm it is comfortably under 6 minutes.

**Revert** Single commit revert. If the actuals import has already run, `clearActuals()` +
re-import restores.

---

## Phase 7 — Access-control hardening

**Branch** `phase-7-access-control` · **Effort** S · **Backend only**

S4, S5, S6, S7, S11. S6 joins this phase now that D3 confirms the portal is never embedded.

**What changes**

- **S4** — `weeklySnapshotJob` **keeps its name** (CLAUDE.md: preserve names wired to
  triggers — renaming it silently breaks the installed trigger). Instead it gates on caller:
  it proceeds when there is no active user (a trigger) and requires `runCalc` otherwise. The
  55-minute rate limit stays as defence in depth.
- **S5** — `requireMaintenance_` distinguishes "Permissions tab is empty" (bootstrap, allow)
  from "the read threw" (fail **closed**). The bootstrap allowance also gets a log line, so
  its use is visible.
- **S6** — `setXFrameOptionsMode(ALLOWALL)` becomes `DEFAULT`, so the portal can only be
  framed by pages on the same origin. D3 confirms nothing embeds it, so this is invisible.
- **S7** — DENIED audit rows are rate-limited per email per execution window, so a loop
  cannot inflate the log.
- **S11** — **revised, see below.** `include()` is *kept* and constrained to a whitelist of
  known partial names rather than accepting any filename. `createTemplateFromFile` stays.

> **Why S11 changed.** Revision 1 deleted `include()` and swapped `createTemplateFromFile`
> for `createHtmlOutputFromFile`, with a warning that this depended on D8. D8 came back "no
> build step" — which makes `include()` the *only* mechanism for splitting the 2,313-line
> `index.html` into partials, and templating is what makes `include()` work at all. Deleting
> it would remove the tool M6 needs. Constraining it closes the arbitrary-file-read seam
> while leaving the mechanism intact.

**Files** `snapshots.js`, `utils.js`, `auth.js`, `Code.js`

**What could break** S4 is the one with teeth: if the gate is wrong, the Monday trigger stops
running and nobody notices for a week — the failure is silent by nature. S5 fails closed,
which is correct, but a transient sheet read error now blocks maintenance functions rather
than opening them. S6 is safe given D3, but if the portal is ever *later* embedded, the
symptom is a blank frame with a console error and no server-side trace — worth a comment in
`doGet` recording why the mode is what it is.

**How you verify**
1. Editor → `showSnapshotTriggers()`. Confirm the weekly trigger is still installed.
2. Editor → `weeklySnapshotJob()` run manually. Confirm it still works for an admin.
3. **Wait for the Monday run**, or temporarily install a trigger a few minutes out, and
   confirm from Executions that it completed and either stored a snapshot or logged
   "No change". This phase is not verified until a real trigger fire has been observed.
4. History → Forecast over time. Confirm snapshots still list and **Compare** still works.
5. Open the portal directly at its `/exec` URL. It loads and renders normally — this is the
   S6 check, and per D3 it is the only way the portal is ever opened.
6. Editor → `testMyPermissions()` → unchanged output.

**Revert** Single commit revert. If the trigger was disturbed, `installWeeklySnapshot()`
reinstates it.

---

## Phase 8 — Extract the shared write path

**Branch** `phase-8-write-helper` · **Effort** M · **Backend only**

M1, and M2 alongside it. Behaviour-preserving structural work, deliberately last of the
unblocked phases because it has the widest blast radius and is safest once Phases 1–7 have
settled.

**What changes**

- One `saveDatedRow_(tableKey, opts)` replacing the near-identical bodies of `saveBaseRate`,
  `saveSurchargeRate`, `saveLetterParcelMix` and `saveActual` — roughly 250 lines removed,
  and C1's fix now lives in exactly one place.
- **M2** — `runValidation` and `runValidationQuiet_` share one rule list, so a new rule
  cannot run in one path and not the other.

**Recommended prerequisite:** a Node regression harness for the pure engine. `computeModel_`
has no Google dependencies, so a small `test/` directory with a `vm`-based loader can run
`engine.js` + `utils.js` **unmodified** — no build step, no TypeScript, and no changes to
shipped code, which is exactly what D8 asks for. The harness reads the same `.js` files that
`clasp` pushes; it does not transform them. It needs a `.claspignore` so `clasp push` does
not upload the test directory. This is the cheap half of M3 and is worth having before the
refactor, not before Phase 3 (which carries its own in-run check).

**Files** `rates.js`, `mixes.js`, `actuals.js`, `validate.js`, `output.js`, plus `test/` and
`.claspignore` if the harness is built

**What could break** Each of the four save functions has genuinely different validation —
percentage bounds, currency defaulting, the derive-rate-from-spend rule, the
`closePrevious` behaviour. A shared helper that flattens those differences will silently
change validation. The helper must take the type-specific validation as a callback rather
than trying to generalise it. If `.claspignore` is wrong, `clasp push` can stop uploading
real files — check the pushed file list, do not assume.

**How you verify**
1. Editor → all four write diagnostics: `testRateWrite()`, `testMixWrite()`,
   `testStructureWrite()`, `testAuditTrail()`. All must report WORKING.
2. Deliberately trigger each refusal in the UI and confirm the **message text is unchanged**:
   a negative base rate; a percentage surcharge entered as `14` instead of `0.14`; a mix grid
   totalling 95%; a mix grid with a route omitted; an overlapping rate period; an actual with
   neither orders nor a rate.
3. Editor → `runValidation(1)`. Confirm the finding counts by rule are identical to the last
   run before the change.
4. Editor → `publishOutput()`, then `previewOutput()` → `Nothing would change`.
5. `clasp push` and confirm from the output that every `.js` and `index.html` was uploaded.

**Revert** Single commit revert — but this is the phase most worth pushing to the test copy
and living with for a day before promoting.

---

## Decisions required

Each item below is a phase I could plan properly once you answer, and would be guessing at
otherwise. **D1, D3 and D8 are answered** and marked as such; the rest still block.

**D1 — Is there a test copy of the spreadsheet, and a second Apps Script deployment to push
it to?** ✅ **Answered.** The `scriptId` already committed in `.clasp.json` (`1Br6gJdV…`) is
the test/dev copy and stays the target for all of Phases 0–8; production
(`1JHYSult…`) is recorded as documentation only, with no clasp config and no push until the
promotion step. Both are in §Environments, along with the one piece of pre-work — anonymising
the test copy's `Permissions` tab — which happens in the sheet, before Phase 0, outside this
plan.

**D2 — Will row-level scoping ever be switched on?** `Scope_Mapping` is empty and
`SCOPE_DEFAULT_ALLOW` is `TRUE`, so C1, S1 and S3 are all currently unexploitable.
*If yes:* S1 becomes a phase of its own, S3 needs scope filters on four return paths, and
Phase 1 is correctly placed. *If no:* drop S1 and S3, demote Phase 1, and `Scope_Mapping` +
`loadScopes_` + `scopeAllows_` become dead code to delete.

**D3 — Is the portal embedded in any other page?** ✅ **Answered: no.** The portal is only
ever opened directly at its own web app URL. S6 is therefore unblocked and scheduled into
Phase 7 — `ALLOWALL` becomes `DEFAULT`, with no user-visible effect.

**D4 — Is the legacy workbook still reconciled against?** `runParityTest()` and
`verifyPublishedOutput()` still read `SOURCE_SPREADSHEET_ID`. *If it's dead:* M4 removes
~900 lines. *If it's live:* `runParityTest()` is an excellent end-to-end oracle and should be
added to the verification steps for Phases 3, 6 and 8.

**D5 — Are scenarios a real requirement?** Every client call site passes a hardcoded `1`.
*If yes:* C5 becomes urgent — the client loaders need scenario filters and `assertNoOverlap_`
needs scenario in its key, plus a scenario picker in the UI. *If no:* drop the column and the
table.

**D6 — `Applies_To` and `Apply_Order`: implement or remove?** The engine ignores both while
the Admin form invites you to set them. Implementing them means surcharges can compound and
be ordered — a real change to what the numbers mean, and I'd want to know whether any current
surcharge *should* compound before touching the engine. Removing them is a schema change plus
a form change.

**D7 — Cold chain: build an editor, or leave read-only?** Today the screen tells users to
edit `Mix_ColdChain` directly, which bypasses validation, locking and audit. Either build the
editor (U8, effort M) or remove the instruction and accept it as an admin-only tab.

**D8 — Is a build step acceptable?** ✅ **Answered: no.** No TypeScript, no bundler, no
`src/` → `dist/` pipeline. The project stays plain, hand-editable JavaScript that can be
edited directly in the Apps Script editor, and `clasp push` keeps sending exactly what is in
the repo.

Three consequences, all now folded into the plan:

- **`include()` is kept and constrained, not deleted** (Phase 7, S11). Without a bundler it
  is the only mechanism for splitting `index.html` into partials, and `createTemplateFromFile`
  is what makes it work — so both stay. This is the dependency revision 1 flagged in advance.
- **M6's mechanism is settled even though its scope is not.** Under "no build step", breaking
  up the 2,313-line client means multiple `.html` partials stitched together with
  `<?!= include('…') ?>`, not ES modules. Still blocked on D9 for *how far* to go, but the
  technique is no longer an open question.
- **Phase 8's Node harness is unaffected** — it always loaded the shipped `.js` files
  unmodified through `vm`, precisely so it would need no build step. It stays as written.

The trade-off you are accepting, stated plainly so it is not a surprise later: no type
checking across ~11,300 lines that address sheet columns by string key, and no compile-time
catch for the class of typo that `COL.RATE_BASE.Base_Rat` would produce (which evaluates to
`undefined` and silently writes to column 0). The mitigations available without a build step
are the Phase 8 Node harness and the existing `test*Write()` diagnostics. Worth knowing that
is the safety net you have.

**D9 — Is the UI overhaul a restyle of the current eight screens, or a rethink of the
journeys?** This determines everything after Phase 8. Under "restyle", U4–U7 and U9 are a
tidy-up phase and U1 is about surfacing `bulkRateChange` and `getRecordHistory` on existing
screens. Under "rethink", the right move is to design the journeys first and treat M6 as a
rebuild rather than a refactor — and several UI findings become moot.

**D10 — Does anyone edit the spreadsheet directly?** If yes, the audit trail has holes by
design, S8's priority rises, and it's worth discussing an `onEdit` shim that at least records
out-of-band changes.

---

## Phases 9+ — not planned

Deliberately empty. These are the findings I will not sequence without the answers above:

- **C3** (segments with no mix rows are silently 0%) — the fix adds a validation finding.
  Whether it lands as WARN or ERROR matters, because an ERROR blocks publishing and I don't
  know how many existing segments would trip it. Plan: run it report-only first, count the
  findings, then decide. Needs D5 too, since coverage is per scenario.
- **C4, C5** — need D6, D5.
- **S1, S3** — need D2.
- **S13** (raw HTML seams in `openForm`), **U2, U4, U5, U6, U7, U8, U9** — all UI, all need
  D9; several are moot under "rethink".
- **U1** — needs D9 for placement, and a product call on whether `bulkRateChange` should be
  exposed to every role with `editRates` or gated further.
- **M4** — needs D4.
- **M5** (dead code sweep) — overlaps D4, D5 and D6. Best done once, after those land. Note
  that D8 has removed `include()` from the dead-code list: it is now load-bearing for M6.
- **M6** — needs D9 only. D8 has settled the technique (HTML partials joined by `include()`);
  what remains is how far to go, which is the restyle-versus-rethink question.
- **M8** (`SHEET_STRUCTURE.md` regeneration script) — trivial, but it should follow whatever
  schema changes come out of D5 and D6 rather than being written twice.

---

## Summary

| Phase | Branch | Effort | Touches | Ships |
|---|---|---|---|---|
| 0 | `phase-0-environment` | S | backend | switchable environment + measured baseline |
| 1 | `phase-1-write-correctness` | S | backend | C1, C2, C8 |
| 2 | `phase-2-cheap-wins` | S | backend | P4, P8, P11, P12, C10, M7 |
| 3 | `phase-3-engine-scope` | M | backend | P1 + equivalence check |
| 4 | `phase-4-audit-reads` | M | backend | P5, P7 (read half) |
| 5 | `phase-5-client-fetch` | M | **UI** | P2, U3 (fetch half) |
| 6 | `phase-6-robustness` | M | backend | P3, C6, C7 |
| 7 | `phase-7-access-control` | S | backend | S4, S5, S6, S7, S11 |
| 8 | `phase-8-write-helper` | M | backend | M1, M2 (+ optional Node harness) |

Eight backend phases, one UI phase, no phase mixing the two. Phases 3 and 5 should be
re-scored against `docs/BASELINE.md` after Phase 0 and may be dropped.

**Nothing has been implemented.** Phase 0 is ready to start against the confirmed test
project, once the test copy's `Permissions` tab has been anonymised (see §Environments).
