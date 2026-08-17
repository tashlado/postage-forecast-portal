# Postage Forecast Portal — findings register

Produced from a complete read of every `.gs` file, `index.html`, `appsscript.json` and
`.clasp.json` at commit `e299923`. Architecture and context: [ARCHITECTURE.md](ARCHITECTURE.md).

**Severity describes consequence, not effort.** High = wrong numbers, data exposed to the
wrong person, or a hard blocker on the work ahead. Medium = real cost or real risk, but
contained or currently dormant. Low = worth fixing when nearby.

**Effort** is S (under an hour), M (half a day to two days), L (a week or more).

Every file-and-line reference below was re-verified against the source before publishing.

---

## What is already right

This register is a list of defects, so read on its own it gives a misleading impression.
Before the table, the things that are done well and should survive any refactor:

- **The engine is pure.** `computeModel_()` takes plain objects and returns plain objects,
  touching no Google service. That is what makes parity testing against the legacy workbook
  possible and preview-without-writing free.
- **One schema definition.** `TABLES` in `utils.js` drives sheet creation, `COL`, the
  auto-generated Amends tables and the reconciliation in `setupCreateAllTabs()`. There is
  **no hardcoded column index anywhere in application code**, and columns are matched by
  name so the sheet and the code cannot drift.
- **The serial-date trap is handled centrally and correctly**, including the duck-typed
  `isDate_`, the bounds check that stops a rate being read as a date, and the integer
  `dayNum` arithmetic that no timezone can shift.
- **Server-side scope filtering.** Rows outside a user's scope are never serialised, let
  alone hidden client-side. Every client-side permission check is backed by a server check.
- **The Apps Script performance basics are already right.** No `flush()` calls anywhere. No
  per-cell `getValue`/`setValue` in any hot path — three `setValue` calls exist in the whole
  codebase, two in setup/maintenance loops and one single-cell login stamp. Every bulk write
  is a ranged `setValues`. `prewarmSheetCache_` is *written* to collapse many tab reads into one
  HTTP call, with two documented fallbacks. `Utilities.formatDate` was already hunted down and
  replaced with plain JavaScript after it was measured costing 2.6s of a 3s page load.
  **Confirmed by measurement in Phase 0:** in production the batching is genuinely active on
  the REST path, worth a measured 4.9 ×. The two-fallback design earns its keep. The *test*
  project falls through to per-sheet reads and is ~6 × slower on reads as a result, which makes
  it the wrong place to measure — see **P14**.
- **The 100% mix invariant and the priced-at-zero check** exist precisely because the legacy
  workbook could be silently wrong, and the comments say so. Preserve them.
- **Every write follows the same seven-step shape**, and destructive maintenance functions
  are guarded by a `CONFIRM` constant that requires a code edit.

The findings below sit on top of genuinely thoughtful work.

---

## The register

| ID | Category | Sev | Effort | Location | Finding and consequence |
|---|---|---|---|---|---|
| **C1** | correctness bug | High | S | `rates.js:76,172`, `mixes.js:559`, `actuals.js:34`, `structure.js:147` | Update paths authorise the **incoming** foreign key, not the one on the row being edited. Sending `{id: <a rate outside my scope>, modellingId: <one I can edit>}` passes the check and rewrites that row, re-parenting it. Delete paths get this right — they read the row first. **Consequence: a scoped user can silently overwrite and steal rows belonging to segments they cannot see.** |
| **S1** | permissions & data safety | High | S | `auth.js:224` `canSeeHighLevelId_` | View-only scope grants edit-everywhere. With `forEdit=true` the function reads `scopes.edit`; if a person has `Can_View` rows but no `Can_Edit` rows, every `scopeAllows_` returns `null` and it falls through to `SCOPE_DEFAULT_ALLOW`, which ships `TRUE`. **Consequence: granting someone read-only access to one brand gives them edit rights over every brand. Dormant only because `Scope_Mapping` is empty; it fires the day scoping is switched on.** |
| **P1** | performance | High | M | `rates.js:351`, `mixes.js:78,388` | Three read-only screens each run the **entire forecast engine** — `loadEngineInput_` plus `computeModel_`, 8,352 detail rows — purely to display rates, then discard all but a lookup. They fire on every dropdown change on the Rates and Mixes screens. Two of the three swallow failure in a bare `catch`. **Consequence: the largest avoidable cost in the app, on the most-used screens, and a silent-failure path if the engine throws.** |
| **P2** | performance | High | M | `index.html:655` `reloadData`, `index.html:1957` `reloadInit` | Every save triggers a full reload of everything. `reloadData()` re-fetches all 18 tabs; `reloadInit()` re-fetches init *and* bulk. Editing one rate costs three round-trips, one of which re-reads the whole database. Several handlers fire `loadMixGrid()` and `reloadData()` concurrently. **Consequence: multi-second pause after every edit, and it worsens linearly as the data grows.** |
| **P3** | performance | High | M | `actuals.js:141` `importActuals` | Loops `saveActual` per row, so each row costs a lock acquire/release, a `setValues` and two `appendRow`s. A full staging import — roughly 600 segment-months — is about 1,800 API calls. **Consequence: certain 6-minute timeout, leaving a partial import with no record of where it stopped. `recordChangesBatch_` already exists and solves exactly this.** |
| **U1** | UX | High | M | `index.html` vs `rates.js:390`, `audit.js:254,297`, `actuals.js:141`, `snapshots.js:177`, `mixes.js:552`, `output.js:288`, `structure.js:96` | Eight built server capabilities are unreachable from the UI: `bulkRateChange` (the "Royal Mail +6% from April" feature, with a preview mode — arguably the highest-value thing in the codebase), `getRecordHistory` (so the entire Amends layer is invisible to users), `getRecentAudit`, `listSnapshots`, `importActuals`, `saveLetterParcelMix`, `verifyPublishedOutput`, `deleteHighLevelId`. **Consequence: work already paid for delivers nothing, and users do by hand what the server can already do in bulk.** |
| **M3** | maintainability | High | L | `enginetest.js`, `rates.js:481`, `mixes.js:644`, `audit.js:367` | No automated tests. "Tests" are `Logger.log` diagnostics run by hand from the editor, and three of them **write rows into the live spreadsheet and leave them there** (into year 2035 and ID 999999 so they cannot affect a forecast), with `cleanupTestData()` to sweep up afterwards. **Consequence: no regression safety net for the optimisation and UI work ahead — and the pure engine is perfectly testable under plain Node today.** |
| **M6** | maintainability | High | L | `index.html` | 2,313 lines of CSS, markup, state and eight screens in one file with no module boundaries. **Consequence: the principal obstacle to the UI overhaul; every change risks every screen.** |
| **M9** | maintainability | High | S | `setup.js:220` (was), `.clasp.json` | ✅ **Found and fixed during Phase 0.** `SEED_CONFIG` was a **top-level** array literal reading `ENGINE_VERSION`, a top-level `const` in `utils.js`. Apps Script evaluates files in project order in one shared scope, and top-level `const`/`let` sit in the temporal dead zone until their own file is evaluated — so this only ever worked because `utils.gs` happened to be ordered first. `clasp push` sorts alphabetically (`files.js:389`) and `"setup"` sorts before `"utils"`, so the repo's first push reordered the project and every entry point began throwing `ReferenceError: ENGINE_VERSION is not defined` at file-load time — before any function body runs, so `doGet`, all diagnostics, every `google.script.run` call and the weekly trigger all failed identically. **Consequence: the routine deploy command bricked the whole script project, and because the breakage is file *order* rather than file *contents*, `git revert` + `clasp push` could not recover it — the one failure mode in this repo that git cannot undo.** Fixed by making the seed a function (`seedConfigRows_()`) so the read happens after all files load, plus `"filePushOrder": ["utils.js"]` in `.clasp.json` as defence in depth. **The general rule this implies: never reference an identifier declared in another file from a file's top level — only from inside a function body.** It was the only such reference in the codebase; the other two `ENGINE_VERSION` readers (`Code.js:131`, `output.js:234`) are inside functions and were always safe. |
| **C2** | correctness bug | Med | S | `rates.js:452` → `rates.js:85` | Nested `withLock_`: `bulkRateChange` takes the script lock, then calls `saveBaseRate`, whose own `finally { releaseLock() }` frees it. **Consequence: everything after the first route in a bulk change runs unprotected, so a concurrent write can interleave halfway through.** |
| **C3** | correctness bug | Med | M | `engine.js:270` `checkMixInvariant_`, `validate.js:155` `ruleCoverage_` | A segment / regime / class with **no** `Mix_Method` rows at all is silently 0%: `checkMixInvariant_` skips it (`if (!mm) continue`, `if (!pct) continue`) and `RATE_MISSING` requires `methodMix > 0` to fire. `LP_COVERAGE` covers the letter/parcel layer; nothing covers the method layer. **Consequence: a segment can forecast zero with no error raised — the exact silent-understatement failure the invariant was written to prevent.** |
| **C4** | correctness bug | Med | M | `engine.js:355`, `engine.js:186` | `Dim_Surcharge.Applies_To` and `Apply_Order` are loaded into the engine input and never read. The engine always computes `base × (1 + Σpct) + Σamt`. **Consequence: surcharges cannot compound or be ordered, while the schema and the Admin form both invite the user to configure exactly that — a setting that silently does nothing.** |
| **C5** | correctness bug | Med | S | `Code.js:313–396`, `rates.js:92` | Client loaders do not filter by `Scenario_ID`, and `assertNoOverlap_` ignores it for base rates and surcharges. **Consequence: dormant today with one scenario; the moment a second exists the Rates screen interleaves both and scenario 2's periods block edits to scenario 1.** |
| **C6** | correctness bug | Med | S | `structure.js:518` `appendCopies` | Slices `batchGet` rows (which omit trailing empty cells) and writes them at full header width. **Consequence: a ragged source row throws "number of columns does not match" mid-write, after earlier tables have already been written — no rollback. Currently masked because `Updated_By`, the last column, is always populated.** |
| **C7** | correctness bug | Med | M | `output.js:151,178` | `publishOutput` clears `OUTPUT` and `OUTPUT_Detail` then rewrites them, with no transaction and no staging. **Consequence: a timeout or error between the clear and the write leaves the published forecast empty, with no `Calc_Runs` row recording that it happened.** |
| **P4** | performance | Med | S | `utils.js:384` `prewarmForWrite_` | Fetches `Audit_Log` — the fastest-growing tab in the file — on **every single write**, solely so `getNextId_` can find the maximum ID. **Consequence: every save gets slower forever, for a value that could come from a Property or a single-column read.** |
| **P5** | performance | Med | M | `Code.js:431`, `audit.js:304,332` | `Audit_Log` is read in full **twice per page load** — `loadStatus_` runs in both `initApp` and `loadAllAppData`, each calling `getPendingChangeCount_` — and `readRecentAudit_` reads the whole tab to return 50 rows. **Consequence: as the log grows this becomes the dominant page-load cost, and it grows with every action including every refused one.** |
| **P6** | performance | Med | M | `utils.js` §3 | No `CacheService` anywhere. `Dim_Reference`, `Dim_Carrier`, `Dim_Method`, `Dim_Surcharge`, `Dim_Calendar`, `Portal_Roles` and `Config` are effectively static and re-read on every call. `PropertiesService` is used only for the Metabase key. **Consequence: roughly half of `initApp`'s tab reads are re-fetching data that changes a few times a year.** |
| **P7** | performance | Med | M | `audit.js:121` `recordChange_` | Every single-row save costs two `appendRow` calls (Amends plus Audit) on top of the `setValues`. `recordChangesBatch_` exists and is used only by the grid paths. **Consequence: three sheet writes per user action where one would do.** |
| **P8** | performance | Med | S | `rates.js:54,280` | `readRow_` does a live `getRange().getValues()` even when the row is already in the execution cache, and `closePrecedingPeriod_` calls it inside a loop with a `setValues` and a `recordChange_` per matched row. **Consequence: an avoidable round-trip on every save, multiplied per row on period-closing writes.** |
| **P9** | performance | Med | M | `output.js:89` `publishOutput` | Runs the engine, validation, a ~9,000-row write and the audit inside one lock. **Consequence: the closest thing to the 6-minute ceiling today, and it blocks every concurrent save for its duration — other users get "The system is busy" after a 10-second wait.** |
| **P13** | performance | Med | M | `index.html:405` `renderAll` | Replaces the active screen's entire `innerHTML` on every interaction. **Consequence: focus, scroll position and any half-typed input are lost on every dropdown change; combined with P2 the screen visibly flashes after each save.** |
| **S3** | permissions & data safety | Med | M | `output.js:81`, `validate.js:46`, `audit.js:254,304` | Scope leaks in **return values**, not in the loaders. `previewOutput` returns rate movements for all segments; `runValidation` returns findings for all segments; `getRecordHistory` returns any record's full Amends history with only a `viewAudit` check and no row-level scope; `readRecentAudit_` — included in every `loadAllAppData` — returns audit rows for all segments with old and new values. `MODELLER` is non-`All_Access` and has both `runCalc` and `viewAudit`. **Consequence: the careful server-side scope filtering on the loaders is bypassed by four other endpoints.** |
| **S4** | permissions & data safety | Med | S | `snapshots.js:43` `weeklySnapshotJob` | No permission check at all, and no trailing `_`, so any domain user can invoke it via `google.script.run`. Deliberate and documented — a trigger has no signed-in user — and guarded only by a 55-minute rate limit. **Consequence: an ungated write endpoint. Defensible, but it should be a private `_` function with a thin gated wrapper the trigger calls.** |
| **S5** | permissions & data safety | Med | S | `utils.js:663` `requireMaintenance_` | Fails **open**: an empty `Permissions` tab returns `true` (intended, for bootstrap), and so does *any read error* on it. **Consequence: during the bootstrap window, or on a transient read failure, `cleanupTestData()` and `clearActuals()` — both destructive, neither `CONFIRM`-guarded — are open to any domain visitor.** |
| **S6** | permissions & data safety | Med | S | `Code.js:30` | `setXFrameOptionsMode(ALLOWALL)`. **Consequence: the portal can be embedded by any site, and one click in a framed page publishes the forecast or deletes a rate. No token, no re-authentication.** |
| **S7** | permissions & data safety | Med | S | `auth.js:154,164` | Every DENIED call appends a row to `Audit_Log` with no rate limit. **Consequence: any domain user can inflate the log cheaply, compounding P5 and eventually the file's cell limit.** |
| **S8** | permissions & data safety | Med | S | `setup.js:559,699`, `migrate.js:403`, `actuals.js:381` | Maintenance writers take no lock. `fixValidationWarnings` does a read-modify-write of the entire `Modelling_IDs` range; `migrate*`, `setup*`, `cleanupTestData` and `clearActuals` all write unlocked. **Consequence: a concurrent portal edit made during a maintenance run is silently overwritten.** |
| **S10** | permissions & data safety | Med | S | `auth.js:78`, `setup.js:221` | `BOOTSTRAP_OWNER_EMAIL` is a plaintext admin backdoor stored in the **Config tab**. Anyone with edit access to the spreadsheet grants themselves ADMIN by typing in a cell — as they equally can by editing `Permissions` directly. **Consequence: the spreadsheet's own sharing is the real trust boundary, and it is managed outside this code. Worth stating explicitly in the runbook rather than leaving implicit.** |
| **P14** | environment | Med | S | test script project's GCP project `736650435578`; `utils.js:289` `prewarmSheetCache_` | **Measured during Phase 0. Blocks the measurement half of Phase 0, not production.** The **test** project falls through to path 3 of `prewarmSheetCache_` — one round-trip per tab — because its GCP project does not have the Sheets API enabled (`REST batchGet returned HTTP 403: … has not been used in project 736650435578 before or it is disabled`). **Production is on path 2 (`rest`) and gets a measured 4.9 × speed-up.** Reading the same 8 tabs of the same spreadsheet: production **559 ms** batched, test **3,468 ms** per-sheet — test is ~**6 × slower on reads than production**. **Consequence: the test environment is not a valid place to measure read cost, so every read-heavy figure in BASELINE.md is inflated about sixfold and Phase 3's "under 1.5 s" decision rule would be applied to a number production never sees. Phases 3, 4 and 5 must be re-scored against production's cost model.** Cheapest fix: enable the Sheets API in the test project's own GCP project, so test's REST path behaves as production's does — no code or manifest change, and it makes test *match* production. Adding the Advanced Service to the repo's `appsscript.json` also works but moves *both* environments onto a third path, dragging production along with it, so it is the worse choice while the goal is a trustworthy baseline. If it is ever added, note the M9-shaped trap: enabling it via the editor's **Services +** menu writes `dependencies.enabledAdvancedServices` into the *remote* `appsscript.json`, and the next `clasp push` silently strips it — it must live in the repo. **Severity note:** first written up as High on the assumption this was a production performance defect; the production `diagnoseBatchGet()` run disproved that, leaving a contained measurement-integrity problem. |
| **S14** | permissions & data safety | Med | S | `utils.js:20` `SPREADSHEET_ID_FALLBACK`, `.clasp.json` | Nothing binds a script project to its spreadsheet. Since Phase 0 the target is a Script Property with a literal fallback, but **the fallback is production while the committed `scriptId` is test** — so a test project with no property set reads and writes production, and the only signal is running `showEnvironment()` and reading the output. The self-cleaning write diagnostics (`testRateWrite`, `testMixWrite`, `testStructureWrite`, `testAuditTrail`) would then append rows and audit entries to the production file. Partially mitigated by M3's convention of writing test rows at year 2035 / ID 999999 so they cannot reach a forecast. **Consequence: the environment is switchable but not enforced; the safe default is still the dangerous one. A `SPREADSHEET_ID`-to-`scriptId` assertion, or making the write diagnostics refuse when the ID came from the fallback, would close it.** Note also that `showEnvironment` is a non-underscore global gated only by `requireMaintenance_`, which fails open (S5), so it discloses the spreadsheet ID and file name on the same terms as every other maintenance function. |
| **U2** | UX | Med | M | `rates.js:39`, `index.html:587` | Inserting a rate change *before* an existing future period dead-ends. The error says "Close or shorten that period first", but the UI only ever sends `closePrevious:true` and offers no way to shorten a later period. **Consequence: the user has to open the spreadsheet and edit it by hand — bypassing validation and audit to complete a routine task.** |
| **U3** | UX | Med | M | `index.html:405,655` | Every save re-renders the whole screen (P13) and re-fetches everything (P2). **Consequence: the interface flashes and loses the user's place after every edit, which reads as instability.** |
| **U5** | UX | Med | M | `index.html:251,1111` | The modal closes on Escape but has no focus trap, so Tab escapes to the page behind it. Tables have no `scope` or caption. The SVG chart has an `aria-label` but no accessible data alternative. **Consequence: keyboard and screen-reader users cannot complete the core editing journeys.** |
| **U6** | UX | Med | S | `index.html:216` `showError` | Errors are injected as a box into the current screen, self-destruct after 8 seconds, and vanish on any re-render. Failures from `quiet` calls skip the save indicator entirely. **Consequence: a failed save can leave no visible trace, and the user believes their change was applied.** |
| **U8** | UX | Med | M | `index.html:983` | Cold chain is read-only, with on-screen guidance telling users to edit the `Mix_ColdChain` tab directly. **Consequence: the app instructs users to bypass its own validation and audit trail to change a value that feeds the forecast.** |
| **M1** | maintainability | Med | M | `rates.js`, `mixes.js`, `actuals.js` | The rate, surcharge, letter/parcel and actual save functions are near-identical 60-line blocks differing only in table and value validation. **Consequence: about 250 duplicated lines, and C1 had to be found five times because it is copied five times. One shared `saveDatedRow_` would fix it once.** |
| **M2** | maintainability | Med | S | `validate.js:36`, `output.js:263` | `runValidation` and `runValidationQuiet_` maintain duplicate copies of the rule list. **Consequence: adding a rule and forgetting one copy means it runs on demand but not before publish, or vice versa — the failure is invisible.** |
| **M4** | maintainability | Med | S | `migrate.js`, `output.js:288`, `enginetest.js` | `migrate.js` is 894 lines of one-time code that has already run, still shipped and still holding a live reference to the legacy workbook — on which `verifyPublishedOutput` and `runParityTest` also depend. **Consequence: 8% of the codebase is dead weight whose removal is blocked by an unanswered question (see ARCHITECTURE.md §7).** |
| **M5** | maintainability | Med | M | various | Dead or vestigial: `include()` (`Code.js:40` — `index.html` has no scriptlets); `Permissions.Tab_Visibility` (read, sent to the client, never used); `Applies_To`/`Apply_Order` (C4); `Rate_CTS` (computed, written, never read); `OUTPUT_Detail` (written every publish, never read back by the app); `FX_Rates` (768 rows, explicitly "not used by the engine"); `Scenarios` (every client call site passes a hardcoded `1`). **Consequence: each one is a thing a new maintainer must investigate and discard.** |
| **C8** | correctness bug | Low | S | `audit.js:26` `logAudit_` | Computes `Log_ID` as max+1 then `appendRow`, frequently **outside** any lock — notably on every DENIED path. **Consequence: concurrent executions can produce duplicate `Log_ID`s. Harmless today since nothing keys on it, but it is presented as an identifier.** |
| **C9** | correctness bug | Low | S | `validate.js:135` | `TERMINAL_MID_MONTH` rebuilds a date with `new Date(dayNum × 86400000)` — a UTC instant then read through local accessors. **Consequence: correct for Europe/London (offset ≥ 0); on a negative-offset project timezone the rule would misfire by one day.** |
| **C10** | correctness bug | Low | S | `auth.js:184,197`, `utils.js:593` | `_hlMetaCache_`, `_midMetaCache_`, `_permsCache_` and `_configCache_` are never cleared when their source tab is written, although `invalidateSheetCache_` clears the row cache. `repairConfigValues` resets `_configCache_` manually, showing the author was aware. **Consequence: within one execution, code after a write can read stale metadata — e.g. a just-created High Level ID is absent from `hlMeta_()`.** |
| **C11** | correctness bug | Low | S | `audit.js:332` `getPendingChangeCount_` | Counts audit rows **per changed field**, and counts changes that cannot move the forecast (a reference label edit, a note). **Consequence: the "N changes unpublished" badge overstates, so users learn to ignore it — which defeats its purpose.** |
| **P10** | performance | Low | S | `structure.js:640` `getReferenceUsage` | Calls `referenceUsage_` once per reference value, each re-scanning up to six full tables. O(values × rows), on every visit to Admin → Reference lists. **Consequence: a slow Admin screen that gets slower as rate tables grow; one pass could build all counts.** |
| **P11** | performance | Low | S | `mixes.js:291,544` | `clearMixWindowBatch_` and `clearLetterParcelWindow_` write the **entire** data range back to change a handful of rows. **Consequence: acceptable at 374 rows, wasteful and increasingly risky as the table grows.** |
| **P12** | performance | Low | S | `utils.js:19`, `output.js:51`, `validate.js:311` | `PERF_LOG_ENABLED = true` in production, and `previewOutput`/`runValidation` build 20+ `Logger.log` lines per call that no user ever sees. **Consequence: shipped debug output on every request; small but pure waste.** |
| **S9** | permissions & data safety | Low | S | `utils.js:14`, `migrate.js:20` | ✅ **Fixed in Phase 0.** Two production spreadsheet IDs hardcoded and committed to the repository. **Consequence: not secret in themselves, but they pin the code to one environment — there is no way to point a copy at a test sheet without editing source.** Both are now Script Properties (`SPREADSHEET_ID`, `SOURCE_SPREADSHEET_ID`) with the original literals as fallbacks, read through `spreadsheetId_()` / `sourceSpreadsheetId_()`. The literals are still committed, deliberately, so an unconfigured project keeps working — see **S14** for the residual risk that creates. |
| **S11** | permissions & data safety | Low | S | `Code.js:40` `include` | Exposed to `google.script.run` and returns the content of any HTML file in the project. Also dead: `index.html` contains no scriptlets, so `createTemplateFromFile` could be `createHtmlOutputFromFile`. **Consequence: an ungated read endpoint with no current use.** |
| **S13** | permissions & data safety | Low | S | `index.html:258,293` | `openForm`'s `cfg.intro` and `field.hint` are injected as raw HTML. Every current caller passes literals or `esc()`-wrapped values, so **there is no live XSS**. **Consequence: an unguarded seam one careless caller away from injecting stored data — a Note, a Source_Ref or a reference label — into the DOM of a page that can call `google.script.run` as the user.** |
| **U4** | UX | Low | S | `index.html:617,651,1415` | Native `confirm()` for destructive actions, sitting beside a carefully built custom modal. **Consequence: inconsistent, and unstyleable inside the Apps Script iframe.** |
| **U7** | UX | Low | S | `index.html:756,873` | `2028-12-31` is hardcoded as the default "To" date in both mix editors. **Consequence: correct only for the current 36-month horizon; extending the horizon silently leaves new periods ending early.** |
| **U9** | UX | Low | S | `auth.js:69`, `index.html:160` | `Permissions.Tab_Visibility` is read, sent to the client as `user.tabVisibility` and never used — the nav filters on capability instead. **Consequence: a governance column administrators may believe controls something.** |
| **M7** | maintainability | Low | S | `output.js:240`, `migrate.js:726`, `audit.js:339,352` | Duplicated helpers: `currentUserEmail_` vs `getActiveEmail_`; `logAction_` vs `logAudit_`; the serial-to-milliseconds conversion open-coded with the literal `25569` instead of `SERIAL_EPOCH_OFFSET`. **Consequence: the date-handling discipline described in ARCHITECTURE.md §5.4 has three exceptions, which is how that discipline erodes.** |
| **M10** | maintainability | Low | S | `Code.js:477` `showPortalUrl` | `ok = url.indexOf('/exec') >= 0`, so a project served from its `/dev` head URL is reported as *"not a web app deployment"* and never reaches `READY`. A `/dev` URL is a perfectly valid way to run the portal during development — it serves HEAD, which is what you want after a `clasp push` — and the test project has no versioned deployment at all (confirmed 2026-08-14). Two further wrinkles: `ScriptApp.getService().getUrl()` can return the `/dev` URL when called from the editor even where an `/exec` deployment exists, so the check is not a reliable discriminator either way; and **Deploy → Manage deployments** is the only authoritative source for the real `/exec` address. **Consequence: the function reports a healthy development project as broken, which sent a verification run looking for a deployment problem that did not exist. It should report both URLs and describe them, rather than failing anything that is not `/exec`.** |
| **M8** | maintainability | Low | S | `docs/SHEET_STRUCTURE.md` | A dated snapshot (2026-08-13) with no regeneration script. **Consequence: it will drift from the sheet and be trusted anyway.** |

---

## Detail on the high-severity findings

### C1 — authorising the payload's foreign key instead of the row's

`saveBaseRate` reads `modellingId` from the request, checks the caller may edit *that*
route, then looks up the row by `p.id` and overwrites it:

```js
const modellingId = safeInt(p.modellingId);
assertCanEditModellingId_(perms, modellingId);   // checks the INCOMING id
…
rowIndex = findRowById_(t.sheet, C.Rate_ID, p.id);
before = readRow_(t.sheet, rowIndex, width);      // row could belong to anyone
row[C.Modelling_ID] = modellingId;                // and is re-pointed here
```

The row's existing `Modelling_ID` is never checked. The delete paths do it correctly — they
read the row first, then `assertCanEditModellingId_(perms, before[C.Modelling_ID])` — so the
correct pattern already exists in the same file. The fix is to authorise **both** the
current owner and the proposed new owner on update. Because M1 has copied this shape five
times, fixing it once in a shared helper fixes all five.

### S1 — the scope fall-through

```js
const set = forEdit ? scopes.edit : scopes.view;
const direct  = scopeAllows_(set, 'HIGH_LEVEL_ID', String(hlId));  // null if no rule
const byBrand = scopeAllows_(set, 'BRAND', meta.brand);            // null if no rule
const byGeo   = scopeAllows_(set, 'GEO', meta.geo);                // null if no rule
return (direct === null && byBrand === null && byGeo === null)
       ? configBool('SCOPE_DEFAULT_ALLOW', true) : false;
```

`scopeAllows_` returns `null` for "no rule of this type", which the caller reads as
"undecided, fall back to the default". That is right for a person with no scope rows at all.
It is wrong for a person whose `edit` set is empty *because they were deliberately given
view-only access* — their empty edit set is a decision, not an absence. The distinction the
code needs is "this person has scope rows, therefore silence means no", and `scopes.any`
already records it — it is just not consulted separately per direction.

### P1 — running the whole engine to render a table

`getResolvedByMonth`, `getMethodMixGrid` and `getLetterParcelGrid` all do this:

```js
const input  = loadEngineInput_(scenarioId);   // reads 10 tabs
const result = computeModel_(input);           // computes all 8,352 detail rows
// …then use a handful of rows for one segment and throw the rest away
```

In two of the three it is wrapped in `try { … } catch (e) { /* rates are a nicety here */ }`,
so if the engine throws the screen renders with zeros and no error. Each fires on every
dropdown change. The engine is pure and already accepts a filtered input — the same result
could come from computing one segment, or from reading the already-published `OUTPUT_Detail`
(which is currently written every publish and never read back — see M5).

### U1 — built and unreachable

`bulkRateChange` (`rates.js:390`) implements scoped bulk uplifts by carrier, by segment or by
explicit route list, in percentage / absolute / set-to modes, defaulting to preview and
returning a per-route before-and-after plan. It respects row-level scope, silently skipping
routes the caller cannot edit. Nothing in `index.html` calls it. `getRecordHistory`
(`audit.js:254`) is the only reader of the eight `*_Amends` tables, which are written on
every change and never shown to anyone.

---

## Cross-cutting themes

Four patterns explain most of the register:

1. **Authorisation is checked against the request, not the record** (C1, S3). The loaders
   filter correctly; the mutations and four read endpoints do not.
2. **Reads are all-or-nothing.** There is no way to ask for part of a table, so every screen
   re-reads everything (P2, P5) and every screen that needs one rate runs the whole model
   (P1).
3. **The batch primitives exist but are used unevenly.** `recordChangesBatch_` and
   `prewarmSheetCache_` are exactly right, and the single-row paths ignore them (P3, P7).
4. **Copy-paste instead of extraction** (M1, M2, M7) — which is also why C1 is a five-place
   bug rather than a one-place bug.

---

## What I have not verified

The performance severities are reasoned from API call counts, not measured. `perfReport()`
is already instrumented and returned in every payload; one `testApiPayload()` run from the
editor and one browser network trace would confirm or demolish P1, P2 and P5 in about ten
minutes, and should happen before any of that work is scheduled.

The severities on S1, S2/C1 and S3 assume row-level scoping is **intended to be switched
on**. `Scope_Mapping` is currently empty, so none of them can be exploited today. If scoping
is never going to be used, all three drop a severity and the `Scope_Mapping` machinery
becomes a candidate for M5.

Full list of open questions and inferred-versus-observed: [ARCHITECTURE.md §7](ARCHITECTURE.md#7-what-i-inferred-rather-than-read).
