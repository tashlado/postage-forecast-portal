# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Google Apps Script web app (the "Postage Forecast Portal") that replaces a
spreadsheet-formula forecasting workbook (originally an 8,352-row `Modelling`
tab driven by SUMPRODUCT/SUMIFS) with a normalised data model, a pure
calculation engine, role/scope-based access control, and a full audit trail.
All data lives in one Google Sheet (resolved by `spreadsheetId_()` in `utils.js` —
the `SPREADSHEET_ID` Script Property if set, else `SPREADSHEET_ID_FALLBACK`); the app
itself is a single HTML file served by `doGet`, backed by `.js` files that are
really Apps Script `.gs` files (renamed for this repo/editor).

This is clasp-managed: `.clasp.json` points at the Apps Script project
(`scriptId`), and `.js` files here map 1:1 to script files there.

> **Which project does `clasp push` hit, and which spreadsheet does it then write?**
> These are two separate questions and the answers pull in opposite directions.
> The `scriptId` committed in every worktree's `.clasp.json` is the **TEST** script
> project (`1Br6gJdV…`); production (`1JHYSult…`) appears in **no** `.clasp.json`.
> But `SPREADSHEET_ID_FALLBACK` **is the production spreadsheet**, so a test project
> with no `SPREADSHEET_ID` property set reads and writes **production data**
> (FINDINGS.md S14). The safe default is the dangerous one, and the only signal is
> `showEnvironment()`. IDs, the trap, and the test/production performance gap are all
> in **`docs/TEST_PORTAL.md`** — read it before running anything that writes.

## Commands

There is no build step, package manager, bundler, or test runner in the
conventional sense — this is Apps Script, deployed via `clasp`.

- `clasp push` — push local files to the Apps Script project
- `clasp pull` — pull from the Apps Script project (rarely needed; this repo is push-of-record)
- `clasp open-script` — open the project in the Apps Script editor (clasp 3.x renamed
  `open`; `clasp open-web-app` opens the deployed portal, `clasp status` still works)

"Tests" are diagnostic functions run manually from the Apps Script editor
(Run menu), not from a CLI. They log to `Logger.log` and return a summary
object. Key ones, roughly in the order you'd use them on a fresh setup:

- `showEnvironment()` (`utils.js`) — **run this first**: which spreadsheet this project is pointed at, and whether that came from a Script Property or the committed fallback
- `pointThisProjectAtTestSpreadsheet()` / `clearTestSpreadsheetOverride()` (`utils.js`) — set or remove the `SPREADSHEET_ID` override. Both refuse to run anywhere but the test script project, and the setter refuses if the ID it is given is production's
- `setupAll()` (`setup.js`) — idempotent: creates all 30+ tabs and seeds reference data
- `migratePreflight()` / `migrateAll()` / `migrateVerify()` (`migrate.js`) — one-time load from the original workbook (`SOURCE_SPREADSHEET_ID` property, else `SOURCE_SPREADSHEET_ID_FALLBACK`)
- `runEngineUnitTests()` / `runParityTest()` (`enginetest.js`) — engine correctness and parity against the original workbook's Output tab
- `testMyPermissions()` (`auth.js`) — shows what the current user can see/do
- `testRateWrite()`, `testMixWrite()`, `testStructureWrite()`, `testAuditTrail()` — round-trip create/refuse/update/delete checks for their respective areas, self-cleaning
- `diagnoseBatchGet()` — which bulk-read path (Advanced Sheets Service / REST / per-sheet) is active, and the speed-up
- `showPortalUrl()` — prints the actual `/exec` web app URL and sanity-checks the deployment
- `testApiPayload()` — times `initApp()` + `loadAllAppData()` end-to-end and prints payload sizes
- `runEnginePreview()` (`engine.js`), `previewOutput()` (`output.js`) — run the model without writing anything
- `diagnoseActuals(hlId)` — inspects imported actuals for a segment, flags rows where cost parsing failed

Functions ending in `_` are private (Apps Script hides nothing, but the
convention is enforced by review, not the runtime) and are never exposed to
the client. Maintenance-only functions (setup, migration, diagnostics) are
named *without* a trailing underscore so they show up in the Apps Script
editor's function picker, but self-gate via `requireMaintenance_()`.

## Architecture

### Layering

```
index.html          single-page client: state object S, google.script.run calls
Code.js              doGet, initApp()/loadAllAppData() — the only two payload shapes
auth.js              identity + permissions + scope filtering (re-derived every call)
utils.js             SHEET/COL/TABLES schema, caching, coercion, dates, config, perf
engine.js            computeModel_() — pure forecast calculation
structure.js         High Level IDs / Modelling IDs / dimension CRUD
rates.js             Rate_Base / Rate_Surcharge CRUD, bulk rate changes
mixes.js             Mix_Method / Mix_LetterParcel grid CRUD (100% invariant)
output.js            previewOutput() / publishOutput() — engine -> OUTPUT tabs
validate.js          rule pack (ERROR/WARN/INFO), run standalone or pre-publish
actuals.js           actuals CRUD, extract import, forecast-vs-actual comparison
audit.js             Audit_Log (per-field) + *_Amends (full-row snapshots)
snapshots.js         weekly forecast snapshots, stored only when changed
setup.js             idempotent tab creation + reference data seeding
migrate.js           one-time load from the legacy workbook
enginetest.js        parity test against the legacy workbook, engine unit tests
```

### The spreadsheet is the schema, once

`TABLES` in `utils.js` is the single source of truth for every tab: sheet
name, column headers (in order), and whether it has history (`amends: true`).
`SHEET` (table key -> sheet name), `COL` (table key -> column name -> 0-based
index) and `TABLE_BY_SHEET` are all derived from `TABLES` at load time. Any
table flagged `amends: true` automatically gets a matching `_AMENDS` table
generated (same headers, prefixed with `Amend_ID/Amend_TS/Amend_By/Amend_Type`).
Add a column in `TABLES` and `setupCreateAllTabs()` will append it to the real
sheet by name — never by position — so column order in the sheet and in code
cannot drift apart. Read/write code should always go through `COL.<TABLE>.<Header>`,
never a literal column index.

### Two-phase page load

`doGet` serves `index.html`, which calls `initApp()` first (permissions,
dropdown lists, structure — fast) and only then `loadAllAppData()` (rates,
mixes, output, validation, history — the bulk). This is deliberate: the page
frame renders in ~1s instead of blocking on everything. Don't merge these two
into one call.

### Trust boundary: the server, not the browser

Every entry point the client can call re-derives permissions from the
spreadsheet (`requirePermissions_()`) rather than trusting anything the page
sends. Two layers, both in `auth.js`:

1. **Portal_Roles** — the ceiling: what a *role* may ever do (`write`,
   `editRates`, `editMixes`, `editStructure`, `runCalc`, `publishOutput`,
   `manageUsers`, `viewAudit`). A role with `All_Access` skips scoping
   entirely.
2. **Scope_Mapping** — the slice: which High Level IDs / brands / geos /
   carriers a *person* may view or edit. Filtering happens server-side before
   data is serialised (`visibleHighLevelIds_`, `visibleModellingIds_`) — a
   scoped-out row is never sent to the browser to be hidden client-side.

Every write function follows the same shape: prewarm the sheets it will
touch (one HTTP call) -> `requirePermissions_()` -> capability check
(`requireEditRates_` etc.) -> row-level scope check
(`assertCanEditHighLevelId_` / `assertCanEditModellingId_`) -> validate ->
`withLock_()` -> write -> `recordChange_()` -> return. Follow this shape for
new write functions rather than inventing a new order.

### The calculation engine is pure

`computeModel_()` in `engine.js` takes plain objects in and returns plain
objects out — no `SpreadsheetApp`, no `Session`, nothing Google-specific.
`loadEngineInput_()` is the only thing that touches the sheet; it turns rows
into the shape the engine wants. This split is what makes `runParityTest()`
possible (replay the same inputs, diff against the legacy workbook) and lets
`previewOutput()` run the whole model without writing anything.

The maths (see the header comment in `engine.js` for the full spec):
- `base(d,m)` — point-in-time value on the 1st of the month
- `fuelPct`/`otherAmt` — day-weighted average across the month (proration matters — see `DAY-WEIGHTED` vs `POINT_IN_TIME` surcharges)
- `rate = base * (1 + fuelPct) + otherAmt`
- `methodMix = mixCC * ccShare + mixAmbient * (1 - ccShare)`
- `lpMix = letterMix` if the route is LETTER, else `1 - letterMix`
- `contribution = rate * methodMix * lpMix`, summed per High Level ID per month into `OUTPUT`

The 100%-mix invariant (`checkMixInvariant_`) and the "priced at zero while
carrying volume" check (`RATE_MISSING`) exist because the legacy workbook
could silently understate the forecast with no error anywhere — preserve
these checks when touching the engine.

### Dates are the recurring hazard

Google Sheets serialises dates as numbers (days since 1899-12-30) when read
via `batchGet`, not as `Date` objects. `normaliseDate()` / `dateKey()` /
`dayNum()` in `utils.js` and `engine.js` are the only places that should ever
interpret a raw sheet value as a date — route everything through them.
`isDate_()` is duck-typed (checks for a working `getTime()`) rather than
`instanceof Date`, because a `Date` crossing an execution boundary can belong
to a different JS realm.

### Bulk reads, per-execution cache

`prewarmSheetCache_()` fetches many tabs in one HTTP call, trying the
Advanced Sheets Service, then the raw REST API (via the script's own OAuth
token), then falling back to per-sheet reads — see the comment in `utils.js`
for why REST exists as a fallback (Workspace admins can allowlist away the
Advanced Service). `getAllData_()` caches per sheet for the lifetime of one
execution (`_sheetDataCache_`); call `invalidateSheetCache_(sheetName)` after
any write to that sheet before reading it again in the same execution.

### History: two records, two questions

`Audit_Log` (one row per changed *field* — cheap, drives the History screen)
and `<Table>_Amends` (full-row snapshot per write — lets you reconstruct "the
rate card exactly as it stood on 3 March"). `recordChange_()` writes both and
is the one call every single-row write path should make;
`recordChangesBatch_()` is the equivalent for grid/bulk writes (mix grids,
segment copies) — it turns what would be hundreds of `appendRow` calls into a
handful of ranged `setValues` calls. Amends are matched by column *name*, not
position, so a new column added to a source table lands correctly in its
Amends table with no extra code.

### Client is one HTML file, no framework

`index.html` holds a global `S` state object, a `TABS` array driving the nav,
and a `call(fn, args, onOk, onErr, quiet)` wrapper around
`google.script.run` that also drives the "Saved/Saving…/Not saved" indicator
in the header. There's a custom `openForm()` modal builder (replacing
`window.prompt`, which can't label fields) used everywhere data is edited.
Any new client-side write should go through `call()`, not raw
`google.script.run`, to keep the save indicator consistent.

## Working in this repo

- **Never reference an identifier declared in another file from a file's top
  level** — only from inside a function body. Apps Script evaluates files in the
  project's file order in one shared scope, and a top-level `const`/`let` is in
  the temporal dead zone until its own file has been evaluated, so a top-level
  read of another file's constant throws `ReferenceError` depending purely on file
  order. `clasp push` orders files alphabetically, so `setup.gs` loads before
  `utils.gs`. This bricked the whole project once (FINDINGS.md M9): the error
  fires at load time, before any function runs, so every entry point including
  `doGet` fails, and `git revert` cannot fix it because the breakage is file
  *order*, not file *contents*. `.clasp.json` pins `utils.js` first as defence in
  depth, but the rule is what actually protects you.
- Changing a table's shape means editing `TABLES` in `utils.js`, not the
  sheet directly — `setupCreateAllTabs()` reconciles the real sheet to match.
- Never hardcode a brand, geo, carrier, method, or surcharge type in code —
  these come from `Dim_Reference` (via `assertInReferenceList_`), `Dim_Carrier`,
  `Dim_Method`, `Dim_Surcharge`. Adding a new one is a data operation, done
  through `saveReferenceValue`/`saveCarrier`/`saveMethod`/`saveSurchargeType`.
- `LETTER_PARCEL` and `TEMP_REGIME` reference lists are locked
  (`LOCKED_REFERENCE_LISTS`) because the engine branches on their literal
  codes (`LETTER`/`PARCEL`, `CC`/`AMBIENT`) — don't remove entries from these,
  and don't add new branch values without also updating `engine.js`.
- Soft-delete is the norm for user-facing records (`Active = false`), not row
  deletion — history and audit trail depend on the row surviving.
- `Config` tab values are always stored as text but Sheets will silently coerce
  anything date-shaped when written — always read config dates through
  `configDate()`, never `getConfig_()` directly for a date-typed key.
- The `docs/SHEET_STRUCTURE.md` snapshot describes the current tab layout
  (headers, row/column counts) as of the date in its first line — regenerate
  rather than hand-edit it if the schema changes.

## Constraints

- No external network transmission of data. Everything stays in Google Workspace.
- 6-minute execution limit per invocation; Sheets API calls are the main cost.
- Batch reads/writes (getValues/setValues over ranges). Never loop getValue/setValue.

## House rules

- Ask before changing the sheet schema — column moves break formulas elsewhere.
- Preserve existing function names that are wired to triggers or menu items.
- This handles postage cost forecasting data — flag anything that looks like it touches employee or customer personal data before changing it.
