# Postage Forecast Portal — architecture

An as-built description of the application, written from a complete read of every `.gs`
file, `index.html`, `appsscript.json` and `.clasp.json` at commit `e299923`.

This document describes **what is there**, not what should be there. Defects and
opportunities are catalogued separately in [FINDINGS.md](FINDINGS.md). Where I inferred
intent rather than reading it, §7 says so explicitly.

---

## 1. What the application does

A postage-cost forecasting portal. It exists to answer one question, 612 times over:

> For a given business segment, in a given month, what will the blended postage cost per
> order be?

Two levels of identity carry the model:

- A **High Level ID** is one business segment — brand × geography × treatment type ×
  weight-loss split (e.g. `MEDEXPRESS_GB_WL_MOUNJARO`). There are 17.
- A **Modelling ID** is one route within a segment — carrier × method × letter-or-parcel
  (e.g. Royal Mail / RoyalMailTracked24 / PARCEL). There are 232.

The forecast for a segment-month is the sum, across that segment's routes, of each route's
rate multiplied by the share of volume it carries. 17 segments × 36 months = 612 output
rows; 232 routes × 36 months = 8,352 detail rows.

It replaces an 8,352-row `Modelling` tab in a spreadsheet workbook that computed the same
answer with SUMPRODUCT and SUMIFS. The stated motivation, visible throughout the comments,
is that the workbook could be **quietly wrong**: a delivery mix that summed to 90% instead
of 100% produced a plausible forecast that was 10% low, with no error anywhere. Several
design decisions only make sense in that light and are called out below.

### The user journeys

**1. Change a rate and see the forecast move.** *Rates tab.* Pick a segment, pick a route.
Two tables — base rate periods and surcharge periods — with Add / Edit / Delete. Adding a
rate change from a date automatically closes whichever period was running, so "Royal Mail
goes up on 1 April" is one action, not two edits. Below them, a **Resolved by month** panel
shows base, surcharge %, surcharge amount, rate per parcel, mix and contribution for every
month, with changed months highlighted. That panel is the point of the whole rebuild: in
the workbook, a fuel change on the 4th of the month resolved to a blended figure buried
inside a 2,500-row SUMPRODUCT that nobody could check. Here it is a column you can read.

**2. Change the mix.** *Mixes tab*, three sub-screens.

- **Delivery method mix** is edited as a **grid**, never row by row. For a segment, a
  temperature regime (cold chain or ambient) and a class (letter or parcel), every route's
  percentage is shown together with a running total and a Save button that stays disabled
  until the total is 100%. This is a deliberate structural choice, not a UI preference: you
  cannot get from one valid 100% state to another one row at a time without passing through
  an invalid state, so the whole set is validated and written together or not at all. A
  "Normalise to 100%" button scales entered values, giving the largest row the rounding
  residual.
- **Letter / parcel split** — a two-row grid where typing in one row adjusts the other.
  Only the letter share is stored; parcel is the remainder, so the two can never disagree.
  Existing overlapping periods are trimmed rather than the save being refused.
- **Cold chain share** — read only, with a note directing users to edit the sheet directly.

**3. Check, then publish.** *Checks tab* runs a rule pack returning ERROR / WARN / INFO.
*Output tab* offers **Preview changes** (runs the model, diffs against the published
`OUTPUT`, writes nothing) and **Publish** (runs the model, validates, then overwrites
`OUTPUT` and `OUTPUT_Detail` and records a `Calc_Runs` row). An ERROR blocks publishing
unless `VALIDATION_BLOCKS_PUBLISH` is turned off in Config. Below is the full published
grid, every segment against every month, with changed cells shaded and peak months tinted.

**4. Compare against reality.** *Actuals tab.* Enter orders and total spend for a
segment-month and the blended rate is derived as spend ÷ orders, so the three figures
cannot drift apart. A rate on its own is accepted but forfeits the volume-weighted
comparison. The screen shows mean forecast, mean actual, variance, a volume-weighted actual
rate, and a month-by-month table with rows outside the configured tolerance shaded.

**5. Administer the model as data.** *Admin tab*, four sub-screens: carriers and methods;
High Level and Modelling IDs (with **Clone routes** and **Copy rates & mixes** between
segments, both previewing before writing); surcharge types; reference lists. Nothing in the
code hardcodes a brand, geography, carrier, method or surcharge type — adding a courier is
a data operation. Reference codes become immutable once anything points at them; labels are
always editable. `LETTER_PARCEL` and `TEMP_REGIME` are locked entirely because the engine
branches on their literal values.

**6. See what changed.** *History tab.* **Forecast over time** lists weekly snapshots and
compares any two (or one against the live forecast), reporting which segments moved and by
how much. **Change log** shows recent activity from the audit trail.

**Dashboard.** Forecast status, validation counts, a forecast-vs-actual quarterly bar chart
(hand-rolled inline SVG — no chart library), and a quarterly summary table where clicking a
row jumps to that segment's rates.

---

## 2. Entry points

This inventory is complete — it was produced by grep, not assumption.

| Entry point | Kind | File | Notes |
|---|---|---|---|
| `doGet(e)` | Web app | `Code.js:24` | The **only** HTTP entry point. Serves `index.html`; falls back to a self-describing status page if that file is missing. |
| `weeklySnapshotJob()` | Time-driven trigger | `snapshots.js:43` | Monday 06:00 Europe/London. Installed manually by `installWeeklySnapshot()`, not automatically. |
| ~35 named globals | `google.script.run` | various | Any global function whose name does not end in `_` is callable from the browser. |

**There is no `onOpen`, no `onEdit`, no `onInstall`, no `doPost`, and no custom menu.**
`SpreadsheetApp.getActive()` is never called — the spreadsheet is always reached by
`openById`, so the script is not bound to the sheet.

Two consequences worth stating plainly:

- **Direct edits to the Sheet bypass everything.** No validation, no 100% mix check, no
  lock, no audit row. The audit trail is complete only for changes made through the portal.
- **`ScriptApp.getProjectTriggers()` is the only trigger inventory.** If nobody has run
  `installWeeklySnapshot()` on this deployment, the weekly snapshot simply does not happen;
  `showSnapshotTriggers()` reports whether it did.

### Deployment posture

```json
"webapp": { "access": "DOMAIN", "executeAs": "USER_DEPLOYING" }
```

The script runs as the deploying account, so **the spreadsheet does not need to be shared
with users** — the `Permissions` tab is the entire access model. Identity comes from
`Session.getActiveUser().getEmail()`, which returns a real address here only because access
is restricted to the same Workspace domain. If it ever returns empty, `getUserPermissions_`
finds nobody and refuses access, which is the correct direction to fail.

OAuth scopes requested: `spreadsheets`, `script.external_request` (for the REST batch-read
fallback and the Metabase probes), `userinfo.email`, `script.scriptapp` (triggers).

---

## 3. Control flow

### 3.1 Page load — two calls, deliberately

```
browser                                server
───────                                ──────
boot()
 │
 ├─ call('initApp')            ──►     prewarmSheetCache_(14 tabs)      1 HTTP call
 │                                     requirePermissions_()            re-read from sheet
 │                                     recordLogin_()                   throttled 15 min
 │                                     visibleHighLevelIds_ / visibleModellingIds_
 │                                     9 × load*_ + loadStatus_
 │                             ◄──     { user, reference, carriers, methods,
 │                                       surchargeTypes, calendar, highLevelIds,
 │                                       modellingIds, scenarios, config, status, perf }
 │
 ├─ render nav + frame                 page is now visible (~1s)
 │
 └─ call('loadAllAppData')     ──►     prewarmSheetCache_(18 tabs)      1 HTTP call
                                       requirePermissions_()            fresh execution
                                       9 × load*ForClient_ (scope-filtered)
                               ◄──     { rateBase, rateSurcharge, mixMethod,
                                         mixLetterParcel, mixColdChain, output,
                                         actuals, snapshots, validation,
                                         recentAudit, status, perf }
```

The split exists so the frame paints in about a second instead of blocking on the bulk. The
header comment in `Code.js` is explicit that these two should not be merged.

**Each `google.script.run` call is a separate server execution.** Module-level state —
`_ssCache_`, `_sheetCache_`, `_sheetDataCache_`, `_nextIdCache_`, `_permsCache_`,
`_configCache_`, `_hlMetaCache_`, `_midMetaCache_` — is built up during one execution and
discarded at the end. That is exactly the lifetime the design wants, and it is why
permissions can be re-derived on every call without the cost being unbearable.

### 3.2 The shape of every write

Every write function follows the same seven steps in the same order. This consistency is
the codebase's best structural property and any new write should copy it:

```js
prewarmForWrite_([sheets…])       // one batched read of the tabs this write will touch
const perms = requirePermissions_();   // identity + role, re-read from the sheet
requireEditRates_(perms);              // capability gate — the role ceiling
assertCanEditModellingId_(perms, id);  // row-level gate — this person's slice
validateDates_(…); /* value checks */  // refuse bad input before taking the lock
return withLock_(function () {         // ScriptLock, 10-second wait
  invalidateSheetCache_(sheet);
  /* … setValues … */
  recordChange_(table, id, before, after, type);   // Amends snapshot + Audit_Log rows
  return { ok: true, … };
});
```

### 3.3 Trust boundary

The comments state the model outright: **the browser is assumed hostile.** Nothing the page
sends about who the user is or what they may do is believed. Two layers, both in `auth.js`:

1. **`Portal_Roles` — the ceiling.** What a *role* may ever do: `write`, `editRates`,
   `editMixes`, `editStructure`, `runCalc`, `publishOutput`, `manageUsers`, `viewAudit`. A
   role flagged `All_Access` skips layer 2 entirely. Seeded roles are ADMIN, MODELLER,
   ANALYST, VIEWER.
2. **`Scope_Mapping` — the slice.** Which High Level IDs, brands, geographies or carriers a
   *person* may view and edit, with separate view and edit grants.

Filtering happens **server-side before serialisation**: `visibleHighLevelIds_` and
`visibleModellingIds_` are computed first and every `load*ForClient_` skips rows outside
them. A scoped-out row is never sent to the browser to be hidden with CSS. The comment on
this is worth preserving verbatim — *"anything else is theatre, because the payload is one
keystroke away in developer tools."*

The client's own permission checks (`can()` filtering the nav, `canEdit` disabling buttons)
are cosmetic conveniences. Every one of them is backed by a server-side check; I traced all
of them.

### 3.4 Call graph — significant paths

```
doGet ──► index.html ──► google.script.run
│
├─ initApp ─────────────► prewarmSheetCache_ ─► Sheets.batchGet │ REST │ per-sheet
│                         requirePermissions_ ─► getUserPermissions_ ─► loadScopes_
│                         loadReferenceLists_ / loadCarriers_ / loadMethods_ /
│                         loadSurchargeTypes_ / loadCalendarForClient_ /
│                         loadHighLevelIdsForClient_ / loadModellingIdsForClient_ /
│                         loadScenarios_ / loadStatus_ ─► getPendingChangeCount_
│
├─ loadAllAppData ──────► load{RateBase,RateSurcharge,MixMethod,MixLP,MixCC,Output,
│                              Actuals,Snapshots,Validation}ForClient_
│                         + readRecentAudit_ + loadStatus_
│
├─ saveBaseRate ────────┐
│  saveSurchargeRate    │  closePrecedingPeriod_ ─► recordChange_
│  saveHighLevelId      ├─ assertNoOverlap_ / duplicate checks
│  saveModellingId      │  findRowById_ ─► readRow_ ─► setValues
│  saveActual           │  recordChange_ ─┬─ writeAmend_       (appendRow)
│  saveLetterParcelMix ─┘                 └─ logFieldChanges_ ─► logAudit_ (appendRow)
│
├─ saveMethodMixGrid ───► clearMixWindowBatch_    ─► recordChangesBatch_  (ranged setValues)
│  saveLetterParcelGrid ► clearLetterParcelWindow_ ─► recordChangesBatch_
│  cloneModellingIds   ─►                            recordChangesBatch_
│  copySegmentData     ─► appendCopies ×4          ─► recordChangesBatch_
│
├─ getResolvedByMonth ──┐
│  getMethodMixGrid     │  loadEngineInput_ ─► readDated_ ×5
│  getLetterParcelGrid  ├─ computeModel_ ─┬─ pointInTime_
│  previewOutput        │                 ├─ dayWeighted_
│  runValidation        │                 ├─ resolveSurcharge_
│  publishOutput ───────┘                 └─ checkMixInvariant_
│                          │
│                          ├─ runValidation(Quiet)_ ─► ruleValueRanges_ / ruleDateOrder_ /
│                          │     ruleRangeIssues_ / ruleCoverage_ / ruleDuplicates_ /
│                          │     ruleMethodCarrier_ / ruleTbcCarrier_ / ruleOutputSwing_
│                          └─ writeOutputTab_ / writeOutputDetailTab_ /
│                             writeValidationResults_ / writeCalcRun_
│
├─ getActualsComparison ─► loadActualsForClient_ ⋈ OUTPUT
├─ compareSnapshots ─────► valuesOf(snapshot │ live OUTPUT), scope-filtered
├─ takeForecastSnapshot ─┐
└─ weeklySnapshotJob ────┴► takeSnapshotInternal_ ─► withLock_ ─► Snapshot_Values + Snapshots
```

### 3.5 The engine is pure

`computeModel_()` takes plain objects and returns plain objects. It never touches
`SpreadsheetApp`, `Session`, or any Google service. `loadEngineInput_()` is the only
function that reads the sheet, and its whole job is turning rows into the shape the engine
wants. **This split is the single best decision in the codebase**: it is what makes
`runParityTest()` possible (replay the same inputs, diff against the legacy workbook) and
what lets `previewOutput()` run the entire model without writing anything.

The maths, restated from the engine's own header:

```
base(d,m)      = Σ base rates live on the 1st of month m           [point in time]
fuelPct(d,m)   = day-weighted average % across month m             [prorated]
otherAmt(d,m)  = day-weighted average amount across month m        [prorated]
rate(d,m)      = base × (1 + fuelPct) + otherAmt
methodMix(d,m) = mixCC × ccShare + mixAmbient × (1 − ccShare)
lpMix(d,m)     = letterMix if the route is LETTER, else 1 − letterMix
contribution   = rate × methodMix × lpMix
OUTPUT(h,m)    = Σ contributions across every Modelling ID in h
```

Proration is the subtlety: a base rate changing on the 15th does **not** apply until the
following month (point-in-time on the 1st), while a fuel surcharge changing on the 4th of a
31-day month contributes 28/31 to that month's blended figure (day-weighted). Which
behaviour applies is a property of the surcharge type, read from `Dim_Surcharge.Proration`.

Two guards exist specifically because the legacy workbook lacked them:

- **`checkMixInvariant_`** — for every segment / regime / class / month the method
  percentages must total 100%, else `MIX_SUM`.
- **`RATE_MISSING`** — a route carrying volume but priced at zero is reported rather than
  silently contributing nothing.

Preserve both when touching the engine.

---

## 4. Client

One HTML file, 2,313 lines, containing CSS, markup and all application JavaScript. No
framework, no build step, no external dependency — the logo is an inline base64 PNG and the
dashboard chart is hand-written SVG. For a single-file Apps Script deployment this is a
defensible choice, and it is the reason the app has no supply chain.

Structure:

- **`S`** — one global state object holding the user, all reference data, all loaded rows,
  the current screen, and a `sel` sub-object of current selections.
- **`TABS`** — an array of `[id, label, requiredCapability]` driving the nav; tabs whose
  capability the user lacks are not rendered.
- **`call(fn, args, onOk, onErr, quiet)`** — the sole wrapper around `google.script.run`.
  It maintains an `inflight` counter that drives the Saved / Saving… / Not saved indicator
  in the header, and a `beforeunload` handler warns if a write is in flight. Any new
  client-side write should go through it rather than raw `google.script.run`.
- **`openForm(cfg)`** — a hand-built modal form builder replacing `window.prompt`. The
  comment explains why: a browser prompt cannot label its fields, so everything had to be
  crammed into one comma-separated string, and a raw date serial leaking into a prefill was
  invisible. It supports text / number / date / select / month fields, hints, suffixes,
  required markers, and an "open-ended" checkbox that pins a date to `9999-12-31`.
- **`renderAll()`** — replaces the active screen's entire `innerHTML` from a `view*()`
  function, then runs a matching `wire*()` function to attach handlers.

Screens: `viewDashboard`, `viewRates`, `viewMixes` (3 sub-views), `viewOutput`,
`viewActuals`, `viewAdmin` (4 sub-views), `viewValidation`, `viewHistory` (2 sub-views).

---

## 5. The data model as implemented

### 5.1 One schema definition, derived everywhere

`TABLES` in `utils.js` is the single source of truth for every tab: sheet name, ordered
column headers, a display width, and an optional `amends: true`. Everything else is derived
from it at load time:

| Derived | Shape | Example |
|---|---|---|
| `SHEET` | table key → sheet name | `SHEET.RATE_BASE` → `'Rate_Base'` |
| `COL` | table key → header → 0-based index | `COL.RATE_BASE.Base_Rate` → `4` |
| `TABLE_BY_SHEET` | sheet name → table key | `'Rate_Base'` → `'RATE_BASE'` |
| `TABLES.*_AMENDS` | auto-generated per `amends: true` table | same headers, prefixed with `Amend_ID/TS/By/Type` |

**Column addressing is disciplined throughout.** Every read and write goes through
`COL.<TABLE>.<Header>`; I found **no hardcoded column index in application code**. The two
exceptions are both legitimate and both concern *foreign* data: `migrate.js` reads the
legacy workbook by position (`r[8]`, `r[10]`…) because that workbook's layout is fixed and
external, and `enginetest.js`/`verifyPublishedOutput` do the same against its `Output` tab.
`actuals.js` resolves the staging tab's columns **by normalised header name**, which is why
it survives the mojibake `Sum of Cost Of Shipping (Â£)` already present in the real sheet.

**Named ranges are not used anywhere.**

`setupCreateAllTabs()` reconciles a real sheet to `TABLES` by appending missing columns *by
name* — never reordering, never deleting. Add a header to `TABLES`, re-run setup, and the
sheet catches up. Amends tables are matched by column name too, so a new column on a source
table lands correctly in its history table with no extra code.

### 5.2 The tables

30 defined tables plus 8 auto-generated Amends tables. Row counts are from
`docs/SHEET_STRUCTURE.md`, snapshotted 2026-08-13.

| Group | Tables | Rows (data) | Written by |
|---|---|---|---|
| Dimensions | `Dim_Reference` 29, `Dim_Carrier` 7, `Dim_Method` 35, `Dim_Surcharge` 4, `Dim_Calendar` 36 | small, static | Admin screens, `setup.js` |
| Structure | `High_Level_IDs` 18, `Modelling_IDs` 254 | small | `structure.js` |
| Rates | `Rate_Base` 286, `Rate_Surcharge` 1,649 | grows | `rates.js` |
| Mixes | `Mix_Method` 374, `Mix_LetterParcel` 18, `Mix_ColdChain` 141 | grows | `mixes.js` |
| Outputs | `OUTPUT` 648, `OUTPUT_Detail` 9,144, `Actuals` 130, `Actuals_Import` 4,956, `Snapshots` 2, `Snapshot_Values` 1,260, `Calc_Runs` 6 | rewritten wholesale | `output.js`, `actuals.js`, `snapshots.js` |
| Governance | `Permissions` 2, `Portal_Roles` 4, `Scope_Mapping` **0**, `Audit_Log` 767, `Scenarios` 1, `Config` 13, `FX_Rates` 768, `Validation_Results` 2 | mixed | `auth.js`, `audit.js`, `validate.js` |
| History | 8 × `*_Amends` | grows unbounded | `audit.js` |

Effective-dated tables (`Rate_Base`, `Rate_Surcharge`, `Mix_*`) all carry
`Valid_From` / `Valid_To` / `Scenario_ID` / `Active`, with `9999-12-31` as the open-ended
sentinel. **Soft delete is the norm**: `Active = false` rather than row removal, so history
and audit keep making sense.

Note `Scope_Mapping` is empty. **Row-level scoping is not currently configured**, and with
`SCOPE_DEFAULT_ALLOW = TRUE` every user therefore sees every segment their role permits.

### 5.3 Reading and writing the sheet

**Bulk read.** `prewarmSheetCache_(names)` fetches many tabs in one HTTP call, trying three
paths in order and falling through cleanly: the Advanced Sheets Service, then the raw Sheets
REST API using the script's own OAuth token, then per-sheet `getDataRange().getValues()`.
The REST fallback exists because a Workspace admin can allowlist the Advanced Service away.
`diagnoseBatchGet()` reports which path is live and how much it saves.

**Per-execution cache.** `getAllData_(sheet)` memoises in `_sheetDataCache_` for the
lifetime of one execution. Any write must call `invalidateSheetCache_(sheet)` before the
same execution reads that tab again — and the write paths do, sometimes twice.

**Bulk write.** Every write of more than one row uses a single ranged `setValues`. The
9,000-row `OUTPUT_Detail` write is one call, with a comment noting that chunking measured
slower. `clearDataRows_` clears contents rather than deleting rows, because
`deleteRows(2, lastRow-1)` fails once the sheet contains only header plus data
("cannot delete all non-frozen rows") and clearing avoids the reflow.

### 5.4 Where the sheet assumptions could silently break

This is the section to read before changing anything.

**1. `batchGet` returns ragged rows.** Trailing empty cells are omitted, so `row[i]` past
the last non-empty cell is `undefined`, and a row array can be shorter than its header list.
Almost everything is protected because reads go through `safeStr` / `safeInt` / `safeNum` /
`safeBool`, which all handle `undefined`. The unprotected path is `copySegmentData`, which
does `data[i].slice()` on a cached row and later writes at full header width — see
[FINDINGS.md](FINDINGS.md) C6.

**2. Dates arrive as serial numbers, not `Date` objects.** `batchGet` is called with
`dateTimeRenderOption: 'SERIAL_NUMBER'`, so a date is a count of days since 1899-12-30.
`new Date(47391)` is 1970, not 2029. Since the entire model is date-range arithmetic, this
is the highest-consequence hazard in the system, and it is handled well: `normaliseDate` /
`dateKey` / `dayNum` are the only functions permitted to interpret a raw cell as a date, and
`normaliseDate` bounds-checks so that a genuine number (a rate of 3.78, a percentage of
0.14) is not misread as a date. `isDate_` is duck-typed on a working `getTime()` rather than
`instanceof Date`, because a `Date` crossing an execution boundary belongs to a different JS
realm. `dayNum` reduces every date to an integer via `Date.UTC`, so comparisons and day
counts cannot be shifted by a timezone or a DST boundary. **Route everything through these
and do not reinvent them.**

**3. Sheets coerces anything date-shaped on write.** `Month_Label` "Jan-26" becomes 26
January; `HORIZON_START` "2026-01-01" becomes the serial 46023. Mitigated three ways:
`formatForColumn_` sets `@` (plain text) on label and code columns at tab creation;
`configDate()` normalises on read so it cannot matter how the cell ended up; `monthLabel_()`
rebuilds a label from the month start if the stored value looks numeric. `repairConfigValues()`
and `repairCalendarLabels()` exist to fix sheets where it already happened. Anyone editing
Config by hand can reintroduce it.

**4. Row index = array index + 1.** `findRowById_`, `findRowByText_` and `recordLogin_` all
convert a cached-array position directly to a sheet row. True only because reads start at A1
and the header is row 1. Inserting a row above the header would silently redirect every
write.

**5. Soft delete means tables only grow.** `Rate_Surcharge` is already 1,649 rows and every
read is a full-table scan filtered in JavaScript. `Audit_Log` and the eight Amends tables
grow without bound; only `pruneSnapshots()` exists to thin anything, and it is unscheduled.

**6. Two spreadsheet IDs are hardcoded and committed** — `SPREADSHEET_ID` (`utils.js:14`)
and `SOURCE_SPREADSHEET_ID` (`migrate.js:20`).

**7. `Dim_Surcharge.Applies_To` and `Apply_Order` are loaded and never used.** The engine
always computes `base × (1 + Σpct) + Σamt`. The Admin form invites you to set an apply
order that does nothing.

**8. `Scenario_ID` is honoured by the engine but not by the client loaders**, and
`assertNoOverlap_` ignores it for base rates and surcharges. Invisible while one scenario
exists.

---

## 6. Personal data

Flagged per the house rule in `CLAUDE.md`.

The only personal data in the system is **staff identifiers**: email addresses and display
names in `Permissions` and `Scope_Mapping`, `Last_Login_TS`, the `Email` column of
`Audit_Log`, and `Created_By` / `Updated_By` / `Amend_By` on every table that carries them.
It is used for access control and for attributing changes — not for any automated decision
that affects an individual.

No customer data and no special-category data appears in the schema or in any code path I
read. `Actuals_Import` holds aggregate shipment counts and costs grouped by country, brand,
carrier, method, treatment type and month — no individual-level records. `Treatment_Type`
and `WL_Split` are health-adjacent at the *product* level (weight loss, Mounjaro, WeGovy)
but describe product lines, not people.

One thing to watch during the UI overhaul: the History screen surfaces `Audit_Log` rows
including the acting user's email, and does so to anyone whose role carries `viewAudit`.
That is appropriate for a change log, but it is employee personal data on screen, and it is
currently not filtered by scope.

---

## 7. What I inferred rather than read

Kept deliberately separate from everything above, because the difference matters.

### Observed — high confidence

The entry-point inventory (grepped, not assumed); the absence of `flush()`, of per-cell
`getValue`/`setValue` in hot paths, and of any `CacheService` use; the column-addressing
discipline; the two-phase load; the seven-step write shape; the purity of `computeModel_`;
and every file-and-line reference in this document and in FINDINGS.md, each re-verified
against the source before publishing.

### Inferred — please correct

1. **Who uses this and how many.** `Permissions` holds 2 people and `Scope_Mapping` is
   empty, so scoping is effectively off. Several findings are rated on the assumption that
   scoping is *intended* to be switched on. If it never will be, they drop a severity.
2. **Whether the legacy workbook still matters.** `runParityTest()` and
   `verifyPublishedOutput()` still read `SOURCE_SPREADSHEET_ID`. If cutover is complete that
   entire dependency can go; if it is still reconciled monthly, it must stay.
3. **Whether `Scenarios` is live or aspirational.** The table exists, the engine filters on
   it, and every client call site passes a hardcoded `1`.
4. **Whether `bulkRateChange` was wired up and removed, or never wired up.** It is the
   highest-value-looking feature in the codebase — "Royal Mail plus 6% from April", with a
   preview mode — and nothing in `index.html` calls it. Git history has only three commits,
   so I cannot tell which.
5. **Whether cold chain is genuinely out of scope or deferred.** The code says out of scope;
   the read-only screen directs users to edit the sheet by hand, which bypasses the audit
   trail.
6. **Whether anyone edits the Sheet directly.** If they do, the audit trail has holes by
   design and locking protects less than it appears to.
7. **Whether the 6-minute execution limit has actually been hit.** I reasoned about it from
   API call counts; I have not observed a timeout.
8. **Real production timings.** `perfReport()` is instrumented and returned in every payload,
   but I have no captured numbers. The performance severities in FINDINGS.md come from
   call-count reasoning, not measurement — one `testApiPayload()` run and one browser network
   trace would confirm or demolish them.

### Questions whose answers change the recommendations

- **Is the UI overhaul a restyle of the current eight screens, or a rethink of the
  journeys?** The findings about the client are priced very differently under each.
- **Is a build step acceptable** (bundler, TypeScript, `clasp` pushing from `src/`), or must
  this stay hand-editable in the Apps Script editor? This is the single largest constraint
  on any refactoring plan.
- **Is there appetite for automated tests?** The pure engine could run under plain Node
  today with no Apps Script runtime at all.
