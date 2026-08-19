# Bulk rate update — feature notes and verification

The bulk-update feature on the Rates tab: pick a charge type, then name a value for any
dimension or leave it as **All**. Branch `feature-bulk-rate-update`, based on `main`.

**This is deliberately not a phase.** Phases 0–8 of [IMPROVEMENT_PLAN.md](IMPROVEMENT_PLAN.md)
are behaviour-preserving by design and verified with `previewOutput()` reporting
`Nothing would change`. This feature exists to change what the forecast says, so that oracle
does not apply to it and the plan's phase numbering is untouched.

---

## The flow

**Bulk update…** sits next to the Modelling ID dropdown, with the pickers rather than on either
rate card — it changes many routes at once, so belonging to one card would misdescribe it.

**Step 1 — which charge.** A popup, before anything else, listing Base rate plus every active
surcharge type read live from `Dim_Surcharge`, plus **All charge types**. No default and no
primary button: the only ways out are to click a charge or to cancel. This comes first because
it decides what the value on the next page even means.

**Step 2 — which routes.** Six dimensions, each defaulting to All, plus the dates, the value and
a source reference.

**Step 3 — the preview.** Shows what would change and what it would become. Nothing is written
until this is confirmed.

### Why the charge type has to come first

| Charge | `Value_Type` | A value of `0.14` means |
|---|---|---|
| Base rate | — | £0.14 per parcel |
| Fuel surcharge | `PCT` | 14% |
| Green levy, Peak surcharge, ShipStation fee | `AMT` | £0.14 |

One number cannot mean all of those, so the two modes are:

- **one charge type → set-to.** Every matching row takes the value entered.
- **All charge types → percentage uplift.** Every matching row moves by the percentage
  entered, off its own current value. A 6% uplift means the same thing to a base rate, a
  fraction and an amount, which is the only way "All" can be coherent.

### The charge types, as they stand

`Dim_Surcharge` holds four (5 rows × 8 cols in the 2026-08-13 snapshot; nothing in `migrate.js`
touches that tab, only `setup.js` seeds it):

| Code | `Surcharge_Name` | `Value_Type` | Shown as |
|---|---|---|---|
| `FUEL` | Fuel surcharge | PCT | Fuel surcharge — percentage of base |
| `GREEN` | Green levy | AMT | Green levy — fixed amount |
| `PEAK` | Peak surcharge | AMT | Peak surcharge — fixed amount |
| `SHIPSTATION` | ShipStation fee | AMT | ShipStation fee — fixed amount |

Plus **Base rate — rate per parcel**, which is not in `Dim_Surcharge` at all; it is a different
table, so `listChargeTypes()` prepends it. Nothing is hardcoded in the client: the popup calls
`listChargeTypes()`, which reads `Dim_Surcharge` (Active only, in `Apply_Order`) and takes each
label from `Surcharge_Name`. A surcharge type added next year appears on its own.

### The six dimensions

Neither rate table stores a dimension. Both key on `Modelling_ID`, and every dimension is
reached by a two-hop join — `Rate_Base` / `Rate_Surcharge` → `Modelling_IDs` → `High_Level_IDs`:

| Selector | Column | Table |
|---|---|---|
| Brand | `Brand` | High_Level_IDs |
| Country | `Geo` | High_Level_IDs |
| Treatment type | `Treatment_Type` | High_Level_IDs |
| WL split | `WL_Split` | High_Level_IDs |
| Method | `Method_Code` | Modelling_IDs |
| Class | `Letter_Parcel` | Modelling_IDs |

Method and class are two separate columns, both on `Modelling_IDs`.

### Courier group, and why it was removed

`Modelling_IDs.Carrier_Code` is **not** selectable. A bulk change always spans every courier,
and nothing in `resolveDimensionTargets_` reads the field — it is ignored rather than defaulted,
so a hand-built payload cannot narrow by carrier either. `describeDimensions_` states
`carrier=All` explicitly in the audit trail so the scope is never ambiguous.

Worth recording that it is a **genuinely separate field**, not one implied by the others, even
though the Modelling ID label `DPD / DPDNEXTDAY / PARCEL` makes it look composite:

- `migrateModellingIds_` (`migrate.js:468`) populated carrier from source column 6 and method
  from column 7 — neither derived from the other. `Modelling_Code` is built *from* the three,
  not the reverse.
- `Dim_Method` gives each method one owning carrier, so naming a method usually implies one —
  but `METHOD_CARRIER` (`validate.js:216`) only **WARNs** when a route disagrees. Divergence is
  permitted by design.
- `TBC` is a real carrier code with its own rule, `ruleTbcCarrier_` (`validate.js:225`). A route
  can carry a real method and a placeholder carrier, which `Method_Code` cannot express.

**The capability this costs:** "courier = Royal Mail, method = All" used to update every Royal
Mail route in one action. That is no longer expressible — methods must be picked one at a time,
and `Dim_Method` holds 35. Removed by explicit request; noted here because it is the most
All-shaped selection there was.

### Four things it deliberately does not do

- **It never creates a rate for a route that has none.** A selection matching 20 routes where 3
  have no rate updates 17 and reports 3 skipped. A missing rate is a structural gap worth seeing
  on the Rates screen, not filling in silently.
- **It does not filter on the current value.** Selection is by charge type and dimension; the
  preview instead lists the distinct current values found and warns when set-to mode is about to
  flatten more than one of them.
- **It does not edit rows in place.** Every change is a new period, so the old number survives in
  the rate table itself, not only in the Amends snapshot.
- **"All" is never stored.** In the dimensions it is the empty string in the selector only; every
  row written is an ordinary row against one concrete `Modelling_ID`.

**`WL_Split` stores `'*'`, meaning "not applicable"**, on segments that are not weight-loss. That
is a real value the selector offers as "Not applicable (\*)" and matches literally — not a
wildcard, and treating it as one would silently widen every change. There is also no magic `ALL`
string in the dimensions, for the same reason in reverse: a brand or method code called `ALL` is
one data operation away, and a wildcard a data entry can impersonate is a wildcard waiting to
widen somebody's rate change.

---

## How it satisfies the design constraints

**Identical behaviour for a single route.** The bulk path is additive — `saveBaseRate`,
`saveSurchargeRate` and their delete paths are untouched apart from one refactor described below.

**The same per-row rules, not a second implementation.** Three pieces are shared:

| Was | Now |
|---|---|
| `assertNoOverlap_(sheetName, …)` read the sheet and applied the rule | `assertNoOverlapIn_(data, …)` *is* the rule; `assertNoOverlap_` is a one-line wrapper that reads the sheet and calls it |
| `closePrecedingPeriod_` decided which rows to shorten *and* wrote them | `precedingPeriodCloses_(data, C, keyMatch, newFrom)` decides — no reads, no writes; `closePrecedingPeriod_` applies it row by row exactly as before |
| The value checks lived inline in each save | `validateChargeValue_(shape, value)` holds them, and both modes are held to it — an uplift can drive a rate negative as easily as a typed value can |

The bulk path calls all three, plus `assertCanEditModellingId_` per row. Nothing is skipped and
nothing is duplicated.

**Batched writes, because the alternative times out.** A single save costs about seven Sheets API
calls. `Rate_Surcharge` holds ~1,650 rows and a wide selection can match hundreds, which at seven
calls each would breach the six-minute limit and leave a half-applied price change. So the plan is
built against **one working copy per table**, then each table is flushed as one ranged write for
the closed periods, one append for the new rows, and one `recordChangesBatch_` for the history.
Measured in the Node harness on an 8-route change: **3 sheet calls versus 39** for the same
result, with the resulting table and audit trail byte-identical to looping the single-save path.

One copy *per table*, not per charge type, matters: "All" puts several surcharge groups on
`Rate_Surcharge`, and separate copies would each flush over the last one's work. There is a
regression check for exactly that.

**Preview, then explicit confirmation.** `bulkUpdateRates` previews by default and writes nothing.
The preview returns a `planKey` — a digest of the charge type, the matched rows, their current
values and their intended new ones — and applying must send it back. If anything moved in between
(somebody else edited a rate, a route was deactivated) the recomputed key differs and the write is
refused. Keying on the charge type too means a plan cannot be replayed against a different charge.

**One batch, visible as one action.** Every row of a batch shares a reference
`BULK-yyyymmdd-hhmmss`, written to:

- `Notes` on every new rate row, with the move it made (`set to 0.165`, or `0.14 +6% -> 0.1484`)
  and the selection that produced it
- `Audit_Log.Detail` on every audit row, via the optional third argument to
  `recordChangesBatch_`
- one summary `Audit_Log` row, `Entity = BULK_RATE_UPDATE`

History → Change log has a **Batch** column showing that reference as a pill, so fifty rows read
as one thing somebody did. `Source_Ref` keeps whatever the user typed and only falls back to the
batch reference when they typed nothing.

---

## Files changed

| File | Change |
|---|---|
| `rates.js` | `assertNoOverlapIn_` and `precedingPeriodCloses_` extracted from the single-save path; `dimensionMatches_`, `resolveDimensionTargets_`, `describeDimensions_`, `activeSurchargeCodes_`, `listChargeTypes`, `bulkRateShape_`, `chargeTypeGroups_`, `validateChargeValue_`, `roundRate_`, `planBulkRateUpdate_`, `distinctMoves_`, `planKey_`, `bulkUpdateRates`, `applyBulkRateUpdate_`, `describeMove_`, `describeChange_`, `bulkRowCurrency_`, `bulkBatchRef_`, `padRow_`, `testBulkRateUpdate` added |
| `audit.js` | `recordChangesBatch_` takes an optional `detail`, stamped on every audit row it writes |
| `index.html` | **Bulk update…** moved to the picker row; charge-type popup, dimension form and confirmation; `cfg.body`, `cfg.afterRender` and `saveLabel:null` support in `openForm`; Batch column in the change log |

`bulkRateChange` is **left alone**. It is the older carrier / High-Level-ID-scoped bulk uplift,
still unreachable from the UI (finding **U1**), and reconciling the two belongs with U1.

---

## Verification

Run in the **test** project. Run `showEnvironment()` first — if it says the spreadsheet came from
the fallback, you are pointed at production and must not run any of this (finding **S14**).

### 1. Editor → `showEnvironment()`

**Pass:** it names the test spreadsheet and says the ID came from a **property**.

### 2. Editor → `testBulkRateUpdate()`

The main automated check. It lists the charge types, checks dimension resolution (including that
a carrier in the payload is ignored), previews in both modes, refuses a stale confirmation, then
applies a +6% uplift across **all** charge types — the only path that writes to both rate tables
in one action — verifies the result, and restores both tables byte for byte, deactivating the rows
it created. Safe to re-run.

**Pass:** the log ends `BULK RATE UPDATE WORKING (n checks)` with no `FAIL` lines. Expect:

```
--- charge types, read from Dim_Surcharge ---
  BASE  Base rate  (rate per parcel)
  FUEL  Fuel surcharge  (percentage of base)
  GREEN  Green levy  (fixed amount)
  PEAK  Peak surcharge  (fixed amount)
  SHIPSTATION  ShipStation fee  (fixed amount)
  ok   base rate is offered first
  ...
--- dimension resolution ---
  every dimension All        : 232 editable routes
  carrier=ROYALMAIL (ignored): 232  (must equal the All-everything count)
  ok   a carrier in the payload is ignored, not honoured
  ok   WL_Split '*' is matched as a value
  ...
  ok   ALL gives uplift mode
  ok   every uplift is the entered percentage of its own current value
  ok   both rate tables were written
  ok   every pre-existing row is back exactly as it was
BULK RATE UPDATE WORKING  (n checks)
```

If it logs `SKIPPING THE WRITE`, the ALL selection matched more than 40 rows and it declined to
touch that many; every preview check still passed, and the write half is covered by steps 6–8.

**Check the charge type list it prints matches the table above.** If a fifth type appears, it was
added to `Dim_Surcharge` since the snapshot — that is the feature working, not a fault.

### 3. Editor → `testRateWrite()`, `testMixWrite()`, `testStructureWrite()`, `testAuditTrail()`

The refactor touched the single-save path's rule functions, so these must still pass unchanged.

**Pass:** `RATE WRITES WORKING`, `MIX WRITES WORKING`, `STRUCTURE GUARDS WORKING`,
`AUDIT TRAIL WORKING`.

### 4. Editor → `previewOutput()`

**Pass:** `Nothing would change. OUTPUT already matches the inputs.` This also confirms
`testBulkRateUpdate()` really did restore everything. If it reports differences, the restore did
not complete — read step 2's log before doing anything else.

### 5. App → the button, and the popup that cannot be skipped

Rates → any segment.

**Pass:** **Bulk update…** sits on the same row as the High Level ID and Modelling ID dropdowns,
aligned with them, not in either card header.

Click it.

**Pass:** a popup titled "Which charge are you changing?" listing Base rate and every surcharge
type with how each is expressed, plus **All charge types**. There is **no Save/Continue button** —
only Cancel. Pressing Escape or clicking outside closes the whole thing without going on.

### 6. App → set-to mode on one surcharge

**Bulk update…** → **Fuel surcharge**. Then:

- Brand, country, treatment, WL split: leave as **All** · Method: **RM24** · Class: **Letter**
- Effective from: the 1st of a month a few months out · Until: tick **No end date**
- New value: `0.1777` · **Preview**

**Pass:** the confirmation is titled "Apply to N charges?" and shows a table of Charge / Now /
Becomes / Rates. Every row's "Becomes" is `0.1777`. If more than one current value is listed, an
amber warning says applying will flatten them. There is no courier row anywhere in the form.

**Cancel**, then open the Rates screen for a route that would have matched.

**Pass:** nothing has changed — the preview wrote nothing.

Repeat and **Apply**. **Pass:** a message naming the batch, e.g.
`8 charge(s) updated as batch BULK-20260819-143012 — 8 period(s) closed, 8 created.` For two or
three matching segments, the previous FUEL period now ends the day before your start date and a
new period runs from it at `0.1777`. A route with a different method, or Parcel rather than
Letter, is untouched.

### 7. App → uplift mode across all charge types

**Bulk update…** → **All charge types** → leave every dimension as All → dates a few months out →
**Change every charge by: 6 %** → **Preview**.

**Pass:** the confirmation lists several charge types, each with its own Now and Becomes — the base
rate moving from its rate to ~6% more, fuel from its fraction to ~6% more, and so on. It says how
many charge types it spans and that it crosses both rate tables. It does **not** warn about
flattening, because nothing is being flattened.

**Apply. Pass:** the message names one batch. On the Rates screen for a matching route, both the
base rate and the surcharges have a new period from your date, each about 6% above its old value.

### 8. App → History → Change log

**Pass:** every row from step 7 carries the same reference in the **Batch** column, spanning both
`Rate_Base` and `Rate_Surcharge` entries, with both halves present — the `UPDATE` rows shortening
`Valid_To` and the `CREATE` rows for the new periods, each with your email.

A bulk update of N charges writes roughly 2N + 1 audit rows and the change log shows the 50 most
recent, so after a large batch it will show little else. That is the trade for "every row touched
must appear in the change log"; see the open question below.

### 9. App → a stale preview is refused

Set up a preview as in step 6 but do not confirm. In a second tab, edit one of the matching rates
by hand. Return and **Apply**.

**Pass:** *"These rates have changed since the preview was taken, so nothing was applied."* and
nothing is written.

### 10. App → a selection matching nothing

**Bulk update…** → any charge → Method set to something no route uses → **Preview**.

**Pass:** a "Nothing matches" dialogue giving the counts and explaining that a bulk update never
creates a rate for a route that has none. Its button returns you to the dimension form with your
charge type and selection intact.

### 11. App → the single-route path is unchanged

Rates → any route → **Add rate change** with a date inside an existing period.

**Pass:** exactly as before — the previous period's "To" moves to the day before, both rows
appear, and History shows the close and the create with **no** batch reference.

### 12. Clean up, then Editor → `previewOutput()`

Delete the test rates added in steps 6 and 7, then run `previewOutput()`.

**Pass:** `Nothing would change.`

---

## Revert

`git revert` the commit. The refactor of `assertNoOverlap_` and `closePrecedingPeriod_` reverts
with it and the single-save path returns to its previous shape. No schema change and no data
migration, so nothing needs undoing in the spreadsheet — though any batch already applied stays
applied, and would need reversing through the UI like any other rate change.

---

## Open question, not guessed at

**The change log window.** A bulk update of 34 charges writes ~69 audit rows and the change log
reads the 50 most recent, so one batch fills it. The options are to raise the limit, to collapse a
batch into one expandable row, or to leave it — and which is right depends on how often this gets
used. Left as it is, with the Batch column making the situation legible rather than confusing.

## Not covered

- **Bulk delete.** Not built. Removing a surcharge across many routes is still route by route.
- **Selecting by courier group.** Removed by request; see above for the capability that costs.
- **A set-to value across mixed charge types.** Impossible by construction — All is an uplift.
  Setting one figure on several charge types means running the bulk update once per charge type.
- **Scenarios.** Every call passes scenario 1, as the rest of the client does. `Scenario_ID` is
  deliberately not part of the match, because the single-save path does not filter on it either
  (finding **C5**); with one scenario the two are identical, and C5 has to be fixed across all the
  rate paths at once rather than diverging here first.
