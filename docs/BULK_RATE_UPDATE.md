# Bulk rate update — feature notes and verification

The "All" bulk-update option on the Rates tab, for both Base Rates and Surcharges.
Branch `feature-bulk-rate-update`, based on `main`.

**This is deliberately not a phase.** Phases 0–8 of [IMPROVEMENT_PLAN.md](IMPROVEMENT_PLAN.md)
are behaviour-preserving by design and verified with `previewOutput()` reporting
`Nothing would change`. This feature exists to change what the forecast says, so that oracle
does not apply to it and the plan's phase numbering is untouched.

---

## What it does

Name a value for any of the seven rate dimensions, or leave it as **All**. Every *existing*
rate matching the selection is superseded by a new period, exactly as **Add rate change** does
for one route: the period running on each matching route is shortened to the day before, and a
new row is written from the date you gave.

Two worked examples, both from the original request:

| Selection | Effect |
|---|---|
| brand **All**, treatment **All**, country **All**, courier **All**, method **RM24**, class **LETTER**, surcharge **FUEL** | Every 24-hour letter FUEL surcharge, across every brand / treatment / country / courier, moves to the new value |
| brand **MedExpress**, everything else **All**, surcharge **FUEL** | Every MedExpress FUEL surcharge moves, whatever its other dimensions |

### The seven dimensions

Neither rate table stores a dimension. Both key on `Modelling_ID`, and every dimension is
reached by a two-hop join — `Rate_Base` / `Rate_Surcharge` → `Modelling_IDs` → `High_Level_IDs`:

| Selector | Column | Table |
|---|---|---|
| Brand | `Brand` | High_Level_IDs |
| Country | `Geo` | High_Level_IDs |
| Treatment type | `Treatment_Type` | High_Level_IDs |
| WL split | `WL_Split` | High_Level_IDs |
| Courier group | `Carrier_Code` | Modelling_IDs |
| Method | `Method_Code` | Modelling_IDs |
| Class | `Letter_Parcel` | Modelling_IDs |

Method and class are **two separate columns**, both on `Modelling_IDs`.

Surcharges add an eighth, `Surcharge_Code`, which is the only dimension stored on the rate row
itself — and the only one that can never be All. `FUEL` is a percentage while `GREEN`, `PEAK`
and `SHIPSTATION` are amounts (`Dim_Surcharge.Value_Type`), so one entered number cannot mean
both a fraction and a currency value; All there would put 0.06 on one row as 6% and on the
next as 6 pence.

### Four things it deliberately does not do

- **It never creates a rate for a route that has none.** A selection matching 20 routes where
  3 have no rate updates 17 and reports 3 skipped. A missing rate is a structural gap worth
  seeing on the Rates screen, not filling in silently.
- **It does not filter on the current value.** Selection is by dimension alone; the preview
  instead lists the distinct current values found and warns when it is about to flatten more
  than one of them.
- **It does not edit rows in place.** Every change is a new period, so the old number survives
  in the rate table itself, not only in the Amends snapshot.
- **"All" is never stored.** It exists only in the selector, as the empty string. Every row
  written is an ordinary row against one concrete `Modelling_ID`.

**`WL_Split` stores `'*'`, meaning "not applicable"**, on segments that are not weight-loss.
That is a real value the selector offers as "Not applicable (\*)" and matches literally. It is
not a wildcard, and treating it as one would silently widen every change. There is also no
magic `ALL` string, for the same reason in reverse: a brand or carrier code called `ALL` is one
data operation away, and a wildcard that a data entry can impersonate is a wildcard waiting to
widen somebody's rate change.

---

## How it satisfies the design constraints

**Identical behaviour when nothing is All.** The bulk path is additive — `saveBaseRate`,
`saveSurchargeRate` and their delete paths are untouched apart from one refactor described
below. Selecting a specific value in all seven dimensions simply produces a one-row batch.

**The same per-row rules, not a second implementation.** Two rules were extracted so that one
implementation serves both paths:

| Was | Now |
|---|---|
| `assertNoOverlap_(sheetName, …)` read the sheet and applied the rule | `assertNoOverlapIn_(data, …)` *is* the rule; `assertNoOverlap_` is a one-line wrapper that reads the sheet and calls it |
| `closePrecedingPeriod_` decided which rows to shorten *and* wrote them | `precedingPeriodCloses_(data, C, keyMatch, newFrom)` decides — no reads, no writes; `closePrecedingPeriod_` applies it row by row exactly as before |

The bulk path calls those same two functions, plus `assertCanEditModellingId_` per row and the
identical value validation. Nothing is skipped and nothing is duplicated.

**Batched writes, because the alternative times out.** A single save costs about seven Sheets
API calls. `Rate_Surcharge` holds ~1,650 rows and a wide selection can match hundreds, which at
seven calls each would breach the six-minute limit and leave a half-applied price change. So
the plan is built against one in-memory copy of the table, then flushed as one ranged write for
the closed periods, one append for the new rows, and one `recordChangesBatch_` for the history.
Measured in the Node harness on an 8-route change: **3 sheet calls versus 39** for the same
result. The harness asserts the resulting table and the resulting audit trail are byte-identical
to looping the single-save path.

**Preview, then explicit confirmation.** `bulkUpdateRates` previews by default and writes
nothing. The preview returns a `planKey` — a digest of the matched rows and their current
values — and applying must send it back. If anything in the matched set moved in between
(somebody else edited a rate, a route was deactivated) the recomputed key differs and the write
is refused, so a confirmation can only ever apply the set that was actually shown.

**One batch, visible as one action.** Every row of a batch shares a reference
`BULK-yyyymmdd-hhmmss`, written to:

- `Notes` on every new rate row, alongside the selection that produced it
- `Audit_Log.Detail` on every audit row, via a new optional third argument to
  `recordChangesBatch_`
- one summary `Audit_Log` row, `Entity = BULK_RATE_UPDATE`

History → Change log gains a **Batch** column showing that reference as a pill, so fifty rows
read as one thing somebody did rather than fifty coincidental edits. `Source_Ref` keeps
whatever the user typed and only falls back to the batch reference when they typed nothing.

---

## Files changed

| File | Change |
|---|---|
| `rates.js` | `assertNoOverlapIn_` and `precedingPeriodCloses_` extracted; `dimensionMatches_`, `resolveDimensionTargets_`, `describeDimensions_`, `bulkRateShape_`, `planBulkRateUpdate_`, `planKey_`, `bulkUpdateRates`, `applyBulkRateUpdate_`, `bulkRowCurrency_`, `bulkBatchRef_`, `padRow_`, `testBulkRateUpdate` added |
| `audit.js` | `recordChangesBatch_` takes an optional `detail`, stamped on every audit row it writes |
| `index.html` | "Bulk update…" on both Rates cards; the dimension form and its confirmation step; `cfg.body` support in `openForm`; Batch column in the change log |

`bulkRateChange` is **left alone**. It is the older carrier / High-Level-ID-scoped bulk uplift,
still unreachable from the UI (finding **U1**), and reconciling the two belongs with U1 rather
than here.

---

## Verification

Run in the **test** project. Run `showEnvironment()` first — if it says the spreadsheet came
from the fallback, you are pointed at production and must not run any of this (finding **S14**).

### 1. Editor → `showEnvironment()`

**Pass:** it names the test spreadsheet and says the ID came from a **property**.

### 2. Editor → `testBulkRateUpdate()`

The main automated check. It snapshots `Rate_Surcharge`, resolves dimensions, previews,
refuses a stale confirmation, applies, verifies the result, then restores every pre-existing row
byte for byte and deactivates the rows it created. Safe to re-run.

**Pass:** the log ends `BULK RATE UPDATE WORKING (n checks)` with no `FAIL` lines. Expect to
see, among others:

```
--- dimension resolution ---
  every dimension All        : 232 editable routes (232 matched, 0 out of scope)
  ok   All-everything matches routes
  ok   naming a dimension narrows the set
  ok   WL_Split '*' is matched as a value
  ok   an unknown value matches nothing
--- preview ---
  ok   preview matched at least one row
  ok   preview wrote nothing
  ok   the same selection gives the same planKey
  ok   a stale planKey is refused
  ...
  ok   every pre-existing row is back exactly as it was
BULK RATE UPDATE WORKING  (n checks)
```

If it logs `SKIPPING THE WRITE`, the seed selection matched more than 25 rows and it declined
to touch that many; the preview half still passed, and the write half is covered by steps 5–8
below.

### 3. Editor → `testRateWrite()`, `testMixWrite()`, `testStructureWrite()`, `testAuditTrail()`

The refactor touched the single-save path's two rule functions, so these must still pass
unchanged.

**Pass:** `RATE WRITES WORKING`, `MIX WRITES WORKING`, `STRUCTURE GUARDS WORKING`,
`AUDIT TRAIL WORKING`.

### 4. Editor → `previewOutput()`

**Pass:** `Nothing would change. OUTPUT already matches the inputs.`

This confirms `testBulkRateUpdate()` really did restore everything. If it reports differences,
the restore did not complete — read the log from step 2 before doing anything else.

### 5. App → the preview shows the right count and writes nothing

Rates → any segment → **Surcharges** card → **Bulk update…**

- Surcharge type: **Fuel surcharge**
- Brand, country, treatment, WL split, courier: leave all as **All**
- Method: **RM24** · Class: **Letter**
- Effective from: the 1st of a month a few months out · Until: tick **No end date**
- New value: something clearly recognisable, e.g. `0.1777`
- **Preview**

**Pass:** a confirmation appears titled "Apply to N rates?", listing the count, the new value,
the effective dates, and a table of the distinct current values with how many rows sit at each.
If more than one current value is listed, an amber warning says applying will flatten them.

Now **Cancel**. Open the Rates screen for one of the routes that would have been affected.

**Pass:** nothing has changed. The preview wrote nothing.

### 6. App → apply, and check the rows

Repeat step 5 and press **Apply to N rates**.

**Pass:** the header shows a message naming the batch, e.g.
`8 rate(s) updated as batch BULK-20260818-143012 — 8 period(s) closed, 8 created.`

Then, for two or three different segments that matched:

**Pass:** the Surcharges table shows the previous FUEL period now ending the day before your
start date, and a new period from your date at `0.1777`. For a segment that did *not* match
(different method, or Parcel rather than Letter), FUEL is untouched.

### 7. App → History → Change log

**Pass:** the rows for this change all carry the same reference in the new **Batch** column,
and both halves are present — the `UPDATE` rows shortening `Valid_To` on the old periods and
the `CREATE` rows for the new ones, each with your email.

Note that a bulk update of N rates writes roughly 2N + 1 audit rows, and the change log shows
the 50 most recent. After a large batch it will show little else. That is the trade for
"every row touched must appear in the change log"; see the open question below.

### 8. App → confirm a stale preview is refused

Open **Bulk update…**, set up the same selection as step 6, and press **Preview** — but do not
confirm yet. In a second browser tab, edit one of the matching rates by hand on the Rates
screen. Return to the first tab and press **Apply**.

**Pass:** it refuses with *"These rates have changed since the preview was taken, so nothing
was applied."* and nothing is written.

### 9. App → a selection matching nothing

**Bulk update…** → set Method to something no route uses → **Preview**.

**Pass:** a "Nothing matches" dialogue explaining how many routes matched the dimensions, how
many had no rate to supersede, and that a bulk update never creates a rate for a route that
has none. Its button returns you to the form with your selection intact.

### 10. App → the single-route path is unchanged

Rates → any route → **Add rate change** with a date inside an existing period.

**Pass:** behaves exactly as before — the previous period's "To" moves to the day before, both
rows appear, and History shows the close and the create with **no** batch reference.

### 11. Clean up, then Editor → `previewOutput()`

Delete the test rates you added in step 6 (Rates → Delete on each), then run `previewOutput()`.

**Pass:** `Nothing would change.`

---

## Revert

`git revert` the commit. The refactor of `assertNoOverlap_` and `closePrecedingPeriod_` reverts
with it and the single-save path returns to its previous shape. No schema change and no data
migration, so nothing needs undoing in the spreadsheet — though any batch already applied stays
applied, and would need reversing through the UI like any other rate change.

---

## Open question, not guessed at

**The change log window.** A bulk update of 34 rates writes ~69 audit rows and the change log
reads the 50 most recent, so one batch fills it. The options are to raise the limit, to collapse
a batch into one expandable row, or to leave it — and which is right depends on how often you
expect to use this. Left as it is for now, with the Batch column making the situation legible
rather than confusing.

## Not covered by this feature

- **Bulk delete.** Not built. Removing a surcharge across many routes still has to be done
  route by route.
- **A percentage or absolute uplift** ("+6% everywhere"). This sets one value on every matching
  row. `bulkRateChange` already implements PCT / AMOUNT / SET uplifts and is unreachable from
  the UI — that is finding **U1**, and wiring it up is where an uplift belongs.
- **Scenarios.** Every call passes scenario 1, as the rest of the client does. `Scenario_ID` is
  deliberately not part of the match, because the single-save path does not filter on it either
  (finding **C5**); with one scenario the two are identical, and C5 has to be fixed across all
  the rate paths at once rather than diverging here first.
