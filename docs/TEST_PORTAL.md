# The test portal — a durable record of the non-production fixture

**Last confirmed: 2026-08-25.** This file exists because the same facts were previously spread
across `BASELINE.md`, `PHASE_0_VERIFICATION.md` and `FINDINGS.md`, and were assumed lost. They
were not lost — they were just not in one place. Keep them here.

---

## The two script projects

| | Script project ID | What it is |
|---|---|---|
| **Test** | `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` | **This is the `scriptId` committed in every worktree's `.clasp.json`.** `clasp push` lands here. |
| **Production** | `1JHYSultHlRsjZkWcJ9JJ8CNEsBFCUYJVe_n1tENX14_g48ohbe80rGWw` | Appears in **no** `.clasp.json`. Pushing to it requires deliberately editing one. |

**Read that table twice before assuming a push is dangerous.** The committed `scriptId` is the
*test* project. A plain `clasp push` from `pfp-bulk`, `pfp-actuals` or `postage-forecast-portal`
does not touch production code. Confirmed alive and reachable on 2026-08-25 via `clasp pull`.

Sources: `BASELINE.md` §"Step 0 — record both environments", `PHASE_0_VERIFICATION.md` step 3
("the ID above is the test project"), `IMPROVEMENT_PLAN.md` §Environments.

---

## The spreadsheets — this is where the danger actually is

| | Spreadsheet ID | What it is |
|---|---|---|
| **Test copy** | `1YTwIvSrLwaObPAdIT0G7-v_o1RG2J2nOcDvVKNJEoSU` | `Copy of Postage Forecast — Data`. **Live and in use** — the test project's `SPREADSHEET_ID` property points here. Confirmed by `showEnvironment()` 2026-08-25 13:43. |
| **Production** | `1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE` | `Postage Forecast — Data`. The live forecast. **Confirmed 2026-08-26 15:01** — see below. Also `SPREADSHEET_ID_FALLBACK` in `utils.js`. |
| **Legacy workbook** | `18a-Pfaa3hy0OlNcl3jYri-qccLku13JyDhfm0lOq4io` | Read-only source for `migrate*()` and `runParityTest()`. From `SOURCE_SPREADSHEET_ID_FALLBACK` in `migrate.js`; no property set. |

### Confirmed 2026-08-26 15:01 — `showEnvironment()` on the PRODUCTION project

```
  script project  : 1JHYSultHlRsjZkWcJ9JJ8CNEsBFCUYJVe_n1tENX14_g48ohbe80rGWw
--- app data: everything the portal reads and writes ---
  spreadsheet ID  : 1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE
  resolved from   : the committed fallback in utils.gs — no SPREADSHEET_ID property is set
  file name       : Postage Forecast — Data
```

Three things this settles, all of which had been recorded here as open:

1. **`1UgIBKz…` is production's live data**, and it **opens fine** — the name resolves. An earlier
   report that the file would not open was something else (wrong account, or the legacy-workbook
   link). `BASELINE.md`'s blank "App spreadsheet ID" row for production can be filled in from here.
2. **`1JHYSult…` is the production script project**, confirming the table above.
3. **Production loads.** `showEnvironment()` returning at all proves it, which matters because a
   load-order fault (M9) fires before any function body runs and would have failed identically.

Production runs on the **committed fallback with no property set**, which is correct for
production: `SPREADSHEET_ID_FALLBACK` names its own spreadsheet. The test copy
(`Copy of Postage Forecast — Data`) is a copy of this file.

### The trap, stated plainly (FINDINGS.md **S14**)

The committed `scriptId` is **test**, but `SPREADSHEET_ID_FALLBACK` is **production** — now
confirmed rather than inferred. So:

> **A test project with no `SPREADSHEET_ID` Script Property set reads and writes the PRODUCTION
> spreadsheet, and the only signal is running `showEnvironment()` and reading the output.**

Nothing in the code enforces the script-project-to-spreadsheet pairing. The property is the only
thing separating the environments, and an unset property fails **open** rather than refusing. The
safe default is the dangerous one.

Anything named `test…Write()` — `testRateWrite`, `testMixWrite`, `testStructureWrite`,
`testAuditTrail` — and any commit-mode maintenance function (`runActualsImport`,
`runFuelScheduleCopy`) writes real rows. **Never run one without checking `showEnvironment()`
first.**

### The fixture was never lost — only its record was

`PHASE_0_VERIFICATION.md` step 16 offered two end states, **recorded neither**, and in the same
breath said *"Delete the throwaway copy → Move to bin → then empty the bin."* Read together, that
looks like the copy is gone.

**It is not.** `showEnvironment()` on 2026-08-25 shows the property set and resolving to
`Copy of Postage Forecast — Data`. **Option A was taken.** A verification pass in August 2026
came within one click of re-copying production to rebuild a fixture that had been working the
whole time — which is the entire reason this file exists. If you change the end state, record it
*here*, in this table, on the same day.

---

## Re-establishing the test spreadsheet, if it is ever lost

**Not needed today.** Only for if the property is deleted or the copy is trashed. Only step 2
needs a browser — there is no CLI action that copies a Google Sheet.

1. **Check the current state first — it is free and it may answer everything.** Open the test
   project (`clasp open-script` from any worktree), then `utils.gs` → function dropdown →
   `showEnvironment` → **▷ Run**. Read `resolved from`. If it already says
   `the SPREADSHEET_ID Script Property` and names a copy, **stop — there is nothing to do.**

2. **Make the copy.** Open the production spreadsheet → **File > Make a copy**. Name it so it
   cannot be mistaken for production.

   > **Sharing.** The copy contains production's real rates, surcharges and actuals frozen at
   > the moment of copying. Copying it does not make the data less real. Set its sharing to
   > match production's — not looser. A copy is how a restricted file quietly becomes an
   > unrestricted one.

3. **Point the test project at it.** Put the copy's file ID in `TEST_SPREADSHEET_ID` in
   `utils.js`, `clasp push`, then run `pointThisProjectAtTestSpreadsheet` from the Run menu.
   That function refuses unless it is running on the test project, refuses if the ID is still
   the placeholder, refuses if the ID *is* the fallback's, and refuses if the sheet cannot be
   opened — then sets the property, clears every cache keyed on the old file, and runs
   `showEnvironment()` to prove the result.

4. **Verify, and do not skip this.** `showEnvironment()` must report:

   | Line | Must say |
   |---|---|
   | `script project` | `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` |
   | `spreadsheet ID` | the **copy's** ID — **not** `1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE` |
   | `resolved from` | `the SPREADSHEET_ID Script Property` |
   | `file name` | the copy's name, and **not** a line reading `CANNOT OPEN IT` |

   If `resolved from` says `the committed fallback in utils.gs — no SPREADSHEET_ID property is
   set`, the project is on **production**. Stop.

5. **Record it.** Fill in the table above with the copy's ID and the date, and commit.

`clearTestSpreadsheetOverride()` reverses step 3. Note what reversing means here: the project
returns to the fallback — i.e. off the test copy and onto whatever `SPREADSHEET_ID_FALLBACK`
names. There is very rarely a reason to want that.

---

## Known differences between test and production

Do not treat the test project as a performance replica.

| | Test `1Br6gJdV…` | Production `1JHYSult…` |
|---|---|---|
| `prewarmSheetCache_` path | **`fallback`** — one round-trip per tab | **`rest`** — real batched read |
| 8 tabs, batched | 219 ms *(not a read — it is the cost of both batch paths failing)* | **559 ms** |
| 8 tabs, per-sheet | **3,468 ms** | 2,767 ms |
| Effective read speed | ~**6 × slower than production** | baseline |

Cause: the test project's GCP project (`736650435578`) does not have the Sheets API enabled, so
the REST batch path returns HTTP 403 and `prewarmSheetCache_` falls through to path 3
(FINDINGS.md **P14**). Enabling the Sheets API in that GCP project is the cheapest fix and makes
test *match* production; adding the Advanced Sheets Service to `appsscript.json` instead moves
**both** environments onto a third path and drags production along with it.

**Consequences to keep in mind:**

- Read-cost measurements taken on test are inflated roughly sixfold. `BASELINE.md` says the same.
- **A long write batch is likelier to hit the 6-minute execution limit on test than on
  production.** `runFuelScheduleCopy()` makes ~65 lock-taking write calls; budget for that being
  slower here.
- The test project has a web app deployment at `@HEAD` (1 deployment, confirmed 2026-08-25), so
  `clasp push` changes what its portal serves immediately. `showPortalUrl()` reports a `/dev`
  URL as broken — that is **M10**, a false alarm, not a problem.

---

## Which code the test project carries

Confirmed by `clasp pull` on 2026-08-25: the test project matches the **`tmp-merge-test`** branch
(the `pfp-bulk` worktree) — `bulkUpdateRates`, `bulkRateShape_`, and the actuals diagnostics
(`diagnoseActualsTreatments`, `diagnoseActualsExtract`, `diagnoseActualsSegments`) are all present.

It does **not** carry the CSV-upload-to-actuals work (`uploadActualsCsv`, `parseCsv_`,
`requireImportActuals_`), which is on `feature-output-table` in the `postage-forecast-portal`
worktree. Reconciling the two branches is a separate task.

Because all three worktrees share one `scriptId`, **a push from any of them overwrites the
others' code on the test project.** If two branches need testing at once, that is the reason to
create a genuinely second test project — not code safety, which the table at the top already
provides.

---

## Confirmed state, 2026-08-25 13:43 — `showEnvironment()` on the test project

```
  script project  : 1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG
  spreadsheet ID  : 1YTwIvSrLwaObPAdIT0G7-v_o1RG2J2nOcDvVKNJEoSU
  resolved from   : the SPREADSHEET_ID Script Property
  file name       : Copy of Postage Forecast — Data
  legacy workbook : 18a-Pfaa3hy0OlNcl3jYri-qccLku13JyDhfm0lOq4io  (fallback in migrate.gs)
```

The test project is correctly isolated for **app data**. Note that the **legacy workbook** is
still on its committed fallback with no property set — read-only for `migrate*()` and
`runParityTest()`, so lower risk, but not isolated. Do not run `migrateAll()` here assuming it is.

---

## First real use of the fixture: FUEL schedule preview, 2026-08-25

`previewFuelScheduleCopy()` (`rates.gs`) was run against the test copy to check whether High Level
ID 1's Royal Mail FUEL schedule needed rolling onto Modelling IDs 9-21. **It did not. Nothing was
written, and `runFuelScheduleCopy()` has never been run.**

**High Level ID 1 (`MEDEXPRESS_GB_WL_MOUNJARO`) has 18 active `ROYALMAIL` routes**, one per
delivery method — not one "RM route". Their FUEL schedules fall into three groups:

| Schedule | 01-01→05-03 | 05-04→06-12 | 06-13→09-30 | 10-01→2029-12-01 | Modelling IDs |
|---|---|---|---|---|---|
| **A** — 3 periods | 0.11 | 0.14 *(runs to 09-30)* | — | 0.14 | **5, 6** — `FIRSTCLASSNOSIG`, `INTERNATIONAL` |
| **B** — 4 periods | 0.11 | **0.14** | **0.16** | 0.14 | **7, 8** — `RM24`, `RM48` |
| **C** — 4 periods | 0.11 | **0.16** | **0.11** | 0.14 | **9-22** (14 routes) |

Schedule **C** is what the copy was meant to produce on 9-21, and 9-22 already had it. The
`06-13→09-30 @ 0.11` rows sit at `Rate_Surcharge` row IDs **1650-1663** — fourteen consecutive
IDs, so one batch — and the `05-04→06-12 @ 0.16` rows above them are original rows shortened from
`05-04→09-30`, which is `closePrevious` behaviour. That batch had already done the job.

**Open, unexplained:** the highest row IDs in the tab are **1748, 1749**, on Modelling IDs 7 and 8
(`RM24`/`RM48`) — i.e. edited *after* the 1650-1663 batch, and landing on schedule **B**, not C.
Whoever edits HL1's Royal Mail FUEL next should know that 7 and 8 diverge from the other sixteen
and that nobody has yet said whether that is deliberate.

**Modelling IDs 5-8 are the outliers** if one schedule across all 18 routes is ever wanted. 5 and
6 have no mid-year change at all.

---

## Consolidated to production only — 2026-08-26

The multi-worktree, multi-branch layout this file was written to document is gone.

| Before | After |
|---|---|
| 4 working directories | **1** — `C:\tools\postage-forecast-portal` |
| 14 branches | **`main`**, carrying everything |
| `.clasp.json` → test `1Br6gJdV…` | **`.clasp.json` → production `1JHYSult…`** |

**`clasp push` from that directory is now a production deploy, live immediately**
because production serves `@HEAD`. Nothing prompts. That is the point of the change —
"which portal am I pushing to" cost real time — but it removes a structural
safeguard and replaces it with a procedural one.

### If the test project still exists

Its `scriptId` is `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` and its
`SPREADSHEET_ID` property points at `1YTwIvSr…` (`Copy of Postage Forecast — Data`).
No local `.clasp.json` points at it any more, so pushing to it means a scratch
directory with its own `.clasp.json` — the pattern used throughout August 2026.

### If the test project is deleted, delete the test SPREADSHEET too

Not optional, and the order matters. A test *project* left alive with its test
*spreadsheet* deleted falls back to `SPREADSHEET_ID_FALLBACK`, **which is
production** — so it would silently read and write live data with no signal but
`showEnvironment()`. That is FINDINGS.md **S14** firing for real. Either both go or
both stay.

Also: `compareRatesWithTestCopy()` and `previewRateCopyFromTestCopy()` need **both**
spreadsheets to exist. Once the copy is binned there is no way to check the two ever
matched. **Confirm parity before deleting, never after.**
