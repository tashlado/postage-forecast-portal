# Phase 0 — verification checklist

Step-by-step verification for Phase 0 of [IMPROVEMENT_PLAN.md](IMPROVEMENT_PLAN.md), which
made the target spreadsheet switchable via Script Properties. Written to be followed by
someone who does not work in the code: every step says exactly what to type or click, and
exactly what result confirms it.

**If any step fails, stop there.** Later steps assume the earlier ones passed, and a failure
part-way through is worth understanding before pressing on.

**You will need:** a PowerShell window in `C:\tools\postage-forecast-portal`, and a browser
signed in as the account that administers the portal.

**Time:** about 25 minutes for steps 1–11 and 13–16. Step 12 (the timings) is another 20 and
can be done separately.

---

## Read this first — the one known risk

`appsscript.json` declares its OAuth scopes explicitly rather than letting Apps Script work
them out, and the declared list is:

```
spreadsheets · script.external_request · userinfo.email · script.scriptapp
```

Phase 0 introduces a `PropertiesService` call into `doGet` — the portal's own page load —
where previously there was none. `PropertiesService` is believed to need no OAuth scope, and
the pre-existing `diagnoseMetabase()` already reads a script property under this same
manifest, which is good evidence. But it is not proven, and if it is wrong the symptom is not
a broken diagnostic — it is a broken portal.

**Step 4 is the test for this.** If step 4 fails with anything mentioning permissions or
scopes, stop: the fix is one line in `appsscript.json` and it should be added deliberately.

---

## Which project each step runs in

Most of this checklist runs in the **test** project. Three steps run in **production**, because
they are *measurements* and the two environments turn out not to be comparable — production
batches its reads, test does not, and test is roughly **6 × slower on reads** as a result
(FINDINGS.md **P14**). Measuring in test would produce numbers no real user ever experiences.

| Step | Project | What it is |
|---|---|---|
| 1–8 | **test** | verification — did Phase 0 break anything |
| 9 `testApiPayload()` | **production**, then once in test | measurement, both columns |
| 10 `previewOutput()` | **production** | read-only check of the forecast maths |
| 11 portal loads | **test** | verification — did Phase 0 break `doGet` |
| 12 browser timings | **production** | measurement |
| 12 save call count | **test** | measurement, but valid from test |
| 13–16 | **test** | the copy, the property, `setupVerify()`, teardown |

**The one rule that matters:** never run a function with **`Write`** in its name against
production — `testRateWrite`, `testMixWrite`, `testStructureWrite`, `testAuditTrail` all append
real rows. None of them appear in this checklist; they belong to Phase 1 onwards, in test.

Everything this checklist does in production is read-only, with a single exception that is not
really an exception: `testApiPayload()` calls `initApp()`, which calls `recordLogin_`
(`auth.js:310`) to stamp one `Last_Login_TS` cell, throttled to once per 15 minutes. That is
precisely what happens when anyone opens the portal.

⚠️ **Order matters in one place:** capture step 9's and step 12's **test** figures *before* step 14
sets the `SPREADSHEET_ID` property. After that, test reads the throwaway copy and its row counts
stop matching production's.

---

## Part A — Get the code onto Google's servers

### Step 1. See what is about to be pushed

- [ ] In PowerShell, type this and press Enter:

```
git status
```

**It worked if** the output says `On branch phase-0-environment` and lists exactly these
files — seven modified, plus new files in `docs/`:

```
modified:   CLAUDE.md
modified:   docs/FINDINGS.md
modified:   docs/IMPROVEMENT_PLAN.md
modified:   enginetest.js
modified:   migrate.js
modified:   setup.js
modified:   utils.js
```

**If it lists source files beyond those seven** — something changed that was not part of
Phase 0. Stop and find out what.

---

### Step 2. Push the code to the Apps Script project

- [ ] Type:

```
clasp push
```

**It worked if** you see a list of filenames followed by `Pushed 17 files.` The list must
contain `utils.js`, `migrate.js`, `setup.js`, `enginetest.js`, `index.html` and
`appsscript.json`.

- [ ] **Check that `utils.js` is the FIRST file in that list.** `.clasp.json` pins it there,
      and the push order becomes the order in which Apps Script evaluates the files. If
      `utils.js` is not first, `"filePushOrder": ["utils.js"]` has gone missing from
      `.clasp.json` — stop and restore it before running anything else. FINDINGS.md M9 explains
      why this matters.

- If it asks `Manifest file has been updated. Do you want to push and overwrite?` — type `y`
  and Enter. Phase 0 did not change the manifest, so this should not appear, but it is
  harmless.
- If it says you are not logged in — run `clasp login`, complete the browser sign-in, then
  run `clasp push` again.
- **If the file count is below 17, or `utils.js` is not in the list** — stop. Nothing below
  would be testing the new code.

---

### Step 3. Open the Apps Script editor

- [ ] Type:

```
clasp open-script
```

(The command is `open-script`, not `open` — clasp 3.x split `open` into one command per
target. `clasp open` gives `unknown command`.)

Depending on your setup this either opens a browser tab or prints a
`https://script.google.com/d/…/edit` address for you to click.

**It worked if** the address bar — or the printed address — contains this exact ID:

```
1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG
```

and the left-hand file list shows `Code.gs`, `utils.gs`, `migrate.gs`, `setup.gs` and the
rest, plus `index.html`.

**If the ID is different, you are in the wrong project** — the ID above is the test project.
Stop.

Keep this tab open. Every step up to 15 uses it.

---

## Part B — Confirm the environment reporting works

### Step 4. Run `showEnvironment()` — the most important step

This is the function Phase 0 added. It reports which spreadsheet the code is pointed at.

- [ ] In the left-hand file list, click **`utils.gs`**. Its code appears in the middle.
- [ ] Look at the toolbar directly above the code: **▷ Run**, a bug icon **Debug**, then a
      **dropdown showing a function name**, then **Execution log**.
- [ ] Click that dropdown, scroll the list, and click **`showEnvironment`**.
- [ ] Click **▷ Run**.
- [ ] If a dialog says **"Authorization required"**: click **Review permissions** → choose
      the portal administrator account → if a warning screen appears, click **Advanced** then
      **Go to Postage Forecast Portal (unsafe)** → click **Allow**. This is normal for your
      own script.
- [ ] A panel opens at the bottom with the log. Wait for `Execution completed`.

**It worked if** the log contains all four of these:

| Line | Must say |
|---|---|
| `script project` | `1Br6gJdVW1-nOeOamrVhpMYCp5uCO79LQK5_m35eTbI9vkIGsDgZbeOiG` |
| `spreadsheet ID` | `1UgIBKzJAEJM-U2MzmcEIBuf21scLDCcu9ZOZKtI2KLE` |
| `resolved from` | `the committed fallback in utils.gs — no SPREADSHEET_ID property is set` |
| `file name` | a real spreadsheet name, **not** a line saying `CANNOT OPEN IT` |

That is Phase 0's safety guarantee proved: with no property set, the code is on exactly the
spreadsheet it was on before the change.

- [ ] **Write down the `file name` and the `url`.** Step 13 needs them.

**If instead you see:**

| Message | What it means |
|---|---|
| `ReferenceError: ENGINE_VERSION is not defined` at `setup.gs:220` | **Happened once, on the first run of this checklist, and is fixed** — see FINDINGS.md M9. It means the project's file order puts `setup.gs` before `utils.gs`. If it reappears, something reintroduced a top-level cross-file reference, or `filePushOrder` was removed from `.clasp.json`. It is not specific to `showEnvironment` — every function fails identically, because the error fires at file load. Stop; re-pushing reverted code will **not** fix it. |
| `That is an administrator function.` | Your email is not an administrator in that spreadsheet's `Permissions` tab. Not caused by Phase 0. |
| Anything about permissions or scopes that persists after the dialog in the step above | The manifest risk described at the top of this file. Stop. |
| `Could not read Script Properties, so the target spreadsheet cannot be confirmed` | Phase 0's new safety guard fired deliberately rather than guess which spreadsheet to use. Stop. |
| A spreadsheet ID different from the one in the table | Stop immediately. This is the outcome the whole phase exists to prevent. |

> ⚠️ **Note what step 4 just told you.** This test *script project* is pointed at the
> **production** spreadsheet. That is expected before step 13, and it is the reason step 13
> exists. Until the property is set, **do not run any function whose name contains `Write`**
> — `testRateWrite`, `testMixWrite`, `testStructureWrite`, `testAuditTrail` all write real
> rows. This is recorded as finding **S14** in [FINDINGS.md](FINDINGS.md).

---

### Step 5. Run `diagnoseBatchGet()`

Checks that the two bulk-reading paths edited in Phase 0 still work.

- [ ] Stay on **`utils.gs`**. Function dropdown → **`diagnoseBatchGet`** → **▷ Run**. Takes
      5–20 seconds.

**It worked if** the log contains a `Path used:` line naming `advanced`, `rest` or `fallback`,
and a `Per-sheet reads:` line — and the run itself completed rather than throwing.

**All three paths are a pass.** The function's job is to report which one is live, not to
insist on a particular one:

| `Path used:` | What it means | Extra lines to expect |
|---|---|---|
| `advanced` | Advanced Sheets Service is on — the fast path | `Speed-up: N.Nx` |
| `rest` | Advanced Service absent, REST works — also fast | `Speed-up: N.Nx` |
| `fallback` | Neither batch path available; reads happen per tab | **No `Speed-up` line.** Instead a block beginning `Neither batch path is available.` |

**Messages that look like errors but are normal.** On the `fallback` path the log will contain
one or both of these. They are the code reporting a handled fall-through, not a failure:

```
Advanced Sheets Service failed, trying REST: …
REST batchGet returned HTTP 403: Google Sheets API has not been used in project … or it is disabled
Using per-sheet reads for 8 tabs. Correct but slower …
```

**One figure is easy to misread.** On the `fallback` path, the milliseconds in
`Path used: fallback (Nms for 8 tabs)` is **not** a read time — path 3 reads nothing, it just
hands off to `getAllData_`. That number is the cost of the two failed batch attempts. The only
real read figure is the `Per-sheet reads:` line.

- [ ] **Record in [BASELINE.md](BASELINE.md):** the path used, whether the Advanced Service is
      present, the batch-attempt milliseconds, the per-sheet milliseconds, and the speed-up if
      one was printed.

**Only treat this as a failure if** the execution stops with a red error and no `Path used:`
line at all — that would mean the code Phase 0 changed is broken, rather than the environment
lacking a fast path.

---

### Step 6. Run `listSourceTabs()`

Checks the seven edits made in `migrate.gs` for the legacy workbook. Read-only.

- [ ] Click **`migrate.gs`** in the file list. Dropdown → **`listSourceTabs`** → **▷ Run**.

**It worked if** the log starts `=== SOURCE WORKBOOK TABS ===`, names a workbook, then lists
tab names including `High Level IDs`, `Modelling IDs` and `Base Rate Card`.

**If it says `CANNOT OPEN SOURCE WORKBOOK`** — this may be pre-existing rather than a Phase 0
fault: the legacy workbook may have been moved or unshared. That is decision **D4** in the
plan. Note it and carry on.

---

### Step 7. Run `showPortalUrl()`

Gets the **test** portal's web address, needed three times later.

- [ ] Click **`Code.gs`** in the file list. Dropdown → **`showPortalUrl`** → **▷ Run**.

**It worked if** the log prints a URL — **either** `/exec` **or** `/dev` is a pass.

- [ ] **Copy that URL somewhere.** Steps 11, 12b and 15 need it.

**If it ends in `/dev`, that is expected and fine.** The test project has no versioned web app
deployment (confirmed 2026-08-14: Manage deployments shows "This project has not been deployed
yet"). `showPortalUrl()` will say *"This does NOT end in /exec, so it is not a web app
deployment"* and stop short of `READY`. It is correctly describing the project; it is wrong to
treat that as a problem to fix. Recorded as **M10** in [FINDINGS.md](FINDINGS.md).

**`/dev` is the better URL for this checklist anyway:**

| URL | Serves | Who can open it |
|---|---|---|
| `/dev` | **HEAD — the code you just pushed** | anyone with edit access to the script |
| `/exec` | a **pinned version**, fixed when the deployment was created | per the deployment's access setting |

`clasp push` updates HEAD but **does not** update a deployment. So verifying on `/exec` would
require a manual "Manage deployments → New version" after every push, and forgetting it would
verify the *old* code and report a false pass. `/dev` is always current and cannot do that.

**Do not create a deployment to satisfy this step.** You will need one at Phase 7 — its
verification opens the portal at `/exec` to check the `X-Frame-Options` change, which `/dev`
cannot test — and anything involving a second user needs one too, since `/dev` requires script
edit access. Create it there, deliberately, with the redeploy-after-push rule written down.

Note this is the **test** portal. Step 12a needs the **production** portal instead — see that step
for how to get it, which is *not* by running this function.

---

### Step 8. Run `testMyPermissions()`

- [ ] Click **`auth.gs`**. Dropdown → **`testMyPermissions`** → **▷ Run**.

**It worked if** the log shows your email, your role, and a list of what you can do — with no
error.

---

### Step 9. Run `testApiPayload()` — in PRODUCTION, then in test

Times a full page load's worth of server work. This is the first half of
[BASELINE.md](BASELINE.md), and it needs **both** environments — production for the figures that
describe real users, test for the figures each later phase compares itself against.

**9a — production.** Open the production project:

```
https://script.google.com/d/1JHYSultHlRsjZkWcJ9JJ8CNEsBFCUYJVe_n1tENX14_g48ohbe80rGWw/edit
```

- [ ] Click **`Code.gs`**. Dropdown → **`testApiPayload`** → **▷ Run**. May take 10–40 seconds.
- [ ] Run it **at least three times** without changing anything.

**Every run is cold — there is no warm run.** All the caches in this codebase are module-level and
reset between executions (`utils.js` §3; no `CacheService` anywhere — finding P6), so run 2 gets no
benefit from run 1. Repeat runs measure **variance, not warm-up**, and the variance is large:
production's first two runs came out 680 ms and 1,516 ms, a 2.2 × spread on identical work.

Take the **median** of your runs, and record the range alongside it. Do not treat the fastest run
as the baseline, and do not expect later runs to be faster than earlier ones.

**This is safe.** `testApiPayload()` calls `initApp()` and `loadAllAppData()`, which read. The
only write in that path is `recordLogin_` stamping one `Last_Login_TS` cell, throttled to once
per 15 minutes — identical to opening the portal. It does not touch a rate, a mix or an actual.

**9b — test.** Back in the test project, same steps, twice again. Do this **before** step 14 sets
the `SPREADSHEET_ID` property, or the row counts will describe the throwaway copy instead.

**It worked if** each run ends with `API READY` — or with
`API WORKS, but slower than expected`, which is also a pass, being a speed comment rather than an
error. Expect test to be markedly slower than production; that is P14, not a regression.

- [ ] **Record in [BASELINE.md](BASELINE.md):** from **both** runs in **both** environments, the
      `initApp (Nms)` and `loadAllAppData (Nms)` figures — the file has a four-column table. From
      the second production run, both `payload size` lines and every row count (`base rates`,
      `surcharges`, `output rows` and the rest).
- [ ] **Sanity check:** the payload sizes should match between production and test, because test
      currently reads the production spreadsheet. If they differ, stop — that assumption is wrong
      and several other figures depend on it.

---

### Step 10. Run `previewOutput()` in PRODUCTION — optional

Confirms the forecast maths is untouched. Writes nothing, by design — it recomputes the model and
diffs it against the published `OUTPUT` tab without touching either.

Run it in **production**, where the published `OUTPUT` is the one that matters.

- [ ] In the production project, click **`output.gs`**. Dropdown → **`previewOutput`** → **▷ Run**.

**Two possible results, both acceptable:**

- `Nothing would change. OUTPUT already matches the inputs.` — clean.
- A list of differences — this means the spreadsheet holds edits that were never published.
  **It does not mean Phase 0 broke anything:** this phase changed no calculation code at all.
  Note the number reported; it is a useful "before" figure for Phase 1.

---

## Part C — Confirm the portal still works

### Step 11. Open the portal

The real test of the new `PropertiesService` call, because it now runs on every page load.

- [ ] Paste the URL from step 7 into a browser tab and press Enter. Wait for it to finish loading.
      For the test project this is the `/dev` URL, which is what you want — it serves the code you
      pushed at step 2.

**It worked if all four are true:**

1. The dashboard appears — no blank page, no grey Google error page.
2. There is **no red error box** on the screen.
3. "Last published" shows a date, and it is the same one it showed before today.
4. Clicking each tab along the top — **Rates**, **Mixes**, **Actuals**, **History**,
   **Admin** — renders a screen each time.

**If the page is blank or shows a Google error** — stop. This is most likely the manifest-scope
failure mode described at the top of this file.

**If `/dev` itself refuses to open** (rather than opening and failing), that is the one case where
the missing deployment genuinely blocks you: create a web app deployment in the test project
(**Deploy → New deployment →** gear **→ Web app →** Execute as **Me** → Who has access **Anyone
within HeliosX**), then use its `/exec` URL — and remember to create a new version after every
later `clasp push`, or you will be verifying stale code.

---

### Step 12. Record the browser timings — can be done later

The second half of [BASELINE.md](BASELINE.md), and what decides whether Phases 3 and 5 are
worth doing at all. Steps 1–11 and 13–16 are the pass/fail verification; this step is
measurement only.

**Two portals, on purpose.** The read timings come from **production**, because that is what real
users experience. The save is measured in **test**, because it is the only measurement that writes
and a rate save in production would mutate real forecast data just to take a timing.

#### 12a — read timings, in the PRODUCTION portal

- [ ] Get production's portal address from **Deploy → Manage deployments** in the **production**
      Apps Script editor, and copy the **Web app URL** shown there. That listing is authoritative.
      Do **not** use the URL from step 7 — that one is the test portal.

      Use Manage deployments rather than `showPortalUrl()`: `ScriptApp.getService().getUrl()` can
      report the `/dev` URL when it is called from the editor, so it is not a reliable way to
      discover the real `/exec` address. For a *measurement* the `/exec` URL is the one you want —
      it serves the pinned version real users actually hit.
- [ ] Open it, press **F12**. A developer panel opens.
- [ ] Click the **Network** tab inside that panel.
- [ ] Click the **Fetch/XHR** filter button.
- [ ] Tick **Disable cache**.
- [ ] Reload the page once and record the page-load figures.
- [ ] Go to the **Rates** tab. Change the **segment** dropdown to a different value. New rows
      appear in the panel.
- [ ] **Count the new rows** and read the **Time** column on each.
- [ ] Repeat twice more and use the middle result — Apps Script round-trips are noisy enough that
      one sample means little.
- [ ] Repeat for changing the **route** dropdown.
- [ ] Optionally do the same on **Mixes** → method mix (Regime dropdown) and letter/parcel.

**All of the above is read-only.** Changing dropdowns issues reads. **Do not save anything in the
production portal.**

#### 12b — the save, in the TEST portal

- [ ] Open the **test** portal — the URL from step 7 (`/dev` for this project).
- [ ] Same F12 → Network → Fetch/XHR → Disable cache setup.
- [ ] Edit one rate and click **Save**.
- [ ] Record: how many `google.script.run` calls fired, **which functions and in what order**,
      whether any two overlap, the largest response size, and the wall-clock to "Saved".

The call count, the function list and the response sizes are all valid for production — they are
client-side behaviour and test reads the same spreadsheet. Only the **wall-clock** is
test-specific, and `BASELINE.md` records it as an upper bound and derives the production figure
from it.

- [ ] Write everything into [BASELINE.md](BASELINE.md), which has a labelled blank for each.

**The two that matter:**

- If a Rates dropdown change **in production** is already **under 1.5 seconds**, Phase 3 is
  cancelled.
- If a save fires **only one request**, Phase 5 is cancelled. This one is sound from test — the
  count does not change between environments.

---

## Part D — Prove the switch actually works

Everything so far proves Phase 0 broke nothing. These steps prove it does something.

### Step 13. Make a throwaway copy of the spreadsheet

- [ ] Open the `url` that step 4 logged (the production spreadsheet).
- [ ] Menu **File** → **Make a copy**.
- [ ] Name it `Postage Forecast Portal — PHASE 0 TEST COPY`.
- [ ] Click **Make a copy**. It opens in a new tab.
- [ ] Look at that tab's address bar. It reads
      `https://docs.google.com/spreadsheets/d/`**`XXXXXXXX`**`/edit`. **Copy the `XXXXXXXX`
      part** — the long string between `/d/` and `/edit`. That is the copy's ID.

> ⚠️ This copy contains real postage cost forecast data. Keep it in your own Drive, do not
> share it, and delete it at step 16.

---

### Step 14. Set the Script Property

- [ ] Go back to the Apps Script editor tab.
- [ ] In the **far-left icon strip**, click the **gear icon (⚙ Project Settings)**.
- [ ] Scroll down to the **Script Properties** section.
- [ ] Click **Add script property**.
- [ ] In **Property**, type exactly: `SPREADSHEET_ID` — capitals, one underscore, no spaces.
- [ ] In **Value**, paste the copy's ID from step 13.
- [ ] Click **Save script properties**.

**It worked if** the page now shows a row with `SPREADSHEET_ID` on the left and the copy's ID
on the right.

---

### Step 15. Prove the code followed the property

- [ ] In the far-left icon strip, click the **`<>` (Editor)** icon to return to the code.
- [ ] Click **`utils.gs`** → dropdown → **`showEnvironment`** → **▷ Run**.

**It worked if all three of these changed:**

| Line | Must now say |
|---|---|
| `spreadsheet ID` | the copy's ID — **not** `1UgIBKzJ…` |
| `resolved from` | `the SPREADSHEET_ID Script Property` |
| `file name` | `Postage Forecast Portal — PHASE 0 TEST COPY` |

**That one line — `resolved from: the SPREADSHEET_ID Script Property` — is the entire
deliverable of Phase 0.**

- [ ] Confirm the code can read the whole copy, not merely open it: click **`setup.gs`** →
      dropdown → **`setupVerify`** → **▷ Run**. Read-only.

**It worked if** the log reports the tabs as present with no `MISSING TAB` lines, and the row
counts match what production had.

- [ ] Reload the portal URL from step 7.

**It worked if** the portal loads normally. It will look identical, because the copy is
identical — that is expected, and it is why the two checks above are the real proof, not this
one.

**Optional decisive proof, to see it with your own eyes:** in the **copy** (not production),
go to the `High_Level_IDs` tab and type `PHASE 0 TEST COPY` into the `Notes` cell of the first
data row. Reload the portal, go to **Admin → High Level & Modelling IDs**, and that text
should appear. It is a throwaway copy, so nothing is at risk.

---

### Step 16. Choose the end state and tidy up

Decide one of two things.

**Option A — recommended.** Keep the property, but point it at a *proper* test spreadsheet
rather than the throwaway copy. This is what stops the test project writing into production,
and it is what makes Phases 1–8 safe to verify — several of them need the `test…Write()`
functions, which write real rows. Repeat step 14 with the real test copy's ID.

**Option B — the plan as written.** Remove the property and return to the fallback:

- [ ] ⚙ **Project Settings** → **Script Properties**.
- [ ] Click the **pencil/edit icon**, then the **bin icon** next to the `SPREADSHEET_ID` row.
- [ ] Click **Save script properties**.

**Then, whichever was chosen:**

- [ ] Run `showEnvironment()` one final time (`utils.gs` → dropdown → **▷ Run**) so the last
      entry in the execution log is a true record of where the project points. Read the
      `resolved from` line and confirm it says what you intended.
- [ ] Delete the throwaway copy: Google Drive → right-click
      `Postage Forecast Portal — PHASE 0 TEST COPY` → **Move to bin** → then empty the bin.

---

## What "Phase 0 passed" means

| Steps | Project | Proves |
|---|---|---|
| 1–3 | test | the new code is actually on Google's servers, in the right file order |
| 4 | test | with no property set, the code is on the same spreadsheet as before — nothing moved |
| 5–8 | test | the diagnostics and the legacy-workbook reads are unaffected |
| 10 | production | the forecast maths is unaffected |
| 11 | test | the portal still loads — `doGet` survived the new `PropertiesService` call |
| 15 | test | setting one property moves the whole app to a different spreadsheet |
| 9, 12 | production + test | [BASELINE.md](BASELINE.md) is filled in, so Phases 3 and 5 can be re-scored |

Steps 1–8, 10, 11 and 13–16 are pass/fail. Steps 9 and 12 are measurement — Phase 0 is not
finished until they are done, because two later phases depend on those numbers, but nothing is
broken while they are outstanding.

**Phase 0's two halves, stated plainly:** the code change is verified by the pass/fail steps and is
already pushed. The measurement half exists to decide whether Phases 3 and 5 are worth doing at
all, and it needs production for the read figures because test is ~6 × slower on reads (P14).

---

## If you need to undo Phase 0

The code change is one commit and touches no data, no schema and no spreadsheet contents.

```
git log --oneline -3
git revert <the Phase 0 commit id>
clasp push
```

Because the fallback literals are the original production IDs, reverting cannot strand the
app on the wrong spreadsheet. If a `SPREADSHEET_ID` property has been set, **delete it too**
(step 16, Option B) — after a revert the code no longer reads it, so leaving it in place
would be misleading rather than dangerous.

**One thing a revert cannot undo: the project's file order.** `clasp push` sets the order of
files in the Apps Script project every time it runs, and that order is the order the files are
evaluated in. A revert changes file *contents*; it does not restore a previous *order*. If the
project ever fails with a `ReferenceError` naming a constant from another file, that is an
ordering problem and reverting will reproduce it exactly. Keep `"filePushOrder": ["utils.js"]`
in `.clasp.json` — that is what holds the order steady across every push. See FINDINGS.md M9.
