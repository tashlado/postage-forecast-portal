/**
 * Postage Forecast Portal — Code.gs
 *
 * The front door. Everything the browser calls arrives here.
 *
 * Two loading calls, deliberately:
 *
 *   initApp()          small and fast. Who you are, the dropdown lists, the
 *                      structure. Enough to draw the page frame.
 *   loadAllAppData()   the bulk. Rates, mixes, output, validation, history.
 *
 * Splitting them means the page appears in about a second instead of staring at
 * a blank screen until everything has arrived.
 *
 * Both filter by scope on the SERVER. A user restricted to GB does not receive
 * US rows and then have them hidden — the rows never leave the server. Anything
 * else is theatre, because the payload is one keystroke away in developer tools.
 */

// ─────────────────────────────────────────────────────────────────────────────
// WEB APP ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Postage Forecast')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // index.html arrives in Step 8. Until then show a status page, so the
    // deployment itself can be tested while the interface is still being built.
    return HtmlService.createHtmlOutput(statusPageHtml_(err))
      .setTitle('Postage Forecast — backend');
  }
}

/** Lets index.html pull in shared CSS or JS partials later. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function statusPageHtml_(err) {
  let who = 'not signed in', role = '', detail = '';
  try {
    const p = getUserPermissions_();
    who  = p.email || 'unknown';
    role = p.found ? (p.role + (p.active ? '' : ' (inactive)')) : 'not in Permissions';
  } catch (e) {
    detail = e.message;
  }
  const missingIndex = String(err && err.message || '').indexOf('index') >= 0;

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
    'background:#F7F8FA;color:#0B0B0F;margin:0;padding:48px 24px;line-height:1.5}' +
    '.card{max-width:520px;margin:0 auto;background:#fff;border:1px solid #E4E6EC;' +
    'border-radius:12px;padding:28px}' +
    'h1{font-size:19px;font-weight:500;margin:0 0 4px}' +
    'p{font-size:14px;color:#3F4451;margin:0 0 16px}' +
    'dl{display:grid;grid-template-columns:auto 1fr;gap:6px 16px;font-size:13px;margin:0}' +
    'dt{color:#6B7280}dd{margin:0;font-variant-numeric:tabular-nums}' +
    '.tag{display:inline-block;background:#F0F0FF;color:#0000C2;border-radius:4px;' +
    'padding:2px 8px;font-size:12px}' +
    '</style></head><body><div class="card">' +
    '<h1>Postage Forecast</h1>' +
    '<p>' + (missingIndex
      ? 'The backend is deployed and running. The interface itself is added in Step 8.'
      : 'The page could not be loaded.') + '</p>' +
    '<dl>' +
    '<dt>Signed in as</dt><dd>' + escapeHtml_(who) + '</dd>' +
    '<dt>Role</dt><dd><span class="tag">' + escapeHtml_(role || 'none') + '</span></dd>' +
    '<dt>Status</dt><dd>backend ready</dd>' +
    '</dl>' +
    (detail ? '<p style="margin-top:16px;color:#C6273E;font-size:13px">' +
              escapeHtml_(detail) + '</p>' : '') +
    '</div></body></html>';
}

function escapeHtml_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}


// ─────────────────────────────────────────────────────────────────────────────
// INIT — the small, fast first call
// ─────────────────────────────────────────────────────────────────────────────

function initApp() {
  perfReset();

  prewarmSheetCache_([
    SHEET.PERMISSIONS, SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING, SHEET.CONFIG,
    SHEET.DIM_REFERENCE, SHEET.DIM_CARRIER, SHEET.DIM_METHOD, SHEET.DIM_SURCHARGE,
    SHEET.DIM_CALENDAR, SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS,
    SHEET.SCENARIOS, SHEET.CALC_RUNS, SHEET.AUDIT_LOG, SHEET.GUIDE
  ]);
  perfMark('init: sheets fetched');

  const perms = requirePermissions_();
  recordLogin_(perms);
  perfMark('init: permissions');

  const visibleHl  = visibleHighLevelIds_(perms);
  const visibleMid = visibleModellingIds_(perms);

  const payload = {
    user: {
      email: perms.email, name: perms.name, role: perms.role,
      allAccess: perms.allAccess, capabilities: perms.capabilities,
      tabVisibility: perms.tabVisibility
    },
    reference:     loadReferenceLists_(),
    carriers:      loadCarriers_(),
    methods:       loadMethods_(),
    surchargeTypes: loadSurchargeTypes_(),
    calendar:      loadCalendarForClient_(),
    highLevelIds:  loadHighLevelIdsForClient_(perms, visibleHl),
    modellingIds:  loadModellingIdsForClient_(perms, visibleMid),
    scenarios:     loadScenarios_(),
    config: {
      horizonStart:  configDate('HORIZON_START', '2026-01-01'),
      horizonMonths: configInt('HORIZON_MONTHS', 36),
      openEndedDate: configDate('OPEN_ENDED_DATE', OPEN_ENDED_DATE),
      autoCalc:      configBool('AUTO_CALC_ON_WRITE', false),
      engineVersion: configStr('ENGINE_VERSION', ENGINE_VERSION),
      /* A link, not a hardcoded URL, so the guide can be moved or replaced in the
         Config tab without a deploy. Empty means the dashboard shows nothing —
         a missing guide should leave no trace rather than a dead link. */
      guideUrl:      configStr('GUIDE_URL', '')
    },
    guide: loadGuideForClient_(),
    status: loadStatus_()
  };
  perfMark('init: payload built');

  payload.perf = perfReport();
  return payload;
}


// ─────────────────────────────────────────────────────────────────────────────
// BULK — everything else, in one call
// ─────────────────────────────────────────────────────────────────────────────

function loadAllAppData() {
  perfReset();

  prewarmSheetCache_([
    SHEET.PERMISSIONS, SHEET.PORTAL_ROLES, SHEET.SCOPE_MAPPING, SHEET.CONFIG,
    SHEET.HIGH_LEVEL_IDS, SHEET.MODELLING_IDS, SHEET.DIM_CALENDAR,
    SHEET.RATE_BASE, SHEET.RATE_SURCHARGE, SHEET.MIX_METHOD,
    SHEET.MIX_LETTERPARCEL, SHEET.MIX_COLDCHAIN,
    SHEET.OUTPUT, SHEET.VALIDATION_RESULTS, SHEET.AUDIT_LOG, SHEET.CALC_RUNS,
    SHEET.ACTUALS, SHEET.SNAPSHOTS
  ]);
  perfMark('bulk: sheets fetched');

  const perms = requirePermissions_();
  const visibleHl  = visibleHighLevelIds_(perms);
  const visibleMid = visibleModellingIds_(perms);
  perfMark('bulk: permissions');

  const data = {
    rateBase:        loadRateBaseForClient_(visibleMid),
    rateSurcharge:   loadRateSurchargeForClient_(visibleMid),
    mixMethod:       loadMixMethodForClient_(visibleMid),
    mixLetterParcel: loadMixLPForClient_(visibleHl),
    mixColdChain:    loadMixCCForClient_(visibleHl),
    output:          loadOutputForClient_(visibleHl),
    actuals:         loadActualsForClient_(visibleHl),
    snapshots:       loadSnapshotsForClient_(),
    validation:      loadValidationForClient_(perms),
    recentAudit:     perms.allAccess || perms.capabilities.viewAudit
                       ? readRecentAudit_(50) : [],
    status:          loadStatus_()
  };
  perfMark('bulk: payload built');

  data.perf = perfReport();
  return data;
}


// ─────────────────────────────────────────────────────────────────────────────
// LOADERS
// ─────────────────────────────────────────────────────────────────────────────

function loadReferenceLists_() {
  const data = getAllData_(SHEET.DIM_REFERENCE), C = COL.DIM_REFERENCE;
  const out = {};
  for (let i = 1; i < data.length; i++) {
    if (!safeBool(data[i][C.Active])) continue;
    const list = safeStr(data[i][C.List_Name]);
    if (!list) continue;
    (out[list] = out[list] || []).push({
      code: safeStr(data[i][C.Code]),
      label: safeStr(data[i][C.Label]),
      sort: safeInt(data[i][C.Sort_Order])
    });
  }
  for (const k in out) out[k].sort((a, b) => a.sort - b.sort);
  return out;
}

function loadCarriers_() {
  const data = getAllData_(SHEET.DIM_CARRIER), C = COL.DIM_CARRIER, out = [];
  for (let i = 1; i < data.length; i++) {
    if (!safeStr(data[i][C.Carrier_Code]) || !safeBool(data[i][C.Active])) continue;
    out.push({ code: safeStr(data[i][C.Carrier_Code]),
               name: safeStr(data[i][C.Carrier_Name]),
               currency: safeStr(data[i][C.Default_Currency]) });
  }
  return out.sort((a, b) => a.name < b.name ? -1 : 1);
}

function loadMethods_() {
  const data = getAllData_(SHEET.DIM_METHOD), C = COL.DIM_METHOD, out = [];
  for (let i = 1; i < data.length; i++) {
    if (!safeStr(data[i][C.Method_Code]) || !safeBool(data[i][C.Active])) continue;
    out.push({ code: safeStr(data[i][C.Method_Code]),
               carrier: safeStr(data[i][C.Carrier_Code]),
               name: safeStr(data[i][C.Method_Name]),
               serviceLevel: safeStr(data[i][C.Service_Level]) });
  }
  return out.sort((a, b) => (a.carrier + a.name) < (b.carrier + b.name) ? -1 : 1);
}

function loadSurchargeTypes_() {
  const data = getAllData_(SHEET.DIM_SURCHARGE), C = COL.DIM_SURCHARGE, out = [];
  for (let i = 1; i < data.length; i++) {
    if (!safeStr(data[i][C.Surcharge_Code]) || !safeBool(data[i][C.Active])) continue;
    out.push({ code: safeStr(data[i][C.Surcharge_Code]),
               name: safeStr(data[i][C.Surcharge_Name]),
               valueType: safeStr(data[i][C.Value_Type]),
               appliesTo: safeStr(data[i][C.Applies_To]),
               applyOrder: safeInt(data[i][C.Apply_Order]),
               proration: safeStr(data[i][C.Proration]) });
  }
  return out.sort((a, b) => a.applyOrder - b.applyOrder);
}

function loadCalendarForClient_() {
  const data = getAllData_(SHEET.DIM_CALENDAR), C = COL.DIM_CALENDAR, out = [];
  for (let i = 1; i < data.length; i++) {
    const ms = normaliseDate(data[i][C.Month_Start]);
    if (!ms) continue;
    out.push({ dateId: safeInt(data[i][C.Date_ID]),
               monthStart: fmtDate(ms),
               monthEnd: fmtDate(normaliseDate(data[i][C.Month_End])),
               daysInMonth: safeInt(data[i][C.Days_In_Month]),
               label: monthLabel_(data[i][C.Month_Label], ms),
               quarter: safeStr(data[i][C.Quarter]),
               inHorizon: safeBool(data[i][C.In_Horizon]) });
  }
  return out.sort((a, b) => a.dateId - b.dateId);
}

function loadHighLevelIdsForClient_(perms, visible) {
  const data = getAllData_(SHEET.HIGH_LEVEL_IDS), C = COL.HIGH_LEVEL_IDS, out = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.High_Level_ID]);
    if (!id || !visible[id]) continue;
    out.push({
      id: id, code: safeStr(data[i][C.High_Level_Code]),
      brand: safeStr(data[i][C.Brand]), geo: safeStr(data[i][C.Geo]),
      treatmentType: safeStr(data[i][C.Treatment_Type]),
      wlSplit: safeStr(data[i][C.WL_Split]),
      currency: safeStr(data[i][C.Currency]),
      active: safeBool(data[i][C.Active]),
      sortOrder: safeInt(data[i][C.Sort_Order]),
      notes: safeStr(data[i][C.Notes]),
      canEdit: canSeeHighLevelId_(perms, id, true)
    });
  }
  return out.sort((a, b) => a.sortOrder - b.sortOrder || a.id - b.id);
}

/**
 * The Guide tab's blocks, in order.
 *
 * Reads through getAllData_, so a tab that does not exist yet returns nothing
 * rather than throwing — the Guide screen then explains how to fill it in, and
 * initApp is not brought down by a missing reference tab. That matters more than
 * it looks: initApp is the first call the page makes, so anything that throws here
 * takes the whole portal with it.
 *
 * No permission check beyond the portal's own: this is documentation, and it is
 * held in the spreadsheet precisely so that everyone who can open the portal can
 * read it.
 */
function loadGuideForClient_() {
  let data;
  try { data = getAllData_(SHEET.GUIDE); } catch (e) { return []; }
  if (!data || data.length < 2) return [];
  const C = COL.GUIDE, out = [];
  for (let i = 1; i < data.length; i++) {
    if (!safeBool(data[i][C.Active])) continue;
    const heading = safeStr(data[i][C.Heading]);
    const body    = safeStr(data[i][C.Body]);
    if (!heading && !body) continue;          // a blank row is spacing, not content
    out.push({ order: safeNum(data[i][C.Sort_Order]), heading: heading, body: body });
  }
  /* Sorted here rather than trusting sheet order, so inserting a row above another
     does not silently reorder the guide. Stable on equal Sort_Order. */
  return out.map(function (r, i) { return { r: r, i: i }; })
            .sort(function (a, b) { return a.r.order - b.r.order || a.i - b.i; })
            .map(function (x) { return x.r; });
}


function loadModellingIdsForClient_(perms, visible) {
  const data = getAllData_(SHEET.MODELLING_IDS), C = COL.MODELLING_IDS, out = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.Modelling_ID]);
    if (!id || !visible[id]) continue;
    out.push({
      id: id, hlId: safeInt(data[i][C.High_Level_ID]),
      carrier: safeStr(data[i][C.Carrier_Code]),
      method: safeStr(data[i][C.Method_Code]),
      letterParcel: safeStr(data[i][C.Letter_Parcel]),
      code: safeStr(data[i][C.Modelling_Code]),
      active: safeBool(data[i][C.Active]),
      notes: safeStr(data[i][C.Notes]),
      canEdit: canSeeModellingId_(perms, id, true)
    });
  }
  return out.sort((a, b) => a.id - b.id);
}

function loadScenarios_() {
  const data = getAllData_(SHEET.SCENARIOS), C = COL.SCENARIOS, out = [];
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][C.Scenario_ID]);
    if (!id || !safeBool(data[i][C.Active])) continue;
    out.push({ id: id, name: safeStr(data[i][C.Scenario_Name]),
               description: safeStr(data[i][C.Description]),
               parentId: safeInt(data[i][C.Parent_Scenario_ID]),
               isDefault: safeBool(data[i][C.Is_Default]),
               locked: safeBool(data[i][C.Locked]) });
  }
  return out.sort((a, b) => a.id - b.id);
}

/** Dated rows, shared shape for the client. */
function loadRateBaseForClient_(visibleMid) {
  const data = getAllData_(SHEET.RATE_BASE), C = COL.RATE_BASE, out = [];
  for (let i = 1; i < data.length; i++) {
    const mid = safeInt(data[i][C.Modelling_ID]);
    if (!mid || !visibleMid[mid]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.Rate_ID]), modellingId: mid,
               validFrom: fmtDate(data[i][C.Valid_From]),
               validTo: fmtDate(data[i][C.Valid_To]),
               value: safeNum(data[i][C.Base_Rate]),
               currency: safeStr(data[i][C.Currency]),
               scenarioId: safeInt(data[i][C.Scenario_ID]),
               sourceRef: safeStr(data[i][C.Source_Ref]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out;
}

function loadRateSurchargeForClient_(visibleMid) {
  const data = getAllData_(SHEET.RATE_SURCHARGE), C = COL.RATE_SURCHARGE, out = [];
  for (let i = 1; i < data.length; i++) {
    const mid = safeInt(data[i][C.Modelling_ID]);
    if (!mid || !visibleMid[mid]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.Surcharge_Rate_ID]), modellingId: mid,
               code: safeStr(data[i][C.Surcharge_Code]),
               validFrom: fmtDate(data[i][C.Valid_From]),
               validTo: fmtDate(data[i][C.Valid_To]),
               value: safeNum(data[i][C.Value]),
               currency: safeStr(data[i][C.Currency]),
               scenarioId: safeInt(data[i][C.Scenario_ID]),
               sourceRef: safeStr(data[i][C.Source_Ref]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out;
}

function loadMixMethodForClient_(visibleMid) {
  const data = getAllData_(SHEET.MIX_METHOD), C = COL.MIX_METHOD, out = [];
  for (let i = 1; i < data.length; i++) {
    const mid = safeInt(data[i][C.Modelling_ID]);
    if (!mid || !visibleMid[mid]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.Mix_ID]), modellingId: mid,
               regime: safeStr(data[i][C.Temp_Regime]),
               validFrom: fmtDate(data[i][C.Valid_From]),
               validTo: fmtDate(data[i][C.Valid_To]),
               value: safeNum(data[i][C.Mix_Pct]),
               scenarioId: safeInt(data[i][C.Scenario_ID]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out;
}

function loadMixLPForClient_(visibleHl) {
  const data = getAllData_(SHEET.MIX_LETTERPARCEL), C = COL.MIX_LETTERPARCEL, out = [];
  for (let i = 1; i < data.length; i++) {
    const hl = safeInt(data[i][C.High_Level_ID]);
    if (!hl || !visibleHl[hl]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.LP_Mix_ID]), hlId: hl,
               validFrom: fmtDate(data[i][C.Valid_From]),
               validTo: fmtDate(data[i][C.Valid_To]),
               value: safeNum(data[i][C.Letter_Mix_Pct]),
               scenarioId: safeInt(data[i][C.Scenario_ID]),
               notes: safeStr(data[i][C.Notes]) });
  }
  return out;
}

function loadMixCCForClient_(visibleHl) {
  const data = getAllData_(SHEET.MIX_COLDCHAIN), C = COL.MIX_COLDCHAIN, out = [];
  for (let i = 1; i < data.length; i++) {
    const hl = safeInt(data[i][C.High_Level_ID]);
    if (!hl || !visibleHl[hl]) continue;
    if (!safeBool(data[i][C.Active])) continue;
    out.push({ id: safeInt(data[i][C.CC_Mix_ID]), hlId: hl,
               validFrom: fmtDate(data[i][C.Valid_From]),
               validTo: fmtDate(data[i][C.Valid_To]),
               value: safeNum(data[i][C.CC_Mix_Pct]),
               scenarioId: safeInt(data[i][C.Scenario_ID]) });
  }
  return out;
}

function loadOutputForClient_(visibleHl) {
  const data = getAllData_(SHEET.OUTPUT), C = COL.OUTPUT, out = [];
  for (let i = 1; i < data.length; i++) {
    const hl = safeInt(data[i][C.High_Level_ID]);
    if (!hl || !visibleHl[hl]) continue;
    /* Brand, Geo, Treatment_Type and WL_Split are sent as OUTPUT itself holds
       them, not looked up from High_Level_IDs on the client. OUTPUT is a snapshot
       taken at publish time, so if a segment has been renamed since, these are
       what the published forecast actually said — which is what the Output screen
       claims to show. Costs nothing extra to read: the whole row is already here. */
    out.push({ hlId: hl, dateId: safeInt(data[i][C.Date_ID]),
               monthStart: fmtDate(data[i][C.Month_Start]),
               brand: safeStr(data[i][C.Brand]),
               geo: safeStr(data[i][C.Geo]),
               treatmentType: safeStr(data[i][C.Treatment_Type]),
               wlSplit: safeStr(data[i][C.WL_Split]),
               currency: safeStr(data[i][C.Currency]),
               rate: safeNum(data[i][C.Forecast_Rate_Per_Order]),
               scenarioId: safeInt(data[i][C.Scenario_ID]),
               calcRunId: safeInt(data[i][C.Calc_Run_ID]) });
  }
  return out;
}

function loadValidationForClient_(perms) {
  const data = getAllData_(SHEET.VALIDATION_RESULTS), C = COL.VALIDATION_RESULTS, out = [];
  const visibleHl = perms.allAccess ? null : visibleHighLevelIds_(perms);
  for (let i = 1; i < data.length; i++) {
    if (!safeStr(data[i][C.Rule_Code])) continue;
    const hl = safeInt(data[i][C.High_Level_ID]);
    if (visibleHl && hl && !visibleHl[hl]) continue;
    out.push({ id: safeInt(data[i][C.Result_ID]),
               rule: safeStr(data[i][C.Rule_Code]),
               severity: safeStr(data[i][C.Severity]),
               hlId: hl, modellingId: safeInt(data[i][C.Modelling_ID]),
               dateId: safeInt(data[i][C.Date_ID]),
               message: safeStr(data[i][C.Message]) });
  }
  return out;
}

/** Header strip: when was it last published, and how much has changed since. */
function loadStatus_() {
  const runs = getAllData_(SHEET.CALC_RUNS), R = COL.CALC_RUNS;
  let last = null;
  for (let i = 1; i < runs.length; i++) {
    const id = safeInt(runs[i][R.Calc_Run_ID]);
    if (!id) continue;
    if (!last || id > last.id) {
      last = { id: id, ts: fmtTs(runs[i][R.Run_TS]), by: safeStr(runs[i][R.Run_By]),
               rows: safeInt(runs[i][R.Rows_Output]),
               durationMs: safeInt(runs[i][R.Duration_Ms]),
               validationStatus: safeStr(runs[i][R.Validation_Status]) };
    }
  }
  return { lastRun: last, pendingChanges: getPendingChangeCount_() };
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — what is my web app address, and is it deployed correctly?
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Prints the real web app URL.
 *
 * Apps Script will happily deploy a project as a Library, which serves an
 * auto-generated page of function signatures instead of your interface. The two
 * URLs look nothing alike, so printing the right one is the quickest way to tell
 * which you are looking at.
 */
function showPortalUrl() {
  requireMaintenance_();
  Logger.log('=== PORTAL ADDRESS ===');

  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }

  if (!url) {
    Logger.log('  No web app deployment found.');
    Logger.log('');
    Logger.log('  Deploy > New deployment > click the GEAR next to "Select type"');
    Logger.log('  > tick "Web app" > Execute as: Me > Who has access: Anyone within HeliosX');
    return { url: '', ok: false };
  }

  Logger.log('  ' + url);
  Logger.log('');
  const ok = url.indexOf('/exec') >= 0;
  if (ok) {
    Logger.log('  This is a WEB APP address — correct. Open it in a browser.');
  } else {
    Logger.log('  This does NOT end in /exec, so it is not a web app deployment.');
    Logger.log('  Deploy > Manage deployments and check the type says "Web app".');
  }

  Logger.log('');
  Logger.log('--- what the portal needs ---');
  let hasIndex = false;
  try { HtmlService.createHtmlOutputFromFile('index'); hasIndex = true; } catch (e) {}
  Logger.log('  index.html present : ' + (hasIndex ? 'yes' : 'NO — add it as an HTML file named "index"'));
  Logger.log('  doGet present      : yes');

  const renamed = (typeof getAllData_ === 'function') && (typeof getAllData === 'undefined');
  Logger.log('  backend updated    : ' + (renamed ? 'yes' :
             'NO — the Step 8 files have not been re-pasted'));

  Logger.log('');
  Logger.log(ok && hasIndex && renamed
    ? 'READY — open the address above.'
    : 'Fix whatever says NO above, then publish a new version.');
  return { url: url, ok: ok && hasIndex && renamed };
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — exercise the API without a browser
// ─────────────────────────────────────────────────────────────────────────────

function testApiPayload() {
  requireMaintenance_();
  Logger.log('=== API PAYLOAD TEST ===');

  let t0 = Date.now();
  const init = initApp();
  const initMs = Date.now() - t0;

  Logger.log('--- initApp (' + initMs + 'ms) ---');
  Logger.log('  user            : ' + init.user.name + '  role ' + init.user.role +
             (init.user.allAccess ? '  (all access)' : ''));
  Logger.log('  reference lists : ' + Object.keys(init.reference).length);
  Logger.log('  carriers        : ' + init.carriers.length);
  Logger.log('  methods         : ' + init.methods.length);
  Logger.log('  surcharge types : ' + init.surchargeTypes.length);
  Logger.log('  calendar months : ' + init.calendar.length);
  Logger.log('  High Level IDs  : ' + init.highLevelIds.length);
  Logger.log('  Modelling IDs   : ' + init.modellingIds.length);
  Logger.log('  scenarios       : ' + init.scenarios.length);
  Logger.log('  payload size    : ~' + Math.round(JSON.stringify(init).length / 1024) + ' KB');

  t0 = Date.now();
  const bulk = loadAllAppData();
  const bulkMs = Date.now() - t0;

  Logger.log('');
  Logger.log('--- loadAllAppData (' + bulkMs + 'ms) ---');
  Logger.log('  base rates      : ' + bulk.rateBase.length);
  Logger.log('  surcharges      : ' + bulk.rateSurcharge.length);
  Logger.log('  method mixes    : ' + bulk.mixMethod.length);
  Logger.log('  letter/parcel   : ' + bulk.mixLetterParcel.length);
  Logger.log('  cold chain      : ' + bulk.mixColdChain.length);
  Logger.log('  output rows     : ' + bulk.output.length);
  Logger.log('  actuals         : ' + bulk.actuals.length);
  Logger.log('  snapshots       : ' + bulk.snapshots.length);
  Logger.log('  validation      : ' + bulk.validation.length);
  Logger.log('  recent audit    : ' + bulk.recentAudit.length);
  Logger.log('  payload size    : ~' + Math.round(JSON.stringify(bulk).length / 1024) + ' KB');

  Logger.log('');
  Logger.log('--- status ---');
  if (bulk.status.lastRun) {
    const r = bulk.status.lastRun;
    Logger.log('  last published  : run ' + r.id + ' by ' + r.by + ' on ' + r.ts +
               '  (' + r.rows + ' rows, ' + r.validationStatus + ')');
  } else {
    Logger.log('  last published  : never');
  }
  Logger.log('  changes since   : ' + bulk.status.pendingChanges);

  const total = initMs + bulkMs;
  Logger.log('');
  Logger.log('  TOTAL: ' + total + 'ms' +
             (total < 3000 ? '   good' : total < 5000 ? '   acceptable' : '   slower than target'));
  Logger.log('');
  Logger.log(total < 5000 ? 'API READY' : 'API WORKS, but slower than expected — send me this log');
  return { initMs: initMs, bulkMs: bulkMs };
}