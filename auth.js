/**
 * Postage Forecast Portal — auth.gs
 *
 * Who is this, and what are they allowed to do?
 *
 * The trust model assumes the browser is hostile. Anything the client sends can
 * be forged, so permissions are re-derived from the spreadsheet on every single
 * call. A flag passed in from the page is never believed.
 *
 * Two layers:
 *   1. Portal_Roles   — the ceiling. What this ROLE may ever do.
 *   2. Scope_Mapping  — the slice. Which brands, geos or carriers this PERSON
 *                       may see and edit.
 *
 * A role with All_Access skips layer 2 entirely.
 */

let _permsCache_ = null;


// ─────────────────────────────────────────────────────────────────────────────
// IDENTITY
// ─────────────────────────────────────────────────────────────────────────────

function getActiveEmail_() {
  try {
    const e = Session.getActiveUser().getEmail();
    return e ? e.toLowerCase().trim() : '';
  } catch (err) {
    return '';
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// PERMISSIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Everything the current user is allowed to do. Memoised for the execution.
 *
 * Returns { email, name, role, active, found, capabilities{}, scopes{}, allAccess }
 */
function getUserPermissions_() {
  if (_permsCache_) return _permsCache_;

  const email = getActiveEmail_();
  const perms = {
    email: email, name: '', role: '', active: false, found: false,
    allAccess: false, capabilities: {}, scopes: null, tabVisibility: ''
  };

  if (!email) {
    // Happens on a personal Gmail deployment, or if the userinfo.email scope
    // was not granted. Nothing can be authorised without knowing who this is.
    perms.reason = 'Could not determine your Google account.';
    _permsCache_ = perms;
    return perms;
  }

  // ---- find the person ---------------------------------------------------
  const pData = getAllData_(SHEET.PERMISSIONS), P = COL.PERMISSIONS;
  for (let i = 1; i < pData.length; i++) {
    if (safeStr(pData[i][P.Email]).toLowerCase() !== email) continue;
    perms.found         = true;
    perms.name          = safeStr(pData[i][P.Display_Name]) || email;
    perms.role          = safeStr(pData[i][P.Role]).toUpperCase();
    perms.active        = safeBool(pData[i][P.Active]);
    perms.tabVisibility = safeStr(pData[i][P.Tab_Visibility]);
    perms.lastLoginRaw  = pData[i][P.Last_Login_TS];
    perms.rowIndex      = i + 1;
    break;
  }

  // ---- the bootstrap escape hatch ----------------------------------------
  // Without this, an owner who forgets to add themselves to Permissions locks
  // themselves out of their own portal and has to edit the sheet by hand.
  const bootstrap = safeStr(configStr('BOOTSTRAP_OWNER_EMAIL', '')).toLowerCase();
  if (bootstrap && bootstrap === email && (!perms.found || !perms.active)) {
    perms.found = true; perms.active = true; perms.role = 'ADMIN';
    perms.name = perms.name || 'Owner (bootstrap)';
    perms.bootstrap = true;
  }

  if (!perms.found)  { perms.reason = 'Your email is not in the Permissions list.'; _permsCache_ = perms; return perms; }
  if (!perms.active) { perms.reason = 'Your access has been switched off.';         _permsCache_ = perms; return perms; }

  // ---- what the role allows ----------------------------------------------
  const rData = getAllData_(SHEET.PORTAL_ROLES), R = COL.PORTAL_ROLES;
  let roleFound = false;
  for (let i = 1; i < rData.length; i++) {
    if (safeStr(rData[i][R.Role]).toUpperCase() !== perms.role) continue;
    roleFound = true;
    perms.allAccess = safeBool(rData[i][R.All_Access]);
    perms.capabilities = {
      write:          safeBool(rData[i][R.Write_Access]),
      editRates:      safeBool(rData[i][R.Can_Edit_Rates]),
      editMixes:      safeBool(rData[i][R.Can_Edit_Mixes]),
      editStructure:  safeBool(rData[i][R.Can_Edit_Structure]),
      runCalc:        safeBool(rData[i][R.Can_Run_Calc]),
      publishOutput:  safeBool(rData[i][R.Can_Publish_Output]),
      manageUsers:    safeBool(rData[i][R.Can_Manage_Users]),
      viewAudit:      safeBool(rData[i][R.Can_View_Audit])
    };
    break;
  }
  if (!roleFound) {
    perms.reason = 'Role "' + perms.role + '" is not defined in Portal_Roles.';
    perms.active = false;
    _permsCache_ = perms;
    return perms;
  }

  // ---- which slice of the data -------------------------------------------
  perms.scopes = perms.allAccess ? null : loadScopes_(email);

  _permsCache_ = perms;
  return perms;
}


/** Scope rows for one person, grouped by type. */
function loadScopes_(email) {
  const data = getAllData_(SHEET.SCOPE_MAPPING), S = COL.SCOPE_MAPPING;
  const scopes = { view: {}, edit: {}, any: false };

  for (let i = 1; i < data.length; i++) {
    if (safeStr(data[i][S.Email]).toLowerCase() !== email) continue;
    if (!safeBool(data[i][S.Active])) continue;

    const type  = safeStr(data[i][S.Scope_Type]).toUpperCase();
    const value = normKey(data[i][S.Scope_Value]);
    if (!type || !value) continue;
    scopes.any = true;

    if (safeBool(data[i][S.Can_View])) (scopes.view[type] = scopes.view[type] || {})[value] = true;
    if (safeBool(data[i][S.Can_Edit])) (scopes.edit[type] = scopes.edit[type] || {})[value] = true;
  }
  return scopes;
}


// ─────────────────────────────────────────────────────────────────────────────
// GATES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * The front door. Every entry point calls this first.
 * Throws with a readable message if the person may not be here at all.
 */
function requirePermissions_() {
  const perms = getUserPermissions_();
  if (!perms.found || !perms.active) {
    logAudit_('DENIED', 'PORTAL', '', '', '', '', perms.reason || 'no access', false);
    throw new Error(perms.reason || 'You do not have access to this portal.');
  }
  return perms;
}

/** Require one named capability. */
function requireCapability_(perms, capability, whatFor) {
  if (perms.allAccess) return true;
  if (perms.capabilities[capability]) return true;
  logAudit_('DENIED', 'CAPABILITY', capability, '', '', '', whatFor || capability, false);
  throw new Error('Your role (' + perms.role + ') cannot ' +
                  (whatFor || capability) + '.');
}

function requireWrite_(perms)         { return requireCapability_(perms, 'write', 'make changes'); }
function requireEditRates_(perms)     { requireWrite_(perms); return requireCapability_(perms, 'editRates', 'edit rates'); }
function requireEditMixes_(perms)     { requireWrite_(perms); return requireCapability_(perms, 'editMixes', 'edit mixes'); }
function requireEditStructure_(perms) { requireWrite_(perms); return requireCapability_(perms, 'editStructure', 'change High Level IDs or Modelling IDs'); }
function requireRunCalc_(perms)       { return requireCapability_(perms, 'runCalc', 'run the calculation'); }
function requirePublish_(perms)       { return requireCapability_(perms, 'publishOutput', 'publish the forecast'); }
function requireManageUsers_(perms)   { return requireCapability_(perms, 'manageUsers', 'manage users'); }
function requireViewAudit_(perms)     { return requireCapability_(perms, 'viewAudit', 'view the change history'); }


// ─────────────────────────────────────────────────────────────────────────────
// SCOPE — which rows this person may touch
// ─────────────────────────────────────────────────────────────────────────────

/** Attributes of every High Level ID, for scope decisions. Memoised. */
let _hlMetaCache_ = null;
function hlMeta_() {
  if (_hlMetaCache_) return _hlMetaCache_;
  const out = {}, data = getAllData_(SHEET.HIGH_LEVEL_IDS), H = COL.HIGH_LEVEL_IDS;
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][H.High_Level_ID]);
    if (!id) continue;
    out[id] = { brand: normKey(data[i][H.Brand]), geo: normKey(data[i][H.Geo]) };
  }
  _hlMetaCache_ = out;
  return out;
}

let _midMetaCache_ = null;
function midMeta_() {
  if (_midMetaCache_) return _midMetaCache_;
  const out = {}, data = getAllData_(SHEET.MODELLING_IDS), M = COL.MODELLING_IDS;
  for (let i = 1; i < data.length; i++) {
    const id = safeInt(data[i][M.Modelling_ID]);
    if (!id) continue;
    out[id] = { hlId: safeInt(data[i][M.High_Level_ID]), carrier: normKey(data[i][M.Carrier_Code]) };
  }
  _midMetaCache_ = out;
  return out;
}

/** Does a scope set of this type permit this value? */
function scopeAllows_(scopeSet, type, value) {
  const s = scopeSet[type];
  if (!s) return null;               // no rule of this type — caller decides
  return !!(s['*'] || s[normKey(value)]);
}

/**
 * Can this person see this High Level ID?
 *
 * With no scope rows at all, behaviour comes from Config SCOPE_DEFAULT_ALLOW,
 * which ships as TRUE. That keeps a fresh install usable, and tightening it is
 * one config change once you decide who owns what.
 */
function canSeeHighLevelId_(perms, hlId, forEdit) {
  if (perms.allAccess) return true;
  const scopes = perms.scopes;
  if (!scopes || !scopes.any) return configBool('SCOPE_DEFAULT_ALLOW', true);

  const set = forEdit ? scopes.edit : scopes.view;
  const meta = hlMeta_()[safeInt(hlId)] || {};

  const direct = scopeAllows_(set, 'HIGH_LEVEL_ID', String(hlId));
  if (direct === true) return true;
  const byBrand = scopeAllows_(set, 'BRAND', meta.brand);
  if (byBrand === true) return true;
  const byGeo = scopeAllows_(set, 'GEO', meta.geo);
  if (byGeo === true) return true;

  // If rules of these types exist and none matched, it is a no.
  return (direct === null && byBrand === null && byGeo === null)
         ? configBool('SCOPE_DEFAULT_ALLOW', true) : false;
}

/** Modelling ID visibility: its High Level ID, plus any carrier restriction. */
function canSeeModellingId_(perms, modellingId, forEdit) {
  if (perms.allAccess) return true;
  const meta = midMeta_()[safeInt(modellingId)];
  if (!meta) return false;
  if (!canSeeHighLevelId_(perms, meta.hlId, forEdit)) return false;

  const scopes = perms.scopes;
  if (!scopes || !scopes.any) return true;
  const byCarrier = scopeAllows_(scopes[forEdit ? 'edit' : 'view'], 'CARRIER', meta.carrier);
  return (byCarrier === null) ? true : byCarrier;
}

function assertCanEditHighLevelId_(perms, hlId) {
  if (canSeeHighLevelId_(perms, hlId, true)) return true;
  logAudit_('DENIED', 'HIGH_LEVEL_ID', hlId, '', '', '', 'edit outside scope', false);
  throw new Error('High Level ID ' + hlId + ' is outside the area you can edit.');
}

function assertCanEditModellingId_(perms, modellingId) {
  if (canSeeModellingId_(perms, modellingId, true)) return true;
  logAudit_('DENIED', 'MODELLING_ID', modellingId, '', '', '', 'edit outside scope', false);
  throw new Error('Modelling ID ' + modellingId + ' is outside the area you can edit.');
}

/**
 * Authorise an update against the row as it stands, not only against the
 * payload — the fix for the "re-parent a row you cannot see" hole.
 *
 * A save carries the foreign key the caller wants the row to end up with.
 * Checking only that key authorises the destination and never the source, so a
 * caller who may edit segment B can send { id: <a row in segment A>, hlId: B }
 * and quietly move a row they were never allowed to touch. Every update path
 * must therefore also authorise the owner already on the row, which is what the
 * delete paths have always done by reading the row before asserting.
 *
 * A no-op for a create — there is no existing row to own. Reads through the
 * execution cache, so inside a write that was going to scan the table anyway it
 * costs nothing. A missing row is left alone: the caller's own "no longer
 * exists" error says something more useful than a scope refusal would.
 *
 * @param {Object} perms          from requirePermissions_()
 * @param {string} sheetName      the tab holding the row being updated
 * @param {number} idColIndex     0-based index of that table's ID column
 * @param {*}      id             ID from the payload; falsy means create
 * @param {number} ownerColIndex  0-based index of High_Level_ID or Modelling_ID
 * @param {string} ownerKind      'HIGH_LEVEL_ID' or 'MODELLING_ID'
 */
function assertCanEditRowOwner_(perms, sheetName, idColIndex, id,
                                ownerColIndex, ownerKind) {
  const want = safeInt(id);
  if (!want) return true;

  const data = getAllData_(sheetName);
  for (let i = 1; i < data.length; i++) {
    if (safeInt(data[i][idColIndex]) !== want) continue;
    const owner = safeInt(data[i][ownerColIndex]);
    return (ownerKind === 'MODELLING_ID')
      ? assertCanEditModellingId_(perms, owner)
      : assertCanEditHighLevelId_(perms, owner);
  }
  return true;
}

function assertCanSeeModellingId_(perms, modellingId) {
  if (canSeeModellingId_(perms, modellingId, false)) return true;
  logAudit_('DENIED', 'MODELLING_ID', modellingId, '', '', '', 'view outside scope', false);
  throw new Error('Modelling ID ' + modellingId + ' is outside the area you can see.');
}

/** The High Level IDs this person may see. Used to filter before serialising. */
function visibleHighLevelIds_(perms) {
  const out = {};
  const meta = hlMeta_();
  for (const id in meta) if (canSeeHighLevelId_(perms, safeInt(id), false)) out[safeInt(id)] = true;
  return out;
}

function visibleModellingIds_(perms) {
  const out = {};
  const meta = midMeta_();
  for (const id in meta) if (canSeeModellingId_(perms, safeInt(id), false)) out[safeInt(id)] = true;
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// SIGN IN
// ─────────────────────────────────────────────────────────────────────────────

const LOGIN_THROTTLE_MINUTES = 15;

/**
 * Stamp Last_Login_TS, but not on every page load.
 *
 * This used to re-read the whole Permissions tab and then write, which cost
 * about 670ms — most of the time it took the portal to start. Two changes:
 *
 *   - the row index and existing timestamp come from the data already cached
 *     by getUserPermissions_, so there is no second read
 *   - the write is skipped if the last one was recent, so a normal session
 *     writes once rather than once per screen
 *
 * A login time accurate to the quarter hour is plenty for "who is using this".
 */
function recordLogin_(perms) {
  try {
    if (!perms.rowIndex) return;

    const raw = perms.lastLoginRaw;
    let lastMs = 0;
    if (isDate_(raw)) lastMs = raw.getTime();
    else if (typeof raw === 'number' && raw > SERIAL_EPOCH_OFFSET) {
      lastMs = Math.round((raw - SERIAL_EPOCH_OFFSET) * MS_PER_DAY);
    }
    if (lastMs && (Date.now() - lastMs) < LOGIN_THROTTLE_MINUTES * 60000) return;

    getSheet_(SHEET.PERMISSIONS)
      .getRange(perms.rowIndex, COL.PERMISSIONS.Last_Login_TS + 1)
      .setValue(new Date());
    invalidateSheetCache_(SHEET.PERMISSIONS);
  } catch (e) { /* never block a login over its own bookkeeping */ }
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — run from the editor
// ─────────────────────────────────────────────────────────────────────────────

function testMyPermissions() {
  requireMaintenance_();
  Logger.log('=== WHO AM I ===');
  const email = getActiveEmail_();
  Logger.log('  Google account: ' + (email || '(could not be determined)'));

  if (!email) {
    Logger.log('');
    Logger.log('  The script cannot see your email address. On a Workspace account this');
    Logger.log('  normally means the userinfo.email scope was not granted — check');
    Logger.log('  appsscript.json, save, and run again.');
    return;
  }

  const perms = getUserPermissions_();
  Logger.log('  In Permissions : ' + (perms.found ? 'yes' : 'NO'));
  if (perms.bootstrap) Logger.log('  (admitted via BOOTSTRAP_OWNER_EMAIL)');

  if (!perms.found) {
    Logger.log('');
    Logger.log('  Add a row to the Permissions tab:');
    Logger.log('    Email = ' + email + '   Role = ADMIN   Active = TRUE');
    return;
  }

  Logger.log('  Name           : ' + perms.name);
  Logger.log('  Role           : ' + perms.role);
  Logger.log('  Active         : ' + perms.active);
  if (!perms.active) { Logger.log('  ' + perms.reason); return; }

  Logger.log('');
  Logger.log('--- what you can do ---');
  Logger.log('  all access     : ' + perms.allAccess);
  const labels = {
    write: 'make changes', editRates: 'edit rates', editMixes: 'edit mixes',
    editStructure: 'change structure', runCalc: 'run the calculation',
    publishOutput: 'publish the forecast', manageUsers: 'manage users',
    viewAudit: 'view history'
  };
  for (const k in labels) {
    const allowed = perms.allAccess || perms.capabilities[k];
    Logger.log('  ' + (allowed ? 'yes' : ' no') + '  ' + labels[k]);
  }

  Logger.log('');
  Logger.log('--- what you can see ---');
  if (perms.allAccess) {
    Logger.log('  everything (role has All_Access)');
  } else if (!perms.scopes || !perms.scopes.any) {
    Logger.log('  no Scope_Mapping rows; falling back to SCOPE_DEFAULT_ALLOW = ' +
               configBool('SCOPE_DEFAULT_ALLOW', true));
  } else {
    ['view', 'edit'].forEach(which => {
      const set = perms.scopes[which];
      const parts = [];
      for (const type in set) parts.push(type + ': ' + Object.keys(set[type]).join(', '));
      Logger.log('  can ' + which + ' — ' + (parts.join(' | ') || 'nothing'));
    });
  }
  const hl = visibleHighLevelIds_(perms), md = visibleModellingIds_(perms);
  Logger.log('  High Level IDs visible : ' + Object.keys(hl).length + ' of ' + Object.keys(hlMeta_()).length);
  Logger.log('  Modelling IDs visible  : ' + Object.keys(md).length + ' of ' + Object.keys(midMeta_()).length);

  return perms;
}