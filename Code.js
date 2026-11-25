/** @OnlyCurrentDoc */

/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * MASTER CONTROLLER, CLIENT ISOLATION ENGINE & UNIFIED TOP MENU
 * ============================================================================
 * File: Code.js / Code.gs
 * ============================================================================
 */

/**
 * 39 Canonical Header Columns for Persona Mapper Settings (LOCKED SCHEMA)
 */
const PERSONA_MAPPER_39_HEADERS = [
  'Campaign ID',
  'Campaign Name',
  'Persona Mapper Enabled',
  'AI Provider',
  'AI Model',
  'Persona Titles',
  'Search Keywords',
  'Campaign Focus',
  'Evidence Period Months',
  'Research Depth',
  'Include ESG Reporting Obligations',
  'Include ESG Reporting Maturity',
  'Include Supply Chain Complexity',
  'Include Scope 3 Exposure',
  'Include Sustainability Data Challenges',
  'Include Digital Transformation',
  'Include ESG Governance',
  'Include Business Triggers',
  'Include eValuater Solution Fit',
  'Weight ESG Reporting Relevance',
  'Weight Supply Chain Complexity',
  'Weight Sustainability Data Needs',
  'Weight Digital Transformation Relevance',
  'Weight Business Triggers',
  'Weight eValuater Solution Fit',
  'Minimum Company Opportunity Score',
  'Minimum Persona Confidence',
  'Minimum Email Priority Score',
  'Human Approval Required',
  'Auto Send Enabled',
  'Email Priority Weight Company Score',
  'Email Priority Weight Persona Fit',
  'Email Priority Weight Buying Signal',
  'Notes',
  'Max Companies Per Run',
  'Email Verification Provider',
  'Email Verification Endpoint',
  'Email Sender Provider',
  'Email Sender Endpoint'
];

/**
 * Master 28 Canonical Header Columns for Campaigns Tab (LOCKED SCHEMA)
 * Column 4 (Index 3) = Status
 */
const CAMPAIGNS_MASTER_HEADERS = [
  'Campaign ID',
  'Campaign Name',
  'Department',
  'Status',
  'Target Countries',
  'Target Industries',
  'Company Size',
  'Discovery Limit',
  'Research Depth',
  'Evidence Period Months',
  'Min Company Score',
  'Min Persona Score',
  'Min Email Priority Score',
  'Target Titles',
  'Campaign Focus',
  'Search Keywords',
  'Custom Business Needs',
  'AI Provider',
  'AI Model',
  'Email Verifier Provider',
  'Email Verifier Endpoint',
  'Email Sender Provider',
  'Email Sender Endpoint',
  'Sender Name',
  'Sender Title',
  'Sender Company',
  'Campaign Config JSON',
  'Last Modified'
];

/**
 * Helper to clean and normalize keys (removes spaces, underscores, symbols, uppercase)
 */
function normalizeStudioKey(str) {
  return String(str || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Helper to verify a string is a real secret/key and not a sheet status placeholder.
 */
function isStudioValidSecretValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const upper = s.toUpperCase();
  if (upper.indexOf('PASTE_') >= 0) return false;
  if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
  if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
  return true;
}

/**
 * Identifies whether a Client / Owner name refers to the local/agency brand ("Growth Target").
 * Local brand runs use default Script Properties; any other client enforces strict CLIENT_{NAME}_ isolation.
 */
function isLocalAgencyClient_(clientName) {
  const clean = String(clientName || '').trim();
  if (!clean) return true;
  const norm = normalizeStudioKey(clean);
  return (
    norm === 'GROWTHTARGET' ||
    norm === 'LOCAL' ||
    norm === 'AGENCY' ||
    norm === 'DEFAULT'
  );
}

/**
 * Builds canonical Script Property prefix for an external client: e.g. "Pilot Acme" -> "CLIENT_PILOT_ACME_".
 */
function normalizeClientPropertyPrefix_(clientName) {
  const clean = String(clientName || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return clean ? ('CLIENT_' + clean + '_') : '';
}

/**
 * Resolves the exact Client / Owner name associated with a Campaign ID.
 * Priority: explicit clientName -> Campaigns tab (Col C) -> Campaign Configuration (if ID matches) -> Growth Target.
 */
function resolveCampaignClientName_(campaignId, explicitClientName, spreadsheet) {
  if (explicitClientName && String(explicitClientName).trim()) {
    return String(explicitClientName).trim();
  }

  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const cid = String(campaignId || '').trim();

  if (ss && cid) {
    const campSheet = ss.getSheetByName('Campaigns');
    if (campSheet && campSheet.getLastRow() > 1) {
      const cData = campSheet.getDataRange().getValues();
      const headers = cData[0].map(function(h) { return normalizeStudioKey(h); });
      const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
      const deptIdx = headers.indexOf('DEPARTMENT') >= 0 ? headers.indexOf('DEPARTMENT') : 2;
      const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON');

      for (let r = 1; r < cData.length; r++) {
        if (String(cData[r][idIdx] || '').trim().toLowerCase() === cid.toLowerCase()) {
          if (cData[r][deptIdx] && String(cData[r][deptIdx]).trim()) {
            return String(cData[r][deptIdx]).trim();
          }
          if (jsonIdx >= 0 && cData[r][jsonIdx]) {
            try {
              const parsed = JSON.parse(cData[r][jsonIdx]);
              if (parsed && parsed.department) return String(parsed.department).trim();
            } catch (e) {}
          }
        }
      }
    }

    const cfgSheet = ss.getSheetByName('Campaign Configuration');
    if (cfgSheet && cfgSheet.getLastRow() > 1) {
      const cfgData = cfgSheet.getDataRange().getValues();
      const dict = {};
      for (let r = 1; r < cfgData.length; r++) {
        const k = normalizeStudioKey(cfgData[r][0]);
        if (k) dict[k] = cfgData[r][1];
      }
      const cfgId = String(dict['CAMPAIGNID'] || '').trim();
      if (cfgId && cfgId.toLowerCase() === cid.toLowerCase()) {
        const cfgDept = String(dict['DEPARTMENT'] || dict['CLIENTOWNER'] || dict['CLIENT'] || '').trim();
        if (cfgDept) return cfgDept;
      }
    }
  }

  if (cid && cid.toUpperCase().indexOf('PILOT') === 0) {
    return 'Pilot Acme';
  }

  return 'Growth Target';
}

/**
 * Universal Strict Credential Resolver (ZERO fallback for external client campaigns).
 *
 * @param {Array<string>|string} basePropertyNames Candidate Script Property base names (e.g. ['GEMINI_API_KEY'])
 * @param {Object|string} context Campaign/client context ({ campaignId, clientName }) or campaignId string
 * @param {Object} [options] Optional controls: { serviceLabel, throwOnMissing, allowAgentSettingsForLocal }
 * @return {Object} { key, clientName, isLocal, propertyUsed, expectedProperty }
 */
function resolveClientScopedSecret_(basePropertyNames, context, options) {
  const bases = Array.isArray(basePropertyNames) ? basePropertyNames : [basePropertyNames];
  const opts = options || {};
  const ctx = (context && typeof context === 'object')
    ? context
    : { campaignId: context ? String(context) : '' };

  const ss = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const resolvedCampaignId = ctx.campaignId || getStudioActiveCampaignId_();
  const clientName = resolveCampaignClientName_(resolvedCampaignId, ctx.clientName, ss);
  const isLocal = isLocalAgencyClient_(clientName);

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();

  if (!isLocal) {
    // STRICT EXTERNAL CLIENT MODE: Only CLIENT_{CLIENT_NAME}_{BASE_KEY} is allowed. Zero local fallback!
    const prefix = normalizeClientPropertyPrefix_(clientName);
    const expectedProperty = prefix + bases[0];
    const normTargets = bases.map(function(b) {
      return normalizeStudioKey(prefix + b);
    });

    // 1. Check Script Properties (preferred encrypted store)
    for (let k in scriptProps) {
      if (normTargets.indexOf(normalizeStudioKey(k)) >= 0 && isStudioValidSecretValue_(scriptProps[k])) {
        return {
          key: String(scriptProps[k]).trim(),
          clientName: clientName,
          isLocal: false,
          propertyUsed: k,
          expectedProperty: expectedProperty
        };
      }
    }

    // 2. Check User Properties (client-prefixed only)
    for (let k in userProps) {
      if (normTargets.indexOf(normalizeStudioKey(k)) >= 0 && isStudioValidSecretValue_(userProps[k])) {
        return {
          key: String(userProps[k]).trim(),
          clientName: clientName,
          isLocal: false,
          propertyUsed: k,
          expectedProperty: expectedProperty
        };
      }
    }

    if (opts.throwOnMissing) {
      const label = opts.serviceLabel || bases[0];
      throw new Error(
        'Missing client credential for "' + clientName + '" (' + label + '). ' +
        'Set Script Property "' + expectedProperty + '" in Project Settings > Script Properties. ' +
        '(Strict client isolation is active — local keys are never used as fallback.)'
      );
    }

    return {
      key: '',
      clientName: clientName,
      isLocal: false,
      propertyUsed: '',
      expectedProperty: expectedProperty
    };
  }

  // LOCAL AGENCY MODE ("Growth Target"): Check default keys, explicitly ignoring any CLIENT_* keys
  const normBases = bases.map(function(b) { return normalizeStudioKey(b); });

  for (let k in scriptProps) {
    if (String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
    if (normBases.indexOf(normalizeStudioKey(k)) >= 0 && isStudioValidSecretValue_(scriptProps[k])) {
      return {
        key: String(scriptProps[k]).trim(),
        clientName: clientName,
        isLocal: true,
        propertyUsed: k,
        expectedProperty: bases[0]
      };
    }
  }

  for (let k in userProps) {
    if (String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
    if (normBases.indexOf(normalizeStudioKey(k)) >= 0 && isStudioValidSecretValue_(userProps[k])) {
      return {
        key: String(userProps[k]).trim(),
        clientName: clientName,
        isLocal: true,
        propertyUsed: k,
        expectedProperty: bases[0]
      };
    }
  }

  if (opts.allowAgentSettingsForLocal !== false && ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const data = agentSheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        const settingKey = normalizeStudioKey(data[r][1] || data[r][0]);
        if (normBases.indexOf(settingKey) >= 0) {
          const val = String(data[r][2] !== undefined ? data[r][2] : '').trim();
          if (isStudioValidSecretValue_(val)) {
            return {
              key: val,
              clientName: clientName,
              isLocal: true,
              propertyUsed: 'Agent Settings:' + String(data[r][1] || data[r][0]),
              expectedProperty: bases[0]
            };
          }
        }
      }
    }
  }

  return {
    key: '',
    clientName: clientName,
    isLocal: true,
    propertyUsed: '',
    expectedProperty: bases[0]
  };
}

/**
 * Returns active registered clients from Tab 15 ("Lists & Settings").
 */
function getRegisteredClientsList_(ss) {
  if (typeof getRegisteredClientsFromSettings_ === 'function') {
    return getRegisteredClientsFromSettings_(ss);
  }
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const clients = [];
  const seen = {};
  let hasAnyClientRow = false;

  if (spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Lists & Settings');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const listName = String(data[i][0] || '').trim().toLowerCase();
        const val = String(data[i][1] || '').trim();
        const status = String(data[i][2] || '').trim().toLowerCase();
        const isActive = (status === 'active' || status === 'true' || status === '');
        if (val && (listName === 'client - account' || listName === 'client / owner' || listName === 'client')) {
          hasAnyClientRow = true;
          if (isActive && !seen[val.toLowerCase()]) {
            clients.push(val);
            seen[val.toLowerCase()] = true;
          }
        }
      }
    }
  }
  if (clients.length > 0) return clients;
  return hasAnyClientRow ? ['Growth Target'] : ['Growth Target', 'Pilot Acme'];
}

/**
 * Returns Client - Account names explicitly marked Inactive in Tab 15 ("Lists & Settings").
 */
function getInactiveClientsList_(ss) {
  if (typeof getInactiveClientsFromSettings_ === 'function') {
    return getInactiveClientsFromSettings_(ss);
  }
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const inactive = [];
  const activeSet = {};
  const seenInactive = {};

  if (spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Lists & Settings');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const listName = String(data[i][0] || '').trim().toLowerCase();
        const val = String(data[i][1] || '').trim();
        const status = String(data[i][2] || '').trim().toLowerCase();
        if (!val || !(listName === 'client - account' || listName === 'client / owner' || listName === 'client')) {
          continue;
        }
        const isActive = (status === 'active' || status === 'true' || status === '');
        if (isActive) {
          activeSet[val.toLowerCase()] = true;
        } else if (!seenInactive[val.toLowerCase()]) {
          inactive.push(val);
          seenInactive[val.toLowerCase()] = true;
        }
      }
    }
  }
  return inactive.filter(function(name) {
    return !activeSet[name.toLowerCase()];
  });
}

/**
 * Filters the "Campaigns" tab (Column C = Department / Client Owner) using native
 * Google Sheets FilterCriteria so only campaigns belonging to clientName are shown.
 * Preserves the locked 28-column schema and leaves blank template rows untouched.
 */
function filterCampaignsSheetByClient_(ss, clientName) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (!campSheet) return;

  const expectedLastRow = campSheet.getLastRow();
  const expectedLastCol = campSheet.getLastColumn();
  if (expectedLastRow <= 1 || expectedLastCol < 1) {
    const staleFilter = campSheet.getFilter();
    if (staleFilter) staleFilter.remove();
    return;
  }

  const headers = campSheet.getRange(1, 1, 1, expectedLastCol).getValues()[0];
  let deptColIndex = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = normalizeStudioKey(headers[c]);
    if (h === 'DEPARTMENT' || h === 'CLIENTOWNER' || h === 'CLIENTACCOUNT' || h === 'CLIENT') {
      deptColIndex = c + 1;
      break;
    }
  }
  if (deptColIndex === -1 && expectedLastCol >= 3) {
    deptColIndex = 3;
  }
  if (deptColIndex === -1) return;

  const cleanClient = String(clientName || '').trim() || 'Growth Target';
  const targetLower = cleanClient.toLowerCase();
  const numDataRows = expectedLastRow - 1;

  let filter = campSheet.getFilter();
  if (filter) {
    const fRange = filter.getRange();
    if (fRange.getRow() !== 1 || fRange.getLastRow() !== expectedLastRow || fRange.getLastColumn() < expectedLastCol) {
      filter.remove();
      filter = null;
    }
  }

  if (!filter) {
    filter = campSheet.getRange(1, 1, expectedLastRow, expectedLastCol).createFilter();
  }

  const colValues = campSheet.getRange(2, deptColIndex, numDataRows, 1).getValues();
  const hiddenValues = [];
  const seen = {};

  for (let i = 0; i < colValues.length; i++) {
    const rawVal = String(colValues[i][0] === null || colValues[i][0] === undefined ? '' : colValues[i][0]);
    const trimmedLower = rawVal.trim().toLowerCase();
    if (trimmedLower && trimmedLower !== targetLower && !seen[rawVal]) {
      hiddenValues.push(rawVal);
      seen[rawVal] = true;
    }
  }

  if (hiddenValues.length > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(hiddenValues)
      .build();
    filter.setColumnFilterCriteria(deptColIndex, criteria);
  } else {
    filter.removeColumnFilterCriteria(deptColIndex);
  }
}

/**
 * Filters the "Dashboard" tab's Campaign Performance Rollup table (Row 17+, Col C = Department)
 * using native Google Sheets FilterCriteria so only campaigns belonging to clientName are shown.
 * Leaves Dashboard KPI rows (Rows 1-16) completely untouched.
 */
function filterDashboardRollupByClient_(ss, clientName) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const dashSheet = spreadsheet.getSheetByName('Dashboard');
  if (!dashSheet) return;

  // Locate the Campaign Performance Rollup header row (default Row 17)
  let headerRow = -1;
  let deptColIndex = 3;
  const scanMax = Math.min(dashSheet.getLastRow(), 25);
  if (scanMax >= 15) {
    const scanCols = Math.min(Math.max(dashSheet.getLastColumn(), 7), 10);
    const scanVals = dashSheet.getRange(15, 1, scanMax - 14, scanCols).getValues();
    for (let r = 0; r < scanVals.length; r++) {
      const colA = normalizeStudioKey(scanVals[r][0]);
      if (colA === 'CAMPAIGNID') {
        headerRow = 15 + r;
        for (let c = 0; c < scanVals[r].length; c++) {
          const h = normalizeStudioKey(scanVals[r][c]);
          if (h === 'DEPARTMENT' || h === 'CLIENTOWNER' || h === 'CLIENTACCOUNT' || h === 'CLIENT') {
            deptColIndex = c + 1;
            break;
          }
        }
        break;
      }
    }
  }

  if (headerRow === -1) {
    const staleFilter = dashSheet.getFilter();
    if (staleFilter) staleFilter.remove();
    return;
  }

  const rollupCols = 7;
  const expectedLastCol = Math.max(dashSheet.getLastColumn(), rollupCols);
  const rawLastRow = dashSheet.getLastRow();
  let dataLastRow = headerRow;

  if (rawLastRow > headerRow) {
    const bodyVals = dashSheet.getRange(headerRow + 1, 1, rawLastRow - headerRow, expectedLastCol).getValues();
    for (let i = bodyVals.length - 1; i >= 0; i--) {
      const hasContent = bodyVals[i].some(function(cell) {
        return cell !== '' && cell !== null && cell !== undefined;
      });
      if (hasContent) {
        dataLastRow = headerRow + 1 + i;
        break;
      }
    }
  }

  if (dataLastRow <= headerRow) {
    const staleFilter = dashSheet.getFilter();
    if (staleFilter) staleFilter.remove();
    return;
  }

  const cleanClient = String(clientName || '').trim() || 'Growth Target';
  const targetLower = cleanClient.toLowerCase();
  const numDataRows = dataLastRow - headerRow;

  let filter = dashSheet.getFilter();
  if (filter) {
    const fRange = filter.getRange();
    if (fRange.getRow() !== headerRow || fRange.getLastRow() !== dataLastRow || fRange.getLastColumn() < rollupCols) {
      filter.remove();
      filter = null;
    }
  }

  if (!filter) {
    filter = dashSheet.getRange(headerRow, 1, numDataRows + 1, rollupCols).createFilter();
  }

  const colValues = dashSheet.getRange(headerRow + 1, deptColIndex, numDataRows, 1).getValues();
  const hiddenValues = [];
  const seen = {};

  for (let i = 0; i < colValues.length; i++) {
    const rawVal = String(colValues[i][0] === null || colValues[i][0] === undefined ? '' : colValues[i][0]);
    const trimmedLower = rawVal.trim().toLowerCase();
    if (trimmedLower && trimmedLower !== targetLower && !seen[rawVal]) {
      hiddenValues.push(rawVal);
      seen[rawVal] = true;
    }
  }

  if (hiddenValues.length > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(hiddenValues)
      .build();
    filter.setColumnFilterCriteria(deptColIndex, criteria);
  } else {
    filter.removeColumnFilterCriteria(deptColIndex);
  }
}

/**
 * Adds or reactivates a Client / Account in "Lists & Settings", pins Client - Account rows
 * at the top, refreshes sheet dropdowns, provisions "Agent Settings", and filters tabs
 * (including Email Deliverability B3/D3).
 * Callable directly from the Sidebar Control Panel "[ + Add New Client... ]".
 */
function addClientAccountToSettings(clientName) {
  const cleanName = String(clientName || '').trim();
  if (!cleanName) {
    throw new Error('Client name cannot be empty.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) throw new Error('No active spreadsheet found.');

  let sheet = ss.getSheetByName('Lists & Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Lists & Settings');
    sheet.appendRow(['List name', 'Value', 'Active status']);
  }

  let foundRow = -1;
  if (sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const listType = String(data[i][0] || '').trim().toLowerCase();
      const val = String(data[i][1] || '').trim();
      if ((listType === 'client - account' || listType === 'client / owner' || listType === 'client') &&
          val.toLowerCase() === cleanName.toLowerCase()) {
        foundRow = i + 1;
        break;
      }
    }
  }

  if (foundRow > 1) {
    sheet.getRange(foundRow, 2).setValue(cleanName);
    sheet.getRange(foundRow, 3).setValue('Active');
  } else {
    sheet.appendRow(['Client - Account', cleanName, 'Active']);
  }

  if (typeof sortListsAndSettingsTab_ === 'function') {
    sortListsAndSettingsTab_(ss);
  }
  if (typeof refreshClientDropdownValidations_ === 'function') {
    refreshClientDropdownValidations_(ss);
  }
  if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
    refreshDeliverabilityFilterDropdowns_(ss);
  }
  if (typeof setupAgentSettingsTabDropdowns === 'function') {
    setupAgentSettingsTabDropdowns(ss, cleanName);
  }
  if (typeof filterAgentSettingsByClient_ === 'function') {
    filterAgentSettingsByClient_(ss, cleanName);
  }
  filterCampaignsSheetByClient_(ss, cleanName);
  filterDashboardRollupByClient_(ss, cleanName);
  if (typeof syncDeliverabilityTabFilters_ === 'function') {
    syncDeliverabilityTabFilters_(ss, cleanName, 'All Campaigns');
  }
  SpreadsheetApp.flush();

  return {
    success: true,
    clientName: cleanName,
    clients: getRegisteredClientsList_(ss),
    inactiveClients: getInactiveClientsList_(ss),
    message: 'Added client "' + cleanName + '" to Lists & Settings, updated dropdowns, and created Agent Settings rows.'
  };
}

/**
 * Refreshes Agent Settings protection status indicators and filters "Agent Settings",
 * "Campaigns", "Dashboard" Rollup, and "Email Deliverability" (B3/D3) to the specified
 * Client / Owner — including 0-campaign clients.
 */
function filterAgentSettingsForClientWorkspace(clientName) {
  const cleanClient = String(clientName || '').trim() || 'Growth Target';
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    return { success: false, message: 'No active spreadsheet found.' };
  }

  if (typeof setupAgentSettingsTabDropdowns === 'function') {
    setupAgentSettingsTabDropdowns(ss, cleanClient);
  }
  if (typeof filterAgentSettingsByClient_ === 'function') {
    filterAgentSettingsByClient_(ss, cleanClient);
  }
  filterCampaignsSheetByClient_(ss, cleanClient);
  filterDashboardRollupByClient_(ss, cleanClient);
  if (typeof syncDeliverabilityTabFilters_ === 'function') {
    syncDeliverabilityTabFilters_(ss, cleanClient, 'All Campaigns');
  }
  SpreadsheetApp.flush();

  return {
    success: true,
    clientName: cleanClient,
    message: 'Dashboard Rollup, Campaigns, Email Deliverability, and Agent Settings filtered to client "' + cleanClient + '".'
  };
}

/**
 * Safely resolves reviewer email without throwing if userinfo.email scope is restricted.
 */
function getSafeReviewerEmail_() {
  try {
    const active = Session.getActiveUser().getEmail();
    if (active) return active;
  } catch (e) {}
  try {
    const effective = Session.getEffectiveUser().getEmail();
    if (effective) return effective;
  } catch (e) {}
  return 'Admin Reviewer';
}

/**
 * Safely parses numeric fields read from Sheets while ignoring accidental Date cells (e.g. 3/20/1900).
 */
function parseSafeCampaignNumeric_(val) {
  if (val === null || val === undefined || val === '' || val instanceof Date) return 0;
  const n = Number(val);
  return isFinite(n) && n > 0 ? n : 0;
}

/**
 * Master onOpen Trigger: Clean menu with isolated audit & deliverability controls.
 * Note: Outreach functions live in EmailSender.js to avoid global namespace collisions.
 */
function onOpen() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    return;
  }

  ui.createMenu('🎯 AI Prospecting Studio')
    .addItem('🚀 Open Campaign Control Panel', 'showPersonaMapperControlPanel')
    .addItem('🔄 Filter Tabs to Active Campaign', 'filterSheetsForActiveCampaignMenu')
    .addSeparator()
    .addItem('🔍 1. Discover Target Companies', 'runCompanyDiscoveryForSelectedCampaign')
    .addItem('👤 2. Run Persona Mapper', 'runPersonaMapperForSelectedCampaign')
    .addSeparator()
    .addItem('📋 Go to Prospect Review Queue', 'navigateToReviewQueueSheet')
    .addItem('✅ Approve Selected Prospect', 'approveSelectedProspect')
    .addItem('✏️ Mark Selected: Needs Editing', 'markSelectedProspectNeedsEditing')
    .addItem('❌ Reject Selected Prospect', 'rejectSelectedProspect')
    .addSeparator()
    .addItem('📝 Create Outreach Drafts (Safe Review)', 'createGmailDraftsForApprovedProspects')
    .addItem('✉️ Send Approved Outreach Emails', 'sendApprovedOutreachEmails')
    .addSeparator()
    .addItem('📊 Open Email Deliverability', 'openEmailDeliverabilitySheet')
    .addItem('🔄 Sync Email Deliverability Events', 'syncEmailDeliverabilityEvents')
    .addItem('⏱️ Enable Hourly Deliverability Sync', 'installDeliverabilitySyncTrigger')
    .addSeparator()
    .addItem('❓ Check Spreadsheet Setup', 'runStudioSpreadsheetAudit')
    .addSeparator()
    .addItem('🧪 Run Authentic Pilot Pipeline', 'triggerPilotRun')
    .addItem('🧹 Clean Up Pilot Test Data', 'triggerPilotCleanup')
    .addToUi();
}

/**
 * Native onEdit Trigger:
 * 1. Syncs Dashboard (D1/C4 Status, B2 Campaign ID) and Campaigns Column 4 (Status).
 * 2. Keeps "Lists & Settings" sorted with Client - Account rows pinned at the top
 *    and refreshes Client dropdown validations whenever a completed row is edited.
 * 3. Automatically resets "Email Deliverability" Campaign Filter (D3) to "All Campaigns"
 *    if Client Filter (B3) is switched away from the selected campaign's owner.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const val = String(e.value || e.range.getValue() || '').trim();

    if (sheetName === 'Dashboard') {
      if (row === 1 && col === 4 && val) {
        syncStatusFromDashboard_(val);
      } else if (row === 4 && col === 3 && val) {
        syncStatusFromDashboard_(val);
      } else if (row === 2 && col === 2 && val) {
        selectAndActivateCampaign(val);
      }
    } else if (sheetName === 'Campaigns' && row > 1 && col === 4) {
      const editedCampId = String(sheet.getRange(row, 1).getValue()).trim();
      const activeCampId = getStudioActiveCampaignId_();
      if (editedCampId && editedCampId === activeCampId) {
        updateDashboardStatusOnly_(activeCampId, val);
      }
    } else if (sheetName === 'Email Deliverability' && row === 3 && col === 2) {
      const newClient = String(val || '').trim();
      const campFilterCell = sheet.getRange(3, 4);
      const currentCamp = String(campFilterCell.getValue() || '').trim();
      if (newClient && newClient !== 'All Clients' && currentCamp && currentCamp !== 'All Campaigns') {
        const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
        const campOwner = resolveCampaignClientName_(currentCamp, '', ss);
        if (campOwner && campOwner.toLowerCase() !== newClient.toLowerCase()) {
          campFilterCell.setValue('All Campaigns');
        }
      }
    } else if (sheetName === 'Lists & Settings' && row > 1 && col >= 1 && col <= 3) {
      const rowVals = sheet.getRange(row, 1, 1, 3).getValues()[0];
      const listType = String(rowVals[0] || '').trim();
      const listVal = String(rowVals[1] || '').trim();
      const lowerType = listType.toLowerCase();

      // Wait until both List name and Value are filled so an in-progress new row doesn't jump mid-entry
      if (listType && listVal) {
        const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
        if (typeof sortListsAndSettingsTab_ === 'function') {
          sortListsAndSettingsTab_(ss);
        }
        if (lowerType === 'client - account' || lowerType === 'client / owner' || lowerType === 'client') {
          if (typeof refreshClientDropdownValidations_ === 'function') {
            refreshClientDropdownValidations_(ss);
          }
          if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
            refreshDeliverabilityFilterDropdowns_(ss);
          }
          if (ss && ss.toast) {
            ss.toast('Lists & Settings organized & Client dropdowns updated.', '✅ Registry Updated', 3);
          }
        }
      }
    }
  } catch (err) {
    Logger.log('[onEdit Error] ' + err.message);
  }
}

/**
 * Synchronizes Status changes from Dashboard (D1/C4) into Campaigns Column 4
 * and updates the Performance Rollup row without shifting columns.
 */
function syncStatusFromDashboard_(newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const activeCampId = getStudioActiveCampaignId_();
  if (!activeCampId) return;

  const campSheet = ss.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const cData = campSheet.getDataRange().getValues();
    for (let r = 1; r < cData.length; r++) {
      if (String(cData[r][0]).trim() === activeCampId) {
        campSheet.getRange(r + 1, 4).setValue(newStatus);
        break;
      }
    }
  }

  updateDashboardStatusOnly_(activeCampId, newStatus);
}

/**
 * Updates status indicators on Dashboard without touching other fields.
 */
function updateDashboardStatusOnly_(cid, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) return;

  dashSheet.getRange('D1').setValue(newStatus);
  dashSheet.getRange('C4').setValue(newStatus);

  const lastRow = dashSheet.getLastRow();
  if (lastRow >= 18) {
    for (let r = 18; r <= lastRow; r++) {
      const rowCampId = String(dashSheet.getRange(r, 1).getValue()).trim();
      if (rowCampId === cid) {
        dashSheet.getRange(r, 4).setValue(newStatus);
        dashSheet.getRange(r, 7).setValue(new Date().toLocaleString());
        break;
      }
    }
  }
  SpreadsheetApp.flush();
}

/**
 * Collision-proof Active Campaign ID getter.
 * Reads directly from Dashboard Cell B2 or Campaign Configuration.
 */
function getStudioActiveCampaignId_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return 'CMP-005';
  const dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet) {
    const b2 = String(dashSheet.getRange('B2').getValue() || '').trim();
    if (b2 && b2 !== 'Select Campaign' && b2 !== 'No Active Campaign') return b2;
  }
  const configSheet = ss.getSheetByName('Campaign Configuration');
  if (configSheet && configSheet.getLastRow() > 1) {
    const cId = String(configSheet.getRange(2, 2).getValue() || '').trim();
    if (cId) return cId;
  }
  return 'CMP-005';
}

/**
 * Restyles headers across operational tabs (kept as internal utility).
 */
function fixAllTabHeaderContrast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const sheets = ss.getSheets();

  sheets.forEach(function(sh) {
    const sName = sh.getName();
    if (sName === 'Dashboard' || sName === 'Email Deliverability') return;
    if (sh.getLastColumn() > 0) {
      const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
      headerRange
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sh.setFrozenRows(1);
    }
  });

  SpreadsheetApp.flush();
}

/**
 * Adaptive sheet finder prioritizing "Persona Mapper Settings" (plural)
 */
function findPersonaMapperSettingsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return null;

  const candidateNames = [
    'Persona Mapper Settings',
    'Persona Mapper Setting',
    'Persona  Mapper Settings',
    'Persona  Mapper Setting',
    'PersonaMapperSettings',
    'Persona_Mapper_Settings'
  ];

  for (let i = 0; i < candidateNames.length; i++) {
    const s = ss.getSheetByName(candidateNames[i]);
    if (s) return s;
  }

  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const norm = allSheets[i].getName().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (norm.includes('personamapper') || (norm.includes('persona') && norm.includes('setting'))) {
      return allSheets[i];
    }
  }

  return ss.insertSheet('Persona Mapper Settings');
}

/**
 * Client-scoped API Key resolver.
 * Delegates to resolveClientScopedSecret_ so external clients strictly require
 * CLIENT_{NAME}_{KEY} and Growth Target safely uses default keys (never CLIENT_*).
 */
function getStudioApiKey(provider, context) {
  const norm = String(provider || '').toUpperCase();
  const baseKey = (norm.indexOf('OPENAI') >= 0 || norm.indexOf('GPT') >= 0)
    ? 'OPENAI_API_KEY'
    : 'GEMINI_API_KEY';

  const res = resolveClientScopedSecret_([baseKey], context || {}, {
    throwOnMissing: false,
    allowAgentSettingsForLocal: true
  });
  return res.key || '';
}

function getApiKey(provider, context) {
  return getStudioApiKey(provider, context);
}

/**
 * Opens the Campaign Control Panel sidebar and pre-loads campaigns, clients, and configuration.
 */
function showPersonaMapperControlPanel() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  const sidebarData = getSidebarInitialData();
  template.initialCampaigns = sidebarData.campaigns;
  template.initialConfigsMap = sidebarData.configsMap;
  template.activeCampaignId = sidebarData.activeCampaignId || '';
  template.initialClients = sidebarData.clients || ['Growth Target'];
  template.initialInactiveClients = sidebarData.inactiveClients || [];

  const html = template.evaluate().setTitle('🎯 Campaign Control Panel').setWidth(380);
  try {
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    Logger.log('showPersonaMapperControlPanel called outside UI context: ' + e.message);
  }
}

/**
 * Returns available campaigns, registered active/inactive clients, and configuration maps.
 */
function getSidebarInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const campaignsList = [];
  const campaignsMap = {};
  const configsMap = {};
  let activeCampaignId = '';

  const cSheet = ss ? ss.getSheetByName('Campaign Configuration') : null;
  const activeConfigDict = {};
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let j = 1; j < cData.length; j++) {
      const k = normalizeStudioKey(cData[j][0]);
      if (k) activeConfigDict[k] = cData[j][1];
    }
    activeCampaignId = String(activeConfigDict['CAMPAIGNID'] || activeConfigDict['ACTIVECAMPAIGNID'] || activeConfigDict['CAMPAIGN'] || '').trim();
  }

  if (!activeCampaignId) {
    activeCampaignId = getStudioActiveCampaignId_();
  }

  const campSheet = ss ? ss.getSheetByName('Campaigns') : null;
  if (campSheet && campSheet.getLastRow() > 1) {
    const campData = campSheet.getDataRange().getValues();
    const headers = campData[0].map(function(h) { return normalizeStudioKey(h); });

    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const nameIdx = headers.indexOf('CAMPAIGNNAME') >= 0 ? headers.indexOf('CAMPAIGNNAME') : 1;
    const deptIdx = headers.indexOf('DEPARTMENT');
    const statusIdx = headers.indexOf('STATUS') >= 0 ? headers.indexOf('STATUS') : 3;
    const countryIdx = headers.indexOf('TARGETCOUNTRIES') >= 0 ? headers.indexOf('TARGETCOUNTRIES') : headers.indexOf('COUNTRIES');
    const indIdx = headers.indexOf('TARGETINDUSTRIES') >= 0 ? headers.indexOf('TARGETINDUSTRIES') : headers.indexOf('INDUSTRIES');
    const sizeIdx = headers.indexOf('COMPANYSIZE') >= 0 ? headers.indexOf('COMPANYSIZE') : headers.indexOf('COMPANYSIZERANGE');
    const limitIdx = headers.indexOf('DISCOVERYLIMIT') >= 0 ? headers.indexOf('DISCOVERYLIMIT') : headers.indexOf('MAXCOMPANIES');
    const depthIdx = headers.indexOf('RESEARCHDEPTH');
    const evIdx = headers.indexOf('EVIDENCEPERIODMONTHS');
    const minCompIdx = headers.indexOf('MINCOMPANYSCORE');
    const minPersIdx = headers.indexOf('MINPERSONASCORE');
    const minEmailIdx = headers.indexOf('MINEMAILPRIORITYSCORE');
    const titlesIdx = headers.indexOf('TARGETTITLES') >= 0 ? headers.indexOf('TARGETTITLES') : headers.indexOf('TITLES');
    const focusIdx = headers.indexOf('CAMPAIGNFOCUS');
    const keywordsIdx = headers.indexOf('SEARCHKEYWORDS');
    const needsIdx = headers.indexOf('CUSTOMBUSINESSNEEDS');
    const aiProvIdx = headers.indexOf('AIPROVIDER');
    const aiModIdx = headers.indexOf('AIMODEL');
    const verifProvIdx = headers.indexOf('EMAILVERIFIERPROVIDER') >= 0 ? headers.indexOf('EMAILVERIFIERPROVIDER') : headers.indexOf('EMAILVERIFIER');
    const verifEndIdx = headers.indexOf('EMAILVERIFIERENDPOINT');
    const senderProvIdx = headers.indexOf('EMAILSENDERPROVIDER');
    const senderEndIdx = headers.indexOf('EMAILSENDERENDPOINT');
    const senderNameIdx = headers.indexOf('SENDERNAME');
    const senderTitleIdx = headers.indexOf('SENDERTITLE');
    const senderCoIdx = headers.indexOf('SENDERCOMPANY');
    const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON');

    for (let i = 1; i < campData.length; i++) {
      const cid = String(campData[i][idIdx] || '').trim();
      if (!cid) continue;

      const name = String(campData[i][nameIdx] || cid).trim();
      campaignsMap[cid] = name;

      let campaignObj = { id: cid, name: name, status: 'Active' };

      if (jsonIdx >= 0 && campData[i][jsonIdx]) {
        try {
          const parsed = JSON.parse(campData[i][jsonIdx]);
          if (parsed && typeof parsed === 'object') {
            campaignObj = Object.assign(campaignObj, parsed);
            if (parsed.senderEmail) campaignObj.senderEmail = String(parsed.senderEmail).trim();
            if (parsed.replyToEmail) campaignObj.replyToEmail = String(parsed.replyToEmail).trim();
          }
        } catch (e) {}
      }

      if (deptIdx >= 0 && campData[i][deptIdx]) campaignObj.department = String(campData[i][deptIdx]).trim();
      if (statusIdx >= 0 && campData[i][statusIdx]) campaignObj.status = String(campData[i][statusIdx]).trim();
      if (countryIdx >= 0 && campData[i][countryIdx]) campaignObj.country = String(campData[i][countryIdx]).trim();
      if (indIdx >= 0 && campData[i][indIdx]) campaignObj.industry = String(campData[i][indIdx]).trim();
      if (sizeIdx >= 0 && campData[i][sizeIdx]) campaignObj.size = String(campData[i][sizeIdx]).trim();

      const parsedLimit = limitIdx >= 0 ? parseSafeCampaignNumeric_(campData[i][limitIdx]) : 0;
      if (parsedLimit > 0) campaignObj.limit = parsedLimit;

      if (depthIdx >= 0 && campData[i][depthIdx]) campaignObj.researchDepth = String(campData[i][depthIdx]).trim();

      const parsedEv = evIdx >= 0 ? parseSafeCampaignNumeric_(campData[i][evIdx]) : 0;
      if (parsedEv > 0) campaignObj.evidencePeriodMonths = parsedEv;

      const parsedMinComp = minCompIdx >= 0 ? parseSafeCampaignNumeric_(campData[i][minCompIdx]) : 0;
      if (parsedMinComp > 0) campaignObj.minCompanyScore = parsedMinComp;

      const parsedMinPers = minPersIdx >= 0 ? parseSafeCampaignNumeric_(campData[i][minPersIdx]) : 0;
      if (parsedMinPers > 0) campaignObj.minPersonaScore = parsedMinPers;

      const parsedMinEmail = minEmailIdx >= 0 ? parseSafeCampaignNumeric_(campData[i][minEmailIdx]) : 0;
      if (parsedMinEmail > 0) campaignObj.minEmailPriorityScore = parsedMinEmail;

      if (titlesIdx >= 0 && campData[i][titlesIdx]) campaignObj.titles = String(campData[i][titlesIdx]).trim();
      if (focusIdx >= 0 && campData[i][focusIdx]) campaignObj.campaignFocus = String(campData[i][focusIdx]).trim();
      if (keywordsIdx >= 0 && campData[i][keywordsIdx]) campaignObj.searchKeywords = String(campData[i][keywordsIdx]).trim();
      if (needsIdx >= 0 && campData[i][needsIdx]) campaignObj.customBusinessNeeds = String(campData[i][needsIdx]).trim();
      if (aiProvIdx >= 0 && campData[i][aiProvIdx]) campaignObj.aiProvider = String(campData[i][aiProvIdx]).trim();
      if (aiModIdx >= 0 && campData[i][aiModIdx]) campaignObj.aiModel = String(campData[i][aiModIdx]).trim();
      if (verifProvIdx >= 0 && campData[i][verifProvIdx]) campaignObj.emailVerifier = String(campData[i][verifProvIdx]).trim();
      if (verifEndIdx >= 0 && campData[i][verifEndIdx] !== '') campaignObj.emailEndpoint = String(campData[i][verifEndIdx]).trim();
      if (senderProvIdx >= 0 && campData[i][senderProvIdx]) campaignObj.emailSenderProvider = String(campData[i][senderProvIdx]).trim();
      if (senderEndIdx >= 0 && campData[i][senderEndIdx] !== '') campaignObj.emailSenderEndpoint = String(campData[i][senderEndIdx]).trim();
      if (senderNameIdx >= 0 && campData[i][senderNameIdx]) campaignObj.senderName = String(campData[i][senderNameIdx]).trim();
      if (senderTitleIdx >= 0 && campData[i][senderTitleIdx]) campaignObj.senderTitle = String(campData[i][senderTitleIdx]).trim();
      if (senderCoIdx >= 0 && campData[i][senderCoIdx]) campaignObj.senderCompany = String(campData[i][senderCoIdx]).trim();

      configsMap[cid] = campaignObj;
    }
  }

  const pSheet = findPersonaMapperSettingsSheet(ss);
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const cid = String(pData[i][0] || '').trim();
      if (!cid) continue;

      const name = String(pData[i][1] || cid).trim();
      campaignsMap[cid] = name;

      configsMap[cid] = configsMap[cid] || { id: cid, name: name };
      const c = configsMap[cid];

      c.id = cid;
      if (!c.name || c.name === cid) c.name = name;
      c.aiProvider = String(c.aiProvider || pData[i][3] || 'GEMINI');
      c.aiModel = String(c.aiModel || pData[i][4] || 'gemini-3.8-flash');
      if (pData[i][5] && !c.titles) c.titles = String(pData[i][5]);
      if (pData[i][6] && !c.searchKeywords) c.searchKeywords = String(pData[i][6]);
      if (pData[i][7] && !c.campaignFocus) c.campaignFocus = String(pData[i][7]);
      c.evidencePeriodMonths = Number(c.evidencePeriodMonths || parseSafeCampaignNumeric_(pData[i][8])) || 12;
      c.researchDepth = String(c.researchDepth || pData[i][9] || 'Deep Persona');
      if (c.incEsgObligations === undefined && pData[i][10] !== '') c.incEsgObligations = String(pData[i][10]).toUpperCase() !== 'FALSE';
      if (c.incEsgMaturity === undefined && pData[i][11] !== '') c.incEsgMaturity = String(pData[i][11]).toUpperCase() !== 'FALSE';
      if (c.incSupplyChain === undefined && pData[i][12] !== '') c.incSupplyChain = String(pData[i][12]).toUpperCase() !== 'FALSE';
      if (c.incScope3 === undefined && pData[i][13] !== '') c.incScope3 = String(pData[i][13]).toUpperCase() !== 'FALSE';
      if (c.incDataChallenges === undefined && pData[i][14] !== '') c.incDataChallenges = String(pData[i][14]).toUpperCase() !== 'FALSE';
      if (c.incDigital === undefined && pData[i][15] !== '') c.incDigital = String(pData[i][15]).toUpperCase() !== 'FALSE';
      if (c.incEsgGovernance === undefined && pData[i][16] !== '') c.incEsgGovernance = String(pData[i][16]).toUpperCase() !== 'FALSE';
      if (c.incTriggers === undefined && pData[i][17] !== '') c.incTriggers = String(pData[i][17]).toUpperCase() !== 'FALSE';
      if (c.incEvaluater === undefined && pData[i][18] !== '') c.incEvaluater = String(pData[i][18]).toUpperCase() !== 'FALSE';
      if (!c.weightEsgReporting && pData[i][19]) c.weightEsgReporting = String(pData[i][19]);
      if (!c.weightSupplyChain && pData[i][20]) c.weightSupplyChain = String(pData[i][20]);
      if (!c.weightDataChallenges && pData[i][21]) c.weightDataChallenges = String(pData[i][21]);
      if (!c.weightDigital && pData[i][22]) c.weightDigital = String(pData[i][22]);
      if (!c.weightTriggers && pData[i][23]) c.weightTriggers = String(pData[i][23]);
      if (!c.weightEvaluaterFit && pData[i][24]) c.weightEvaluaterFit = String(pData[i][24]);
      c.minCompanyScore = Number(c.minCompanyScore || parseSafeCampaignNumeric_(pData[i][25])) || 80;
      c.minPersonaScore = Number(c.minPersonaScore || parseSafeCampaignNumeric_(pData[i][26])) || 85;
      c.minEmailPriorityScore = Number(c.minEmailPriorityScore || parseSafeCampaignNumeric_(pData[i][27])) || 80;
      if (c.humanApprovalRequired === undefined && pData[i][28] !== '') c.humanApprovalRequired = String(pData[i][28]).toUpperCase() !== 'FALSE';
      if (c.autoSendEnabled === undefined && pData[i][29] !== '') c.autoSendEnabled = String(pData[i][29]).toUpperCase() === 'TRUE';
      if (!c.weightEmailCompanyScore && pData[i][30]) c.weightEmailCompanyScore = String(pData[i][30]);
      if (!c.weightEmailPersonaFit && pData[i][31]) c.weightEmailPersonaFit = String(pData[i][31]);
      if (!c.weightEmailBuyingSignal && pData[i][32]) c.weightEmailBuyingSignal = String(pData[i][32]);
      if (pData[i][33] && !c.customBusinessNeeds) c.customBusinessNeeds = String(pData[i][33]);
      c.limit = Number(c.limit || parseSafeCampaignNumeric_(pData[i][34])) || 5;
      c.emailVerifier = String(c.emailVerifier || pData[i][35] || 'ZeroBounce');
      if (pData[i][36] && !c.emailEndpoint) c.emailEndpoint = String(pData[i][36]);
      c.emailSenderProvider = String(c.emailSenderProvider || pData[i][37] || 'Workspace Gmail (Drafts Only)');
      if (pData[i][38] && !c.emailSenderEndpoint) c.emailSenderEndpoint = String(pData[i][38]);
    }
  }

  if (activeCampaignId && configsMap[activeCampaignId]) {
    const c = configsMap[activeCampaignId];
    if (activeConfigDict['CAMPAIGNNAME']) c.name = String(activeConfigDict['CAMPAIGNNAME']);
    if (activeConfigDict['DEPARTMENT']) c.department = String(activeConfigDict['DEPARTMENT']);
    if (activeConfigDict['MAXCOMPANIES']) c.limit = Number(activeConfigDict['MAXCOMPANIES']) || 5;
    if (activeConfigDict['EMAILSENDERPROVIDER']) c.emailSenderProvider = String(activeConfigDict['EMAILSENDERPROVIDER']);
    if (!c.senderEmail && (activeConfigDict['SENDEREMAIL'] || activeConfigDict['SENDERMAILBOX'])) {
      c.senderEmail = String(activeConfigDict['SENDEREMAIL'] || activeConfigDict['SENDERMAILBOX']).trim();
    }
    if (!c.replyToEmail && (activeConfigDict['REPLYTOEMAIL'] || activeConfigDict['REPLYTO'])) {
      c.replyToEmail = String(activeConfigDict['REPLYTOEMAIL'] || activeConfigDict['REPLYTO']).trim();
    }
  }

  for (let cid in configsMap) {
    if (!configsMap[cid].department) {
      configsMap[cid].department = resolveCampaignClientName_(cid, '', ss);
    }
  }

  for (let cid in campaignsMap) {
    const dept = (configsMap[cid] && configsMap[cid].department)
      ? configsMap[cid].department
      : resolveCampaignClientName_(cid, '', ss);
    campaignsList.push({ id: cid, name: campaignsMap[cid], department: dept });
  }

  return {
    campaigns: campaignsList,
    configsMap: configsMap,
    activeCampaignId: activeCampaignId || (campaignsList.length > 0 ? campaignsList[0].id : 'CMP-005'),
    clients: getRegisteredClientsList_(ss),
    inactiveClients: getInactiveClientsList_(ss)
  };
}

function getAvailableCampaignsList() {
  return getSidebarInitialData().campaigns;
}

/**
 * Activates a selected campaign: reads full configuration from all 28 Campaigns
 * columns + Persona Mapper Settings, updates Campaign Configuration + Dashboard
 * WITHOUT overwriting existing Campaigns/Persona Mapper Settings rows, and applies filters.
 */
function selectAndActivateCampaign(cid) {
  if (!cid) return { success: false, message: 'No Campaign ID provided.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return { success: false, message: 'Spreadsheet not accessible.' };

  const sidebarData = getSidebarInitialData();
  const config = (sidebarData.configsMap && sidebarData.configsMap[cid])
    ? Object.assign({}, sidebarData.configsMap[cid])
    : { id: cid, name: cid };

  studioDirectSaveCampaignPipeline(config, true);

  return {
    success: true,
    campaignId: cid,
    config: config,
    message: 'Active campaign set to ' + cid + ' and synced to Dashboard, Campaigns, Persona Mapper Settings, Email Deliverability & Agent Settings!'
  };
}

/**
 * Direct Campaign Pipeline Saver (LOCKED 38 PARAMETERS SCHEMA):
 * Preserves the exact 38 parameter rows in Campaign Configuration without erasing validations.
 * Honors isActivationOnly so switching campaigns never overwrites saved rows in
 * Campaigns or Persona Mapper Settings.
 */
function studioDirectSaveCampaignPipeline(payload, isActivationOnly) {
  if (!payload || !payload.id) {
    throw new Error('Campaign ID is required to save configuration.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) throw new Error('No active spreadsheet found.');

  const campId = payload.id;
  const campName = payload.name || campId;
  const clientOwner = String(payload.department || '').trim() || 'Growth Target';

  // Ensure new Client / Owner is registered in Lists & Settings before writing to strict-validated cells
  const registeredClients = getRegisteredClientsList_(ss);
  const clientExists = registeredClients.some(function(c) {
    return c.toLowerCase() === clientOwner.toLowerCase();
  });
  if (!clientExists && !isActivationOnly) {
    addClientAccountToSettings(clientOwner);
  }

  let status = payload.status || 'Active';
  const campSheetExisting = ss.getSheetByName('Campaigns');
  if (campSheetExisting && campSheetExisting.getLastRow() > 1 && !payload.status) {
    const cRows = campSheetExisting.getDataRange().getValues();
    for (let r = 1; r < cRows.length; r++) {
      if (String(cRows[r][0]).trim() === campId && cRows[r][3]) {
        status = String(cRows[r][3]).trim();
        break;
      }
    }
  }

  let configSheet = ss.getSheetByName('Campaign Configuration');
  if (!configSheet) configSheet = ss.insertSheet('Campaign Configuration');

  const configRows = [
    ['Parameter', 'Configured Value', 'Description'],
    ['CAMPAIGN_ID', campId, 'Active Campaign Identifier'],
    ['CAMPAIGN_NAME', campName, 'Active Campaign Name'],
    ['DEPARTMENT', clientOwner, 'Client / Account Owner'],
    ['TARGET_COUNTRIES', payload.country || '', 'Target Geographies / Regions'],
    ['TARGET_INDUSTRIES', payload.industry || '', 'Target Industry Verticals'],
    ['COMPANY_SIZE_RANGE', payload.size || '', 'Target Employee Size Range'],
    ['MAX_COMPANIES', Number(payload.limit) || 5, 'Maximum target companies to discover per run'],
    ['RESEARCH_DEPTH', payload.researchDepth || 'Deep Persona', 'Depth of company research and profiling'],
    ['EVIDENCE_PERIOD_MONTHS', Number(payload.evidencePeriodMonths) || 12, 'Historical lookback window in months'],
    ['MIN_COMPANY_SCORE', Number(payload.minCompanyScore) || 80, 'Minimum company qualification threshold (0-100)'],
    ['MIN_PERSONA_SCORE', Number(payload.minPersonaScore) || 85, 'Minimum persona fit confidence threshold (0-100)'],
    ['MIN_EMAIL_PRIORITY_SCORE', Number(payload.minEmailPriorityScore) || 80, 'Minimum email priority score for sending (0-100)'],
    ['TARGET_TITLES', payload.titles || '', 'Semicolon-delimited target job titles'],
    ['CAMPAIGN_FOCUS', payload.campaignFocus || '', 'Primary value proposition hook'],
    ['SEARCH_KEYWORDS', payload.searchKeywords || '', 'Search discovery query keywords'],
    ['CUSTOM_BUSINESS_NEEDS', payload.customBusinessNeeds || '', 'Manual client business needs & buying triggers'],
    ['AI_PROVIDER', payload.aiProvider || 'GEMINI', 'Selected AI engine (GEMINI or OPENAI)'],
    ['PRIMARY_AI_PROVIDER', payload.aiProvider || 'GEMINI', 'Selected AI engine (legacy alias)'],
    ['AI_MODEL', payload.aiModel || 'gemini-3.8-flash', 'AI Model version name'],
    ['EMAIL_VERIFICATION_PROVIDER', payload.emailVerifier || 'ZeroBounce', 'Email verification vendor'],
    ['EMAIL_VERIFICATION_ENDPOINT', payload.emailEndpoint || '', 'Email verification REST endpoint URL'],
    ['EMAIL_SENDER_PROVIDER', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform'],
    ['DISPATCH_MODE', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform mode'],
    ['EMAIL_SENDER_ENDPOINT', payload.emailSenderEndpoint || '', 'Sender campaign UUID or webhook URL'],
    ['SENDER_NAME', payload.senderName || '', 'Display name for outreach email signature'],
    ['SENDER_TITLE', payload.senderTitle || '', 'Sender job title for email signature'],
    ['SENDER_COMPANY', payload.senderCompany || '', 'Sender company name for email signature'],
    ['INCLUDE_ESG_OBLIGATIONS', payload.incEsgObligations !== false ? 'TRUE' : 'FALSE', 'Evaluate CSRD/GRI obligations'],
    ['INCLUDE_ESG_MATURITY', payload.incEsgMaturity !== false ? 'TRUE' : 'FALSE', 'Evaluate ESG reporting maturity'],
    ['INCLUDE_SUPPLY_CHAIN', payload.incSupplyChain !== false ? 'TRUE' : 'FALSE', 'Evaluate supply chain complexity'],
    ['INCLUDE_SCOPE3', payload.incScope3 !== false ? 'TRUE' : 'FALSE', 'Evaluate Scope 3 emissions exposure'],
    ['INCLUDE_DATA_CHALLENGES', payload.incDataChallenges !== false ? 'TRUE' : 'FALSE', 'Evaluate sustainability data challenges'],
    ['INCLUDE_DIGITAL', payload.incDigital !== false ? 'TRUE' : 'FALSE', 'Evaluate digital transformation & AI adoption'],
    ['INCLUDE_ESG_GOVERNANCE', payload.incEsgGovernance !== false ? 'TRUE' : 'FALSE', 'Evaluate ESG board oversight'],
    ['INCLUDE_TRIGGERS', payload.incTriggers !== false ? 'TRUE' : 'FALSE', 'Evaluate commercial buying triggers'],
    ['INCLUDE_EVALUATER', payload.incEvaluater !== false ? 'TRUE' : 'FALSE', 'Evaluate eValuater solution fit'],
    ['HUMAN_APPROVAL_REQUIRED', payload.humanApprovalRequired !== false ? 'TRUE' : 'FALSE', 'Requires human review in queue before sending'],
    ['AUTO_SEND_ENABLED', payload.autoSendEnabled === true ? 'TRUE' : 'FALSE', 'Auto-send emails immediately upon approval']
  ];

  const currentConfigRows = configSheet.getLastRow();
  if (currentConfigRows > 0) {
    configSheet.getRange(1, 1, Math.max(currentConfigRows, configRows.length), 3).clearContent();
  }
  configSheet.getRange(1, 1, configRows.length, 3).setValues(configRows);
  configSheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe');
  configSheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  let campSheet = ss.getSheetByName('Campaigns');
  if (!campSheet) campSheet = ss.insertSheet('Campaigns');

  if (campSheet.getLastRow() <= 0 || campSheet.getLastColumn() < CAMPAIGNS_MASTER_HEADERS.length) {
    campSheet.getRange(1, 1, 1, CAMPAIGNS_MASTER_HEADERS.length).setValues([CAMPAIGNS_MASTER_HEADERS]);
    campSheet.getRange(1, 1, 1, CAMPAIGNS_MASTER_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe');
    campSheet.setFrozenRows(1);
  }

  const cLast = campSheet.getLastRow();
  let cTargetRow = 2;
  let cFound = false;

  if (cLast > 1) {
    const campA = campSheet.getRange(2, 1, cLast - 1, 1).getValues();
    for (let cr = 0; cr < campA.length; cr++) {
      if (String(campA[cr][0] || '').trim() === campId) {
        cTargetRow = cr + 2;
        cFound = true;
        break;
      }
    }
    if (!cFound) cTargetRow = cLast + 1;
  }

  if (!isActivationOnly || !cFound) {
    const normalizedPayload = Object.assign({}, payload, {
      department: clientOwner,
      senderEmail: String(payload.senderEmail || '').trim(),
      replyToEmail: String(payload.replyToEmail || payload.senderEmail || '').trim()
    });

    const campMasterRow = [
      campId,
      campName,
      clientOwner,
      status,
      payload.country || '',
      payload.industry || '',
      payload.size || '',
      Number(payload.limit) || 5,
      payload.researchDepth || 'Deep Persona',
      Number(payload.evidencePeriodMonths) || 12,
      Number(payload.minCompanyScore) || 80,
      Number(payload.minPersonaScore) || 85,
      Number(payload.minEmailPriorityScore) || 80,
      payload.titles || '',
      payload.campaignFocus || '',
      payload.searchKeywords || '',
      payload.customBusinessNeeds || '',
      payload.aiProvider || 'GEMINI',
      payload.aiModel || 'gemini-3.8-flash',
      payload.emailVerifier || 'ZeroBounce',
      payload.emailEndpoint || '',
      payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)',
      payload.emailSenderEndpoint || '',
      payload.senderName || '',
      payload.senderTitle || '',
      payload.senderCompany || '',
      JSON.stringify(normalizedPayload),
      new Date().toLocaleString()
    ];

    // Enforce plain integer format on Limit/Period/Score columns so Sheets doesn't treat 80/85 as 1900 dates
    campSheet.getRange(cTargetRow, 8).setNumberFormat('0');
    campSheet.getRange(cTargetRow, 10, 1, 4).setNumberFormat('0');
    campSheet.getRange(cTargetRow, 1, 1, campMasterRow.length).setValues([campMasterRow]);
    SpreadsheetApp.flush();
  }

  let pSheet = findPersonaMapperSettingsSheet(ss);
  if (!pSheet) pSheet = ss.insertSheet('Persona Mapper Settings');

  if (pSheet.getLastRow() <= 0) {
    pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length).setValues([PERSONA_MAPPER_39_HEADERS]);
    pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe');
    pSheet.setFrozenRows(1);
  }

  let pTargetRow = 2;
  const pLast = pSheet.getLastRow();
  let pFound = false;

  if (pLast > 1) {
    const colA = pSheet.getRange(2, 1, pLast - 1, 1).getValues();
    for (let r = 0; r < colA.length; r++) {
      if (String(colA[r][0] || '').trim() === campId) {
        pTargetRow = r + 2;
        pFound = true;
        break;
      }
    }
    if (!pFound) pTargetRow = pLast + 1;
  }

  if (!isActivationOnly || !pFound) {
    const pRow = [
      campId,
      campName,
      'TRUE',
      payload.aiProvider || 'GEMINI',
      payload.aiModel || 'gemini-3.8-flash',
      payload.titles || '',
      payload.searchKeywords || '',
      payload.campaignFocus || '',
      Number(payload.evidencePeriodMonths) || 12,
      payload.researchDepth || 'Deep Persona',
      payload.incEsgObligations !== false ? 'TRUE' : 'FALSE',
      payload.incEsgMaturity !== false ? 'TRUE' : 'FALSE',
      payload.incSupplyChain !== false ? 'TRUE' : 'FALSE',
      payload.incScope3 !== false ? 'TRUE' : 'FALSE',
      payload.incDataChallenges !== false ? 'TRUE' : 'FALSE',
      payload.incDigital !== false ? 'TRUE' : 'FALSE',
      payload.incEsgGovernance !== false ? 'TRUE' : 'FALSE',
      payload.incTriggers !== false ? 'TRUE' : 'FALSE',
      payload.incEvaluater !== false ? 'TRUE' : 'FALSE',
      payload.weightEsgReporting || '20%',
      payload.weightSupplyChain || '15%',
      payload.weightDataChallenges || '15%',
      payload.weightDigital || '15%',
      payload.weightTriggers || '20%',
      payload.weightEvaluaterFit || '15%',
      Number(payload.minCompanyScore) || 80,
      Number(payload.minPersonaScore) || 85,
      Number(payload.minEmailPriorityScore) || 80,
      payload.humanApprovalRequired !== false ? 'TRUE' : 'FALSE',
      payload.autoSendEnabled === true ? 'TRUE' : 'FALSE',
      payload.weightEmailCompanyScore || '40%',
      payload.weightEmailPersonaFit || '35%',
      payload.weightEmailBuyingSignal || '25%',
      payload.customBusinessNeeds || '',
      Number(payload.limit) || 5,
      payload.emailVerifier || 'ZeroBounce',
      payload.emailEndpoint || '',
      payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)',
      payload.emailSenderEndpoint || ''
    ];

    pSheet.getRange(pTargetRow, 26, 1, 3).setNumberFormat('0');
    pSheet.getRange(pTargetRow, 35).setNumberFormat('0');
    pSheet.getRange(pTargetRow, 1, 1, 39).setValues([pRow]);
    SpreadsheetApp.flush();
  }

  const dashboardPayload = Object.assign({}, payload, { department: clientOwner, status: status });
  updateDashboardForActiveCampaign(dashboardPayload);

  // Sync Agent Settings key protection badges & filter to active Client / Owner (preserves edited spend caps)
  if (typeof setupAgentSettingsTabDropdowns === 'function') {
    setupAgentSettingsTabDropdowns(ss, clientOwner);
  }
  if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
    refreshDeliverabilityFilterDropdowns_(ss);
  }

  filterOperationalSheetsByCampaign_(campId, clientOwner);

  return {
    success: true,
    campaignId: campId,
    campaignName: campName,
    message: 'Campaign ' + campId + ' (' + clientOwner + ') synchronized across all tabs and activated on Dashboard!'
  };
}

/**
 * EXACT 4-COLUMN DASHBOARD UPDATER (Client-Aware API Key Check)
 */
function updateDashboardForActiveCampaign(payloadOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) return;

  let cid = '';
  let payload = {};

  if (typeof payloadOrId === 'string') {
    cid = payloadOrId;
    const initial = getSidebarInitialData();
    payload = initial.configsMap[cid] || { id: cid, name: cid };
  } else if (payloadOrId && typeof payloadOrId === 'object') {
    payload = payloadOrId;
    cid = payload.id || payload.campaignId;
  }

  if (!cid) cid = getStudioActiveCampaignId_();

  let status = payload.status || 'Active';
  let dept = payload.department || '';
  let limit = payload.limit || '';
  let dispatchChannel = payload.emailSenderProvider || '';

  const campSheet = ss.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const cData = campSheet.getDataRange().getValues();
    for (let cr = 1; cr < cData.length; cr++) {
      if (String(cData[cr][0]).trim() === cid) {
        if (!status || status === 'Active') status = String(cData[cr][3] || status).trim();
        if (!dept && cData[cr][2]) dept = String(cData[cr][2]).trim();
        if (!limit && cData[cr][7]) limit = cData[cr][7];
        if (!dispatchChannel && cData[cr][21]) dispatchChannel = String(cData[cr][21]).trim();
        break;
      }
    }
  }

  const cfgSheet = ss.getSheetByName('Campaign Configuration');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    let cfgCid = '';
    for (let cr = 1; cr < cfgData.length; cr++) {
      if (normalizeStudioKey(cfgData[cr][0]) === 'CAMPAIGNID') {
        cfgCid = String(cfgData[cr][1]).trim();
        break;
      }
    }
    if (!cfgCid || cfgCid.toLowerCase() === cid.toLowerCase()) {
      for (let cr = 1; cr < cfgData.length; cr++) {
        const k = normalizeStudioKey(cfgData[cr][0]);
        if (!dept && (k === 'DEPARTMENT' || k === 'OPERATINGDEPARTMENT')) dept = String(cfgData[cr][1]).trim();
        if (!limit && (k === 'MAXCOMPANIES' || k === 'DISCOVERYLIMIT')) limit = cfgData[cr][1];
        if (!dispatchChannel && (k === 'EMAILSENDERPROVIDER' || k === 'DISPATCHMODE')) dispatchChannel = String(cfgData[cr][1]).trim();
      }
    }
  }

  dept = dept || resolveCampaignClientName_(cid, '', ss);
  limit = limit || 5;
  dispatchChannel = dispatchChannel || 'Workspace Gmail (Drafts Only)';
  const limitDisplay = String(limit).includes('Companies') ? limit : (limit + ' Companies');

  const campName = payload.name || payload.campaignName || cid;
  const prov = String(payload.aiProvider || 'GEMINI').toUpperCase();
  const model = String(payload.aiModel || (prov.includes('GEMINI') ? 'gemini-3.8-flash' : 'gpt-4o')).trim();

  let aiLabel = '';
  if (prov.includes('GEMINI')) {
    aiLabel = 'Gemini ' + model + ' + Search Grounding';
  } else {
    aiLabel = 'OpenAI ' + model;
  }

  const keyInfo = resolveClientScopedSecret_(
    [prov.includes('OPENAI') ? 'OPENAI_API_KEY' : 'GEMINI_API_KEY'],
    { campaignId: cid, clientName: dept },
    { throwOnMissing: false, allowAgentSettingsForLocal: true, spreadsheet: ss }
  );
  const keyStatus = keyInfo.key
    ? 'Connected'
    : (keyInfo.isLocal ? 'Key Not Configured' : 'Missing: ' + keyInfo.expectedProperty);

  const nowTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

  let discoveredCount = 0;
  const compSheet = ss.getSheetByName('Company Intelligence');
  if (compSheet && compSheet.getLastRow() > 1) {
    const data = compSheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return normalizeStudioKey(h); });
    const campIdCol = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][campIdCol] || '').trim() === cid) discoveredCount++;
    }
  }

  let personasCount = 0;
  const persSheet = ss.getSheetByName('Persona Leads');
  if (persSheet && persSheet.getLastRow() > 1) {
    const data = persSheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return normalizeStudioKey(h); });
    const campIdCol = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][campIdCol] || '').trim() === cid) personasCount++;
    }
  }

  let draftsCount = 0;
  let suppressedCount = 0;
  const qSheet = ss.getSheetByName('Prospect Review Queue');
  if (qSheet && qSheet.getLastRow() > 1) {
    const data = qSheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return normalizeStudioKey(h); });
    const campIdCol = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const sendCol = headers.indexOf('SENDSTATUS') >= 0 ? headers.indexOf('SENDSTATUS') : 24;
    const compCol = headers.indexOf('COMPLIANCESTATUS') >= 0 ? headers.indexOf('COMPLIANCESTATUS') : 41;
    const appCol = headers.indexOf('HUMANAPPROVALSTATUS') >= 0 ? headers.indexOf('HUMANAPPROVALSTATUS') : 21;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][campIdCol] || '').trim() === cid) {
        const sendStatus = String(data[i][sendCol] || '').trim().toLowerCase();
        const isValidOutreach = (
          sendStatus === 'draft created' ||
          sendStatus === 'draft created in gmail' ||
          sendStatus === 'sent' ||
          sendStatus === 'pushed to sequencer'
        );
        if (isValidOutreach) {
          draftsCount++;
        }
        const compStatus = String(data[i][compCol] || '').trim().toLowerCase();
        const appStatus = String(data[i][appCol] || '').trim().toLowerCase();
        if (compStatus.includes('failed') || compStatus.includes('blocked') || appStatus.includes('suppression') || appStatus.includes('rejected')) {
          suppressedCount++;
        }
      }
    }
  }

  dashSheet.getRange('C1').setValue(campName);
  dashSheet.getRange('D1').setValue(status);
  dashSheet.getRange('B2').setValue(cid);

  const kpiMaxRow = Math.min(dashSheet.getLastRow(), 15);
  if (kpiMaxRow >= 4) {
    const labels = dashSheet.getRange(4, 1, kpiMaxRow - 3, 1).getValues();

    for (let r = 0; r < labels.length; r++) {
      const rowNum = r + 4;
      const label = normalizeStudioKey(labels[r][0]);

      if (label.includes('ACTIVECAMPAIGNNAME') || label === 'CAMPAIGNNAME') {
        dashSheet.getRange(rowNum, 2).setValue(campName);
        dashSheet.getRange(rowNum, 3).setValue(status);
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label === 'CAMPAIGNID' || label === 'ACTIVECAMPAIGNID') {
        dashSheet.getRange(rowNum, 2).setValue(cid);
        dashSheet.getRange(rowNum, 3).setValue(cid.includes('PILOT') ? 'Pilot Mode' : 'Production Mode');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('OPERATINGDEPARTMENT') || label.includes('DEPARTMENT') || label.includes('CLIENTOWNER') || label.includes('CLIENTACCOUNT')) {
        dashSheet.getRange(rowNum, 2).setValue(dept);
        dashSheet.getRange(rowNum, 3).setValue(keyInfo.isLocal ? 'Local Owner' : 'Client Owner');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('COMPANYDISCOVERYLIMIT') || label.includes('DISCOVERYLIMIT')) {
        dashSheet.getRange(rowNum, 2).setValue(limitDisplay);
        dashSheet.getRange(rowNum, 3).setValue('Enforced Limit');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('DISCOVERED') && (label.includes('COMPAN') || label.includes('QUALIF'))) {
        dashSheet.getRange(rowNum, 2).setValue(discoveredCount);
        dashSheet.getRange(rowNum, 3).setValue(discoveredCount > 0 ? 'Discovery Complete' : 'Pending Discovery');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('PERSONALEADS') || label.includes('IDENTIFIEDLEAD')) {
        dashSheet.getRange(rowNum, 2).setValue(personasCount);
        dashSheet.getRange(rowNum, 3).setValue(personasCount > 0 ? 'Mapping Complete' : 'Pending Mapping');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('GMAILOUTREACH') || label.includes('DRAFTSCREATED')) {
        dashSheet.getRange(rowNum, 2).setValue(draftsCount);
        dashSheet.getRange(rowNum, 3).setValue(draftsCount > 0 ? 'Ready in Gmail Drafts' : 'No Drafts Yet');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('SUPPRESSED') || label.includes('BLOCKED')) {
        dashSheet.getRange(rowNum, 2).setValue(suppressedCount);
        dashSheet.getRange(rowNum, 3).setValue('Suppression Active');
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('EMAILDISPATCH') || label.includes('DISPATCHCHANNEL')) {
        dashSheet.getRange(rowNum, 2).setValue(dispatchChannel);
        dashSheet.getRange(rowNum, 3).setValue(draftsCount);
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      } else if (label.includes('PRIMARYAIMODEL') || label.includes('AIMODELCONFIGURED')) {
        dashSheet.getRange(rowNum, 2).setValue(aiLabel);
        dashSheet.getRange(rowNum, 3).setValue(keyStatus);
        dashSheet.getRange(rowNum, 4).setValue(nowTimestamp);
      }
    }
  }

  if (dashSheet.getLastRow() >= 13) {
    if (!dashSheet.getRange('A16').getValue()) {
      dashSheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);
    }
    const rollupHeaders = [
      'Campaign ID',
      'Campaign Name',
      'Department',
      'Status',
      'Companies Found',
      'Gmail Drafts',
      'Last Run'
    ];
    dashSheet.getRange(17, 1, 1, 7)
      .setValues([rollupHeaders])
      .setFontWeight('bold')
      .setBackground('#dadce0')
      .setFontColor('#202124');

    let existingRow = -1;
    const currentMax = dashSheet.getLastRow();
    for (let r = 18; r <= currentMax; r++) {
      const rowId = String(dashSheet.getRange(r, 1).getValue()).trim();
      if (rowId === cid) {
        existingRow = r;
        break;
      }
    }

    const campaignRowData = [
      cid,
      campName,
      dept,
      status,
      discoveredCount,
      draftsCount,
      new Date().toLocaleString()
    ];

    if (existingRow !== -1) {
      dashSheet.getRange(existingRow, 1, 1, 7).setValues([campaignRowData]);
    } else {
      const insertAt = Math.max(18, currentMax + 1);
      dashSheet.getRange(insertAt, 1, 1, 7).setValues([campaignRowData]);
    }

    filterDashboardRollupByClient_(ss, dept);
  }

  SpreadsheetApp.flush();
}

/**
 * Ensures Column J in Company Database is set as "Campaign ID" and backfills
 * unassigned pilot rows automatically without corrupting user data.
 */
function ensureCompanyDatabaseCampaignIdColumn_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let campaignColIdx = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = normalizeStudioKey(headers[c]);
    if (h === 'CAMPAIGNID' || h === 'CAMPAIGN' || h === 'CAMPAIGNS') {
      campaignColIdx = c + 1;
      break;
    }
  }

  if (campaignColIdx === -1) {
    sheet.getRange(1, 10).setValue('Campaign ID')
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const colJVals = sheet.getRange(2, 10, lastRow - 1, 1).getValues();
      for (let r = 0; r < colJVals.length; r++) {
        if (!colJVals[r][0]) {
          sheet.getRange(r + 2, 10).setValue('PILOT-CMP-001');
        }
      }
    }
  }
}

/**
 * Applies Google Sheets native FilterCriteria to operational tabs (including API_Usage_Log,
 * Persona Mapper Settings, and EmailEvents) so rows from other campaigns are hidden,
 * displaying exclusively the active campaign. Also filters "Campaigns", "Dashboard" Rollup,
 * "Agent Settings", and "Email Deliverability" (B3/D3) to the active Client / Owner & Campaign.
 */
function filterOperationalSheetsByCampaign_(activeCampaignId, explicitClientName) {
  if (!activeCampaignId) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  const targetSheetNames = [
    'Prospect Review Queue',
    'Persona Leads',
    'Company Intelligence',
    'Company Database',
    'Research Evidence',
    'Execution Logs',
    'API_Usage_Log',
    'Persona Mapper Settings',
    'EmailEvents'
  ];

  targetSheetNames.forEach(function(targetName) {
    let sheet = ss.getSheetByName(targetName);
    if (!sheet) {
      const normTarget = targetName.toLowerCase().replace(/[^a-z0-9]/g, '');
      const allSheets = ss.getSheets();
      for (let s = 0; s < allSheets.length; s++) {
        if (allSheets[s].getName().toLowerCase().replace(/[^a-z0-9]/g, '') === normTarget) {
          sheet = allSheets[s];
          break;
        }
      }
    }

    if (!sheet) return;
    if (sheet.getLastRow() <= 1) {
      const stale = sheet.getFilter();
      if (stale) stale.remove();
      return;
    }

    if (sheet.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === 'companydatabase') {
      ensureCompanyDatabaseCampaignIdColumn_(sheet);
    }

    let headerRow = 1;
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let campaignColIndex = -1;

    for (let c = 0; c < headers.length; c++) {
      const h = normalizeStudioKey(headers[c]);
      if (h === 'CAMPAIGNID' || h === 'CAMPAIGN' || h === 'CAMPAIGNS' || (h.includes('CAMPAIGN') && !h.includes('NAME') && !h.includes('FOCUS'))) {
        campaignColIndex = c + 1;
        break;
      }
    }

    if (campaignColIndex === -1 && sheet.getLastRow() > 2) {
      const row2Headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (let c = 0; c < row2Headers.length; c++) {
        const h = normalizeStudioKey(row2Headers[c]);
        if (h === 'CAMPAIGNID' || h === 'CAMPAIGN' || h === 'CAMPAIGNS' || (h.includes('CAMPAIGN') && !h.includes('NAME') && !h.includes('FOCUS'))) {
          campaignColIndex = c + 1;
          headerRow = 2;
          break;
        }
      }
    }

    if (campaignColIndex === -1) return;

    const startRow = headerRow;
    const expectedLastRow = sheet.getLastRow();
    const expectedLastCol = sheet.getLastColumn();
    const numDataRows = expectedLastRow - startRow;
    if (numDataRows < 1) return;

    let filter = sheet.getFilter();
    if (filter) {
      const fRange = filter.getRange();
      if (fRange.getRow() !== startRow || fRange.getLastRow() !== expectedLastRow || fRange.getLastColumn() < expectedLastCol) {
        filter.remove();
        filter = null;
      }
    }

    if (!filter) {
      const filterRange = sheet.getRange(startRow, 1, numDataRows + 1, expectedLastCol);
      filter = filterRange.createFilter();
    }

    const colValues = sheet.getRange(startRow + 1, campaignColIndex, numDataRows, 1).getValues();

    const hiddenValues = [];
    const seen = {};
    for (let i = 0; i < colValues.length; i++) {
      const val = String(colValues[i][0] || '').trim();
      if (val && val !== activeCampaignId && !seen[val]) {
        hiddenValues.push(val);
        seen[val] = true;
      }
    }

    if (hiddenValues.length > 0) {
      const criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(hiddenValues)
        .build();
      filter.setColumnFilterCriteria(campaignColIndex, criteria);
    } else {
      filter.removeColumnFilterCriteria(campaignColIndex);
    }
  });

  const activeClient = resolveCampaignClientName_(activeCampaignId, explicitClientName || '', ss);
  filterCampaignsSheetByClient_(ss, activeClient);
  filterDashboardRollupByClient_(ss, activeClient);

  if (typeof filterAgentSettingsByClient_ === 'function') {
    filterAgentSettingsByClient_(ss, activeClient);
  }
  if (typeof syncDeliverabilityTabFilters_ === 'function') {
    syncDeliverabilityTabFilters_(ss, activeClient, activeCampaignId);
  }

  SpreadsheetApp.flush();
}

/**
 * Menu action to trigger tab re-filtering manually (also refreshes Campaigns,
 * Dashboard Rollup, Email Deliverability & Agent Settings filters).
 */
function filterSheetsForActiveCampaignMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  let activeId = getStudioActiveCampaignId_();

  const cfg = ss.getSheetByName('Campaign Configuration');
  if (cfg && cfg.getLastRow() > 1) {
    const cfgId = String(cfg.getRange(2, 2).getValue() || '').trim();
    if (cfgId && cfgId !== activeId) {
      activeId = cfgId;
    }
  }

  const activeClient = resolveCampaignClientName_(activeId, '', ss);
  if (typeof setupAgentSettingsTabDropdowns === 'function') {
    setupAgentSettingsTabDropdowns(ss, activeClient);
  }
  if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
    refreshDeliverabilityFilterDropdowns_(ss);
  }

  filterOperationalSheetsByCampaign_(activeId, activeClient);

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
    ui.alert('Tabs Filtered', 'Operational tabs (including Dashboard Rollup, Campaigns, Persona Mapper Settings, API_Usage_Log, EmailEvents, Email Deliverability & Agent Settings) filtered to active campaign: ' + activeId + ' (' + activeClient + ')', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('Operational tabs filtered to active campaign: ' + activeId);
  }
}

function executeMasterPersonaMapperSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const cSheet = ss.getSheetByName('Campaign Configuration');
  const dict = {};

  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const cleanKey = normalizeStudioKey(cData[i][0]);
      if (cleanKey) dict[cleanKey] = cData[i][1];
    }
  }

  let cid = String(dict['CAMPAIGNID'] || dict['ACTIVECAMPAIGNID'] || dict['CAMPAIGN'] || '').trim();
  if (!cid) cid = getStudioActiveCampaignId_();

  const payload = {
    id: cid,
    name: String(dict['CAMPAIGNNAME'] || dict['NAME'] || cid),
    department: String(dict['DEPARTMENT'] || 'Growth Target'),
    country: String(dict['TARGETCOUNTRIES'] || ''),
    industry: String(dict['TARGETINDUSTRIES'] || ''),
    size: String(dict['COMPANYSIZERANGE'] || ''),
    limit: Number(dict['MAXCOMPANIES']) || 5,
    researchDepth: String(dict['RESEARCHDEPTH'] || 'Deep Persona'),
    evidencePeriodMonths: Number(dict['EVIDENCEPERIODMONTHS']) || 12,
    minCompanyScore: Number(dict['MINCOMPANYSCORE']) || 80,
    minPersonaScore: Number(dict['MINPERSONASCORE']) || 85,
    minEmailPriorityScore: Number(dict['MINEMAILPRIORITYSCORE']) || 80,
    titles: String(dict['TARGETTITLES'] || ''),
    campaignFocus: String(dict['CAMPAIGNFOCUS'] || ''),
    searchKeywords: String(dict['SEARCHKEYWORDS'] || ''),
    customBusinessNeeds: String(dict['CUSTOMBUSINESSNEEDS'] || ''),
    aiProvider: String(dict['AIPROVIDER'] || dict['PRIMARYAIPROVIDER'] || 'GEMINI'),
    aiModel: String(dict['AIMODEL'] || 'gemini-3.8-flash'),
    emailVerifier: String(dict['EMAILVERIFICATIONPROVIDER'] || 'ZeroBounce'),
    emailEndpoint: String(dict['EMAILVERIFICATIONENDPOINT'] || ''),
    emailSenderProvider: String(dict['EMAILSENDERPROVIDER'] || dict['DISPATCHMODE'] || 'Workspace Gmail (Drafts Only)'),
    emailSenderEndpoint: String(dict['EMAILSENDERENDPOINT'] || ''),
    senderName: String(dict['SENDERNAME'] || ''),
    senderTitle: String(dict['SENDERTITLE'] || ''),
    senderCompany: String(dict['SENDERCOMPANY'] || ''),
    senderEmail: String(dict['SENDEREMAIL'] || dict['SENDERMAILBOX'] || ''),
    replyToEmail: String(dict['REPLYTOEMAIL'] || dict['REPLYTO'] || '')
  };

  const res = studioDirectSaveCampaignPipeline(payload);

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
    ui.alert('Sync Complete: ' + res.message);
  } catch (e) {
    Logger.log('Sync Complete: ' + res.message);
  }
}

function syncActiveCampaignToPersonaMapperSettings() {
  executeMasterPersonaMapperSync();
}

function navigateToReviewQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return { success: false, message: 'Spreadsheet not found.' };
  const queueSheet = ss.getSheetByName('Prospect Review Queue');
  if (!queueSheet) {
    return { success: false, message: 'Prospect Review Queue sheet not found.' };
  }
  ss.setActiveSheet(queueSheet);
  return { success: true, message: 'Switched to Prospect Review Queue.' };
}

function runStudioSpreadsheetAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    Logger.log('Spreadsheet not accessible.');
    return;
  }

  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {}

  const activeId = getStudioActiveCampaignId_();
  const activeClient = resolveCampaignClientName_(activeId, '', ss);

  if (typeof setupAgentSettingsTabDropdowns === 'function') {
    setupAgentSettingsTabDropdowns(ss, activeClient);
  }

  const requiredTabs = [
    'Dashboard',
    'API_Usage_Log',
    'Campaigns',
    'Company Intelligence',
    'Persona Leads',
    'Persona Mapper Settings',
    'Prospect Review Queue',
    'Campaign Configuration',
    'Company Database',
    'Research Evidence',
    'Agent Settings',
    'Execution Logs',
    'Lists & Settings'
  ];

  const deliverabilityTabs = ['EmailEvents', 'Email Deliverability'];

  const present = [];
  const missing = [];
  const delivPresent = [];
  const delivMissing = [];

  requiredTabs.forEach(function(name) {
    if (name === 'Persona Mapper Settings') {
      if (findPersonaMapperSettingsSheet(ss)) present.push('Persona Mapper Settings');
      else missing.push('Persona Mapper Settings');
    } else {
      if (ss.getSheetByName(name)) present.push(name);
      else missing.push(name);
    }
  });

  deliverabilityTabs.forEach(function(name) {
    if (ss.getSheetByName(name)) delivPresent.push(name);
    else delivMissing.push(name);
  });

  let hourlySyncActive = false;
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'syncEmailDeliverabilityEvents') {
        hourlySyncActive = true;
        break;
      }
    }
  } catch (e) {}

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();
  const detectedLocalKeys = [];
  const detectedClientKeys = [];

  for (let k in scriptProps) {
    const uk = k.toUpperCase();
    if (!isStudioValidSecretValue_(scriptProps[k])) continue;
    if (uk.indexOf('CLIENT_') === 0) {
      detectedClientKeys.push(k);
    } else {
      if (uk.includes('OPENAI')) detectedLocalKeys.push('OpenAI (Project Settings)');
      if (uk.includes('GEMINI')) detectedLocalKeys.push('Gemini (Project Settings)');
    }
  }

  for (let k in userProps) {
    const uk = k.toUpperCase();
    if (!isStudioValidSecretValue_(userProps[k])) continue;
    if (uk.indexOf('CLIENT_') === 0) continue;
    if (uk.includes('OPENAI')) detectedLocalKeys.push('OpenAI (User Properties)');
    if (uk.includes('GEMINI')) detectedLocalKeys.push('Gemini (User Properties)');
  }

  const agentSheet = ss.getSheetByName('Agent Settings');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const data = agentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const val = String(data[i][j] || '').trim().toUpperCase();
        if (val === 'OPENAI_API_KEY') {
          const keyVal = String(data[i][j + 1] !== undefined ? data[i][j + 1] : (data[i][2] || '')).trim();
          if (isStudioValidSecretValue_(keyVal)) detectedLocalKeys.push('OpenAI (Agent Settings)');
        }
        if (val === 'GEMINI_API_KEY') {
          const keyVal = String(data[i][j + 1] !== undefined ? data[i][j + 1] : (data[i][2] || '')).trim();
          if (isStudioValidSecretValue_(keyVal)) detectedLocalKeys.push('Gemini (Agent Settings)');
        }
      }
    }
  }

  let dispatchChannel = 'Workspace Gmail (Drafts Only)';
  const cfgSheet = ss.getSheetByName('Campaign Configuration');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    for (let r = 1; r < cfgData.length; r++) {
      const k = normalizeStudioKey(cfgData[r][0]);
      if ((k === 'EMAILSENDERPROVIDER' || k === 'DISPATCHMODE') && cfgData[r][1]) {
        dispatchChannel = String(cfgData[r][1]).trim();
        break;
      }
    }
  }

  const registeredClients = getRegisteredClientsList_(ss);

  let apiKeyInfo = 'Not Configured (Add key in Project Settings > Script Properties)';
  if (detectedLocalKeys.length > 0) {
    apiKeyInfo = 'Local Connected: ' + Array.from(new Set(detectedLocalKeys)).join(', ');
  }

  let message = '📊 SPREADSHEET SYSTEM HEALTH CHECK:\n\n';
  message += '• Active Campaign: ' + activeId + ' (Client / Owner: ' + activeClient + ')\n';
  message += '• Registered Clients: ' + registeredClients.join(', ') + '\n';
  message += '• Core Project Tabs Verified: ' + present.length + ' of ' + requiredTabs.length + ' Tabs Active\n';
  if (missing.length > 0) {
    message += '⚠️ Missing Core Tabs: ' + missing.join(', ') + '\n';
  } else {
    message += '✅ All 13 core project tabs exist and match system schema.\n';
  }

  message += '• Deliverability Tabs: ' + delivPresent.length + ' of ' + deliverabilityTabs.length + ' Active' +
    (delivMissing.length > 0 ? (' (Missing: ' + delivMissing.join(', ') + ' — run setupEmailDeliverabilityModule)') : ' ✅') + '\n';
  message += '• Hourly Deliverability Sync Trigger: ' + (hourlySyncActive ? 'Active ✅' : 'Not Installed (Enable from menu)') + '\n\n';

  message += '• Local Agency AI Status: ' + apiKeyInfo + '\n';
  message += '• Client-Isolated Properties Detected: ' + detectedClientKeys.length + ' key(s)\n';
  message += '• Email Dispatch Channel: ' + dispatchChannel + '\n';
  message += '• Suppression Engine: Active (checking Lists & Settings)\n';
  message += '• Human Review Queue: Operational\n\n';
  message += 'Your workspace is healthy and ready to run!';

  if (ui) {
    ui.alert('Spreadsheet Setup Status', message, ui.ButtonSet.OK);
  } else {
    Logger.log(message);
  }
}

function showSetupStatus() {
  runStudioSpreadsheetAudit();
}

function approveSelectedProspect() {
  updateSelectedProspectStatus('Approved');
}

function markSelectedProspectNeedsEditing() {
  updateSelectedProspectStatus('Needs Review');
}

function rejectSelectedProspect() {
  updateSelectedProspectStatus('Rejected');
}

/**
 * Updates Review Status, Human Approval Status, Reviewer metadata, and Queue Status
 * across ALL selected visible prospect rows in "Prospect Review Queue" (supporting
 * both dragged selections and Ctrl/Cmd-clicked separate rows, while skipping rows
 * hidden by active filters).
 */
function updateSelectedProspectStatus(statusText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;
  const sheet = ss.getActiveSheet ? ss.getActiveSheet() : null;
  if (!sheet || sheet.getName() !== 'Prospect Review Queue') {
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
      ui.alert('Please switch to the "Prospect Review Queue" tab before updating prospects.');
    } catch (e) {
      Logger.log('Please switch to the "Prospect Review Queue" tab before updating prospects.');
    }
    return;
  }

  const lastDataRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastDataRow <= 1 || lastCol < 1) {
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
      ui.alert('No prospect rows found in Prospect Review Queue.');
    } catch (e) {
      Logger.log('No prospect rows found in Prospect Review Queue.');
    }
    return;
  }

  let ranges = [];
  try {
    const rangeList = sheet.getActiveRangeList ? sheet.getActiveRangeList() : null;
    if (rangeList) {
      ranges = rangeList.getRanges() || [];
    }
  } catch (e) {}

  if (!ranges || ranges.length === 0) {
    try {
      const singleRange = sheet.getActiveRange ? sheet.getActiveRange() : null;
      if (singleRange) ranges = [singleRange];
    } catch (e) {}
  }
  if (!ranges || ranges.length === 0) {
    try {
      const activeCell = sheet.getActiveCell ? sheet.getActiveCell() : null;
      if (activeCell) ranges = [activeCell];
    } catch (e) {}
  }

  const previewCols = Math.min(lastCol, 18);
  const previewRows = sheet.getRange(2, 1, lastDataRow - 1, previewCols).getValues();
  const rowSet = {};

  for (let i = 0; i < ranges.length; i++) {
    const rng = ranges[i];
    if (!rng) continue;
    const startRow = Math.max(2, rng.getRow());
    const endRow = Math.min(lastDataRow, rng.getRow() + rng.getNumRows() - 1);

    for (let r = startRow; r <= endRow; r++) {
      if (rowSet[r]) continue;
      let isHidden = false;
      try {
        if ((typeof sheet.isRowHiddenByFilter === 'function' && sheet.isRowHiddenByFilter(r)) ||
            (typeof sheet.isRowHiddenByUser === 'function' && sheet.isRowHiddenByUser(r))) {
          isHidden = true;
        }
      } catch (e) {}
      if (isHidden) continue;

      const rowPreview = previewRows[r - 2] || [];
      const hasData = rowPreview.some(function(cell) {
        return cell !== '' && cell !== null && cell !== undefined;
      });
      if (hasData) {
        rowSet[r] = true;
      }
    }
  }

  const targetRows = Object.keys(rowSet).map(Number).sort(function(a, b) { return a - b; });
  if (targetRows.length === 0) {
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
      ui.alert('Please select one or more visible prospect rows (row 2 or below).');
    } catch (e) {
      Logger.log('Please select one or more visible prospect rows (row 2 or below).');
    }
    return;
  }

  const user = getSafeReviewerEmail_();
  const now = new Date().toLocaleString();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  function resolveCol(headerName, fallbackCol) {
    const idx = headers.indexOf(headerName);
    return idx >= 0 ? idx + 1 : fallbackCol;
  }

  const reviewStatusCol = resolveCol('Review Status', 18);
  const reviewerCol = resolveCol('Reviewer', 19);
  const reviewedAtCol = resolveCol('Reviewed At', 21);
  const approvalStatusCol = resolveCol('Human Approval Status', 22);
  const approvedByCol = resolveCol('Approved By', 23);
  const approvedAtCol = resolveCol('Approved At', 24);
  const queueStatusCol = resolveCol('Queue Status', 41);

  targetRows.forEach(function(row) {
    sheet.getRange(row, reviewStatusCol).setValue(statusText);
    sheet.getRange(row, reviewerCol).setValue(user);
    sheet.getRange(row, reviewedAtCol).setValue(now);
    sheet.getRange(row, approvalStatusCol).setValue(statusText);

    if (statusText === 'Approved') {
      sheet.getRange(row, approvedByCol).setValue(user);
      sheet.getRange(row, approvedAtCol).setValue(now);
    } else if (statusText === 'Rejected') {
      sheet.getRange(row, queueStatusCol).setValue('Rejected');
    } else if (statusText === 'Needs Review') {
      sheet.getRange(row, queueStatusCol).setValue('Ready for Review');
    }
  });

  SpreadsheetApp.flush();
  const rowPreviewLabel = targetRows.length <= 8
    ? targetRows.join(', ')
    : (targetRows.slice(0, 5).join(', ') + '...');
  const toastMsg = targetRows.length === 1
    ? ('Row ' + targetRows[0] + ' updated to: ' + statusText)
    : ('Updated ' + targetRows.length + ' prospect rows (' + rowPreviewLabel + ') to: ' + statusText);
  ss.toast(toastMsg, 'Status Updated', 4);
}

function triggerPilotRun() {
  if (typeof runPilotOnRealSheets === 'function') {
    runPilotOnRealSheets();
  } else {
    Logger.log('runPilotOnRealSheets() function not found in PilotSandboxRunner.gs.');
  }
}

function triggerPilotCleanup() {
  if (typeof cleanupPilotDataFromRealSheets === 'function') {
    cleanupPilotDataFromRealSheets();
  } else {
    Logger.log('cleanupPilotDataFromRealSheets() function not found in PilotSandboxRunner.gs.');
  }
}

function triggerSetupAllDropdowns() {
  if (typeof setupAllSheetsAndDropdowns === 'function') {
    setupAllSheetsAndDropdowns();
  }
}

function preparePersonaMapperSheets() {
  if (typeof ensurePersonaMapperSheets_ === 'function') {
    ensurePersonaMapperSheets_(SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive());
  }
}

/**
 * Forces Google Apps Script to request & verify ALL manifest scopes (including
 * full Gmail read/search permission via https://mail.google.com/) before
 * reporting authorization complete.
 */
function authorizeStudioPermissions() {
  if (typeof ScriptApp !== 'undefined' && typeof ScriptApp.requireAllScopes === 'function') {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  }

  if (typeof ScriptApp !== 'undefined' && typeof ScriptApp.getAuthorizationInfo === 'function') {
    const gmailAuthInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL, ['https://mail.google.com/']);
    if (gmailAuthInfo && gmailAuthInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      const authUrl = gmailAuthInfo.getAuthorizationUrl();
      const missingMsg =
        'Gmail read permission (https://mail.google.com/) is still NOT granted.\n\n' +
        'When Google shows the permissions screen, please click "Select all" (or check the box to Read/Manage Gmail).\n\n' +
        (authUrl ? ('Open this link to authorize:\n' + authUrl) : '');
      Logger.log(missingMsg);
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Additional Permission Needed', missingMsg, ui.ButtonSet.OK);
      } catch (e) {}
      throw new Error(missingMsg);
    }
  }

  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  SpreadsheetApp.getActiveSpreadsheet();
  GmailApp.getDrafts();
  GmailApp.search('in:sent', 0, 1);
  ScriptApp.getProjectTriggers();

  let grantedScopes = [];
  try {
    if (typeof ScriptApp !== 'undefined' && typeof ScriptApp.getAuthorizationInfo === 'function') {
      grantedScopes = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizedScopes() || [];
    }
  } catch (e) {}

  const successMsg = 'Authorization verified successfully!' +
    (grantedScopes.length > 0 ? ('\n\nGranted permissions:\n• ' + grantedScopes.join('\n• ')) : '');

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
    ui.alert('Authorization Complete', successMsg, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(successMsg);
  }
}