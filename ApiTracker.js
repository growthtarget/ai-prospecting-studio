/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * CENTRALIZED API USAGE TRACKER & CLIENT-ISOLATED BUDGET GUARDRAIL ENGINE
 * ============================================================================
 * File: ApiTracker.gs / ApiTracker.js
 * 
 * PURPOSE:
 * 1. Logs every external API call (LLMs, Verifiers) to "API_Usage_Log" (10 canonical cols).
 * 2. Tags usage entries with Client / Owner context for effortless client spend auditing.
 * 3. Automatically calculates est. dollar cost based on token counts or unit pricing.
 * 4. Enforces pre-flight daily spend hard caps isolated per Client / Owner
 *    (reading CLIENT_{NAME}_DAILY_SPEND_CAP_USD from Agent Settings or Script Properties).
 * 5. Dispatches automated low-balance email alerts using per-client MIN_VERIFIER_CREDITS
 *    and strict client-scoped verifier keys.
 * ============================================================================
 */

/**
 * Resolves the Client / Owner name for a Campaign ID safely even if called standalone.
 */
function resolveTrackerClientName_(ss, campaignId, explicitClientName) {
  if (explicitClientName && String(explicitClientName).trim()) {
    return String(explicitClientName).trim();
  }
  if (typeof resolveCampaignClientName_ === 'function') {
    return resolveCampaignClientName_(campaignId, '', ss);
  }
  if (campaignId && String(campaignId).toUpperCase().indexOf('PILOT') === 0) {
    return 'Pilot Acme';
  }
  return 'Growth Target';
}

/**
 * Normalizes a client name for safe comparison across logs and campaigns.
 */
function normalizeTrackerClientKey_(clientName) {
  if (typeof isLocalAgencyClient_ === 'function' && isLocalAgencyClient_(clientName)) {
    return 'GROWTHTARGET';
  }
  const norm = String(clientName || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!norm || norm === 'LOCAL' || norm === 'AGENCY' || norm === 'DEFAULT') {
    return 'GROWTHTARGET';
  }
  return norm;
}

/**
 * Builds canonical Script Property prefix e.g. "eValuater" -> "CLIENT_EVALUATER_".
 */
function getTrackerClientPrefix_(clientName) {
  if (typeof normalizeClientPropertyPrefix_ === 'function') {
    return normalizeClientPropertyPrefix_(clientName);
  }
  const clean = String(clientName || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return clean ? ('CLIENT_' + clean + '_') : '';
}

/**
 * Builds a lookup map of Campaign ID -> normalized Client key from "Campaigns" and "Campaign Configuration".
 */
function buildCampaignToClientKeyMap_(ss) {
  const map = {};
  if (!ss) return map;

  const campSheet = ss.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const data = campSheet.getDataRange().getValues();
    const headers = data[0].map(function(h) {
      return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const deptIdx = headers.indexOf('DEPARTMENT') >= 0 ? headers.indexOf('DEPARTMENT') : 2;
    for (let r = 1; r < data.length; r++) {
      const cid = String(data[r][idIdx] || '').trim();
      const dept = String(data[r][deptIdx] || '').trim();
      if (cid) {
        map[cid.toLowerCase()] = normalizeTrackerClientKey_(dept || (cid.toUpperCase().indexOf('PILOT') === 0 ? 'Pilot Acme' : 'Growth Target'));
      }
    }
  }

  const cfgSheet = ss.getSheetByName('Campaign Configuration');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    let cfgCid = '';
    let cfgDept = '';
    for (let r = 1; r < cfgData.length; r++) {
      const k = String(cfgData[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (k === 'CAMPAIGNID') cfgCid = String(cfgData[r][1] || '').trim();
      if (k === 'DEPARTMENT') cfgDept = String(cfgData[r][1] || '').trim();
    }
    if (cfgCid && cfgDept && !map[cfgCid.toLowerCase()]) {
      map[cfgCid.toLowerCase()] = normalizeTrackerClientKey_(cfgDept);
    }
  }

  return map;
}

/**
 * Centralized logging function for any external API request.
 * @param {Object} data
 * @param {string} data.campaignId - Active campaign ID
 * @param {string} [data.clientName] - Client / Owner name (auto-resolved if omitted)
 * @param {string} data.serviceName - e.g. 'GEMINI', 'OPENAI', 'ZeroBounce'
 * @param {string} data.operation - e.g. 'CompanyDiscovery', 'PersonaMapping', 'EmailVerification'
 * @param {number} data.httpStatus - HTTP response code (e.g. 200, 429)
 * @param {number} data.inputUnits - Input token count or verification units
 * @param {number} data.outputUnits - Output token count (for LLMs)
 * @param {string} data.status - 'SUCCESS' or 'ERROR'
 * @param {string} data.details - Error messages or context notes
 * @param {string} [data.modelName] - Specific model/service tier for pricing calculation
 */
function logApiUsage_(data) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      Logger.log('>>> WARNING: ApiTracker could not acquire lock. Skipping log entry to prevent timeout.');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
    if (!ss) return;

    let sheet = ss.getSheetByName('API_Usage_Log');
    if (!sheet) {
      sheet = ss.insertSheet('API_Usage_Log');
      sheet.appendRow([
        'Timestamp', 'Campaign ID', 'Service Name', 'Operation', 'HTTP Status',
        'Input Tokens / Units', 'Output Tokens', 'Est. Cost ($)', 'Status', 'Details / Error'
      ]);
    }

    const timestamp = new Date();
    const campaignId = data.campaignId || 'CMP-DEFAULT';
    const clientName = resolveTrackerClientName_(ss, campaignId, data.clientName);
    const serviceName = data.serviceName || 'Unknown';
    const operation = data.operation || 'API Call';
    const httpStatus = data.httpStatus || 200;
    const inputUnits = Number(data.inputUnits) || 0;
    const outputUnits = Number(data.outputUnits) || 0;
    const status = data.status || 'SUCCESS';
    const rawDetails = String(data.details || '');
    const details = rawDetails.indexOf('[Client:') >= 0
      ? rawDetails
      : ('[Client: ' + clientName + '] ' + rawDetails).trim();
    const modelName = data.modelName || serviceName;

    const cost = calculateApiCost_(modelName, inputUnits, outputUnits, serviceName);

    sheet.appendRow([
      timestamp,
      campaignId,
      serviceName,
      operation,
      httpStatus,
      inputUnits,
      outputUnits,
      cost,
      status,
      details
    ]);

  } catch (err) {
    Logger.log('>>> ERROR in logApiUsage_: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Calculates estimated dollar cost based on APP_CONFIG pricing or defaults.
 */
function calculateApiCost_(modelName, inputUnits, outputUnits, serviceName) {
  try {
    const pricing = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails && APP_CONFIG.guardrails.pricing)
      ? APP_CONFIG.guardrails.pricing
      : {};

    const rateKey = String(modelName || '').trim();
    const serviceKey = String(serviceName || '').trim();

    if (pricing[rateKey]) {
      const spec = pricing[rateKey];
      if (spec.inputPer1M !== undefined) {
        const inputCost = (inputUnits / 1000000) * spec.inputPer1M;
        const outputCost = (outputUnits / 1000000) * spec.outputPer1M;
        return Number((inputCost + outputCost).toFixed(6));
      } else if (spec.costPerUnit !== undefined) {
        return Number((inputUnits * spec.costPerUnit).toFixed(4));
      }
    }

    if (rateKey.toLowerCase().indexOf('gemini-3.5-flash-lite') >= 0) {
      const inputCost = (inputUnits / 1000000) * 0.075;
      const outputCost = (outputUnits / 1000000) * 0.30;
      return Number((inputCost + outputCost).toFixed(6));
    }

    if (pricing[serviceKey]) {
      const spec = pricing[serviceKey];
      if (spec.costPerUnit !== undefined) {
        return Number((inputUnits * spec.costPerUnit).toFixed(4));
      }
    }

    return 0.0001;
  } catch (e) {
    return 0;
  }
}

/**
 * Pre-Flight Guardrail: Checks today's API spend scoped strictly to the target Client / Owner.
 * Reads CLIENT_{NAME}_DAILY_SPEND_CAP_USD from Agent Settings first (or Script Properties),
 * falling back to DAILY_SPEND_CAP_USD if unset for that client.
 */
function checkApiBudgetGuardrails_(campaignId, explicitClientName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
    if (!ss) return;

    const sheet = ss.getSheetByName('API_Usage_Log');
    if (!sheet || sheet.getLastRow() <= 1) return;

    const targetClientName = resolveTrackerClientName_(ss, campaignId, explicitClientName);
    const targetClientKey = normalizeTrackerClientKey_(targetClientName);
    const isExternalClient = (targetClientKey !== 'GROWTHTARGET');
    const clientCapProp = isExternalClient ? (getTrackerClientPrefix_(targetClientName) + 'DAILY_SPEND_CAP_USD') : '';
    const normClientCapKey = clientCapProp.toUpperCase().replace(/[^A-Z0-9]/g, '');

    let defaultDailyCap = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails)
      ? APP_CONFIG.guardrails.defaultDailySpendCapUsd
      : 10.00;
    let clientDailyCap = null;

    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const rawKey = String(rows[i][1] || '').trim().toUpperCase();
        const cleanKey = rawKey.replace(/[^A-Z0-9]/g, '');
        const rawCell = rows[i][2];
        const parsedVal = (rawCell !== '' && rawCell !== null && !(rawCell instanceof Date)) ? parseFloat(rawCell) : NaN;

        if (cleanKey === 'DAILYSPENDCAPUSD' && !isNaN(parsedVal) && parsedVal >= 0) {
          defaultDailyCap = parsedVal;
        } else if (isExternalClient && cleanKey === normClientCapKey && !isNaN(parsedVal) && parsedVal >= 0) {
          clientDailyCap = parsedVal;
        }
      }
    }

    if (isExternalClient && clientDailyCap === null && clientCapProp) {
      const propVal = PropertiesService.getScriptProperties().getProperty(clientCapProp);
      const parsedProp = (propVal !== '' && propVal !== null) ? parseFloat(propVal) : NaN;
      if (!isNaN(parsedProp) && parsedProp >= 0) {
        clientDailyCap = parsedProp;
      }
    }

    const dailyCap = (isExternalClient && clientDailyCap !== null) ? clientDailyCap : defaultDailyCap;

    const campaignMap = buildCampaignToClientKeyMap_(ss);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const todayStr = new Date().toDateString();
    let todaySpend = 0;

    for (let r = 0; r < data.length; r++) {
      const rowDate = new Date(data[r][0]);
      if (rowDate.toDateString() !== todayStr) continue;

      const rowCampId = String(data[r][1] || '').trim().toLowerCase();
      const rowDetails = String(data[r][9] || '');
      let rowClientKey = '';

      const tagMatch = rowDetails.match(/\[Client:\s*([^\]]+)\]/i);
      if (tagMatch && tagMatch[1]) {
        rowClientKey = normalizeTrackerClientKey_(tagMatch[1]);
      } else if (rowCampId && campaignMap[rowCampId]) {
        rowClientKey = campaignMap[rowCampId];
      } else if (rowCampId.indexOf('pilot') === 0) {
        rowClientKey = 'PILOTACME';
      } else {
        rowClientKey = 'GROWTHTARGET';
      }

      if (rowClientKey === targetClientKey) {
        const cost = parseFloat(data[r][7]) || 0;
        todaySpend += cost;
      }
    }

    Logger.log('>>> Pre-Flight Budget Check [' + targetClientName + ']: Today Spend = $' + todaySpend.toFixed(4) + ' / Cap = $' + dailyCap.toFixed(2));

    if (todaySpend >= dailyCap) {
      throw new Error(
        '🛑 RUNAWAY LOOP ABORTED: Daily API Spend Hard Cap ($' + dailyCap.toFixed(2) + ') reached for client "' + targetClientName + '"! ' +
        'Today\'s spend for this client is $' + todaySpend.toFixed(4) + '. Adjust cap in Agent Settings or wait until tomorrow.'
      );
    }

  } catch (err) {
    if (err.toString().includes('RUNAWAY LOOP ABORTED')) {
      throw err;
    }
    Logger.log('>>> Warning in checkApiBudgetGuardrails_: ' + err.toString());
  }
}

/**
 * Checks ZeroBounce verifier balance strictly for the specified campaign/client,
 * honoring per-client CLIENT_{NAME}_MIN_VERIFIER_CREDITS from Agent Settings.
 */
function checkVerifierBalanceAndAlert_(campaignId, explicitClientName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
    if (!ss) return;

    const activeCampId = campaignId || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : '');
    const clientName = resolveTrackerClientName_(ss, activeCampId, explicitClientName);
    const clientKey = normalizeTrackerClientKey_(clientName);
    const isExternalClient = (clientKey !== 'GROWTHTARGET');
    const clientMinProp = isExternalClient ? (getTrackerClientPrefix_(clientName) + 'MIN_VERIFIER_CREDITS') : '';
    const normClientMinKey = clientMinProp.toUpperCase().replace(/[^A-Z0-9]/g, '');

    let defaultMinCredits = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails)
      ? APP_CONFIG.guardrails.defaultMinVerifierCredits
      : 500;
    let clientMinCredits = null;
    let verifierProvider = '';

    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const cleanK = String(rows[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        const rawCell = rows[i][2];

        if (cleanK === 'MINVERIFIERCREDITS') {
          const val = (rawCell !== '' && rawCell !== null && !(rawCell instanceof Date)) ? parseInt(rawCell, 10) : NaN;
          if (!isNaN(val) && val >= 0) defaultMinCredits = val;
        } else if (isExternalClient && cleanK === normClientMinKey) {
          const val = (rawCell !== '' && rawCell !== null && !(rawCell instanceof Date)) ? parseInt(rawCell, 10) : NaN;
          if (!isNaN(val) && val >= 0) clientMinCredits = val;
        } else if (cleanK === 'VERIFIERPROVIDER') {
          const val = String(rawCell || '').trim();
          if (val) verifierProvider = val;
        }
      }
    }

    if (isExternalClient && clientMinCredits === null && clientMinProp) {
      const propVal = PropertiesService.getScriptProperties().getProperty(clientMinProp);
      const parsedProp = (propVal !== '' && propVal !== null) ? parseInt(propVal, 10) : NaN;
      if (!isNaN(parsedProp) && parsedProp >= 0) {
        clientMinCredits = parsedProp;
      }
    }

    const minCredits = (isExternalClient && clientMinCredits !== null) ? clientMinCredits : defaultMinCredits;

    // Check campaign-specific verifier provider in Campaigns and Campaign Configuration
    if (activeCampId) {
      const campSheet = ss.getSheetByName('Campaigns');
      if (campSheet && campSheet.getLastRow() > 1) {
        const cData = campSheet.getDataRange().getValues();
        const headers = cData[0].map(function(h) { return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); });
        const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
        const verifIdx = headers.indexOf('EMAILVERIFIERPROVIDER');
        for (let r = 1; r < cData.length; r++) {
          if (String(cData[r][idIdx] || '').trim().toLowerCase() === activeCampId.toLowerCase()) {
            if (verifIdx >= 0 && cData[r][verifIdx]) {
              verifierProvider = String(cData[r][verifIdx]).trim();
            }
            break;
          }
        }
      }

      const cfgSheet = ss.getSheetByName('Campaign Configuration');
      if (cfgSheet && cfgSheet.getLastRow() > 1) {
        const cfgData = cfgSheet.getDataRange().getValues();
        let cfgCid = '';
        let cfgVerif = '';
        for (let r = 1; r < cfgData.length; r++) {
          const k = String(cfgData[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
          if (k === 'CAMPAIGNID') cfgCid = String(cfgData[r][1] || '').trim();
          if (k === 'EMAILVERIFICATIONPROVIDER') cfgVerif = String(cfgData[r][1] || '').trim();
        }
        if (cfgCid && cfgCid.toLowerCase() === activeCampId.toLowerCase() && cfgVerif) {
          verifierProvider = cfgVerif;
        }
      }
    }

    if ((verifierProvider || 'ZeroBounce').toLowerCase() !== 'zerobounce') {
      return;
    }

    let apiKey = '';
    if (typeof resolveClientScopedSecret_ === 'function') {
      const resolved = resolveClientScopedSecret_(
        ['ZEROBOUNCE_API_KEY', 'VERIFIER_API_KEY', 'EMAIL_VERIFICATION_API_KEY'],
        { campaignId: activeCampId, clientName: clientName },
        { throwOnMissing: false, allowAgentSettingsForLocal: true, spreadsheet: ss }
      );
      apiKey = resolved.key || '';
    }

    if (!apiKey) return;

    const url = 'https://api.zerobounce.net/v2/getcredits?api_key=' + encodeURIComponent(apiKey);
    const options = { method: 'get', muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const json = JSON.parse(response.getContentText() || '{}');
      const credits = parseInt(json.Credits, 10);

      if (!isNaN(credits) && credits < minCredits) {
        let alertEmail = '';
        if (isExternalClient) {
          const customAlertProp = getTrackerClientPrefix_(clientName) + 'ALERT_EMAIL';
          alertEmail = String(PropertiesService.getScriptProperties().getProperty(customAlertProp) || '').trim();
        }
        if (!alertEmail) {
          try { alertEmail = Session.getActiveUser().getEmail(); } catch (e) {}
        }
        if (!alertEmail) {
          try { alertEmail = Session.getEffectiveUser().getEmail(); } catch (e) {}
        }

        if (alertEmail) {
          const subject = '⚠️ Low ZeroBounce Credits Alert [' + clientName + '] (' + credits + ' remaining)';
          const body = 'Hello,\n\nYour AI Prospecting Studio automated check detected that the ZeroBounce email verification balance for client "' + clientName + '" (Campaign: ' + (activeCampId || 'Active') + ') has dropped to ' + credits + ' credits (Threshold: ' + minCredits + ').\n\nPlease top up this account to prevent enrichment runs from pausing or failing mid-batch.\n\n- AI Prospecting Studio System';
          MailApp.sendEmail(alertEmail, subject, body);
          Logger.log('>>> Low balance email alert sent to ' + alertEmail + ' for client ' + clientName + ' (Credits: ' + credits + ' / Min: ' + minCredits + ')');
        }
      }
    }
  } catch (err) {
    Logger.log('>>> Warning in checkVerifierBalanceAndAlert_: ' + err.toString());
  }
}