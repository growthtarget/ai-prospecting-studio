/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * CENTRALIZED API USAGE TRACKER, BUDGET GUARDRAILS & CONNECTION AUDIT ENGINE
 * ============================================================================
 * File: ApiTracker.gs / ApiTracker.js
 *
 * PURPOSE:
 * 1. Logs every external API call (LLMs, Serper Search, Verifiers) to "API_Usage_Log" (10 canonical cols).
 * 2. Tags usage entries with Client / Owner context for effortless client spend auditing.
 * 3. Calculates accurate dollar cost via exact & longest-prefix pricing lookup.
 * 4. Enforces pre-flight daily spend hard caps isolated per Client / Owner
 *    (reading CLIENT_{NAME}_DAILY_SPEND_CAP_USD from Agent Settings or Script Properties,
 *    including strict $0.00 lockout support even on empty logs).
 * 5. Checks verifier balances & dispatches throttled low-balance alerts for
 *    ZeroBounce, MillionVerifier, and NeverBounce.
 * 6. Runs $0.00 read-only Connection & Budget Audits for ControlPanel.html without
 *    spending AI/search/verifier credits, sending emails, or exposing raw secrets.
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
 * Safely parses a numeric budget/credit setting cell (handles "$10.00" and rejects Dates).
 */
function parseTrackerNumber_(rawVal) {
  if (rawVal === '' || rawVal === null || rawVal === undefined || rawVal instanceof Date) {
    return NaN;
  }
  const cleaned = String(rawVal).replace(/[$,\s]/g, '').trim();
  if (!cleaned) return NaN;
  const num = parseFloat(cleaned);
  return isFinite(num) ? num : NaN;
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
 * Calculates estimated dollar cost based on APP_CONFIG pricing (exact + longest-prefix match)
 * with conservative LLM/Serper fallback pricing so unmapped models never under-report.
 */
function calculateApiCost_(modelName, inputUnits, outputUnits, serviceName) {
  try {
    const pricing = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails && APP_CONFIG.guardrails.pricing)
      ? APP_CONFIG.guardrails.pricing
      : {};

    const rateKey = String(modelName || '').trim();
    const rateKeyLow = rateKey.toLowerCase();
    const serviceKey = String(serviceName || '').trim();
    const serviceKeyLow = serviceKey.toLowerCase();
    const inUnits = Number(inputUnits) || 0;
    const outUnits = Number(outputUnits) || 0;

    function computeFromSpec(spec) {
      if (!spec) return null;
      if (spec.inputPer1M !== undefined) {
        const inputCost = (inUnits / 1000000) * Number(spec.inputPer1M);
        const outputCost = (outUnits / 1000000) * Number(spec.outputPer1M || 0);
        return Number((inputCost + outputCost).toFixed(6));
      }
      if (spec.costPerUnit !== undefined) {
        return Number((inUnits * Number(spec.costPerUnit)).toFixed(4));
      }
      return null;
    }

    // 1. Exact match on modelName
    if (rateKey && pricing[rateKey]) {
      const exactCost = computeFromSpec(pricing[rateKey]);
      if (exactCost !== null) return exactCost;
    }

    // 2. Case-insensitive & longest-prefix match on modelName (e.g. gpt-4o-2024-11-20 -> gpt-4o)
    if (rateKeyLow) {
      let bestKey = '';
      Object.keys(pricing).forEach(function(k) {
        const kLow = k.toLowerCase();
        if (rateKeyLow === kLow || rateKeyLow.indexOf(kLow) === 0) {
          if (k.length > bestKey.length) bestKey = k;
        }
      });
      if (bestKey) {
        const prefixCost = computeFromSpec(pricing[bestKey]);
        if (prefixCost !== null) return prefixCost;
      }
    }

    // 3. Exact or case-insensitive match on serviceName (e.g. Serper, ZeroBounce, MillionVerifier)
    if (serviceKey && pricing[serviceKey]) {
      const svcCost = computeFromSpec(pricing[serviceKey]);
      if (svcCost !== null) return svcCost;
    }
    const svcMatch = Object.keys(pricing).find(function(k) {
      return k.toLowerCase() === serviceKeyLow;
    });
    if (svcMatch) {
      const svcCaseCost = computeFromSpec(pricing[svcMatch]);
      if (svcCaseCost !== null) return svcCaseCost;
    }

    if (serviceKeyLow.indexOf('serper') >= 0 || rateKeyLow.indexOf('serper') >= 0) {
      const serperCostPerUnit = (pricing['Serper'] && pricing['Serper'].costPerUnit !== undefined)
        ? Number(pricing['Serper'].costPerUnit)
        : 0.001;
      return Number((inUnits * serperCostPerUnit).toFixed(4));
    }

    // 4. Conservative fallback for unmapped LLM models when tokens were consumed
    if (inUnits > 0 || outUnits > 0) {
      if (serviceKeyLow.indexOf('openai') >= 0 || rateKeyLow.indexOf('gpt') >= 0 || /^o\d/.test(rateKeyLow)) {
        return Number((((inUnits / 1000000) * 2.50) + ((outUnits / 1000000) * 10.00)).toFixed(6));
      }
      if (serviceKeyLow.indexOf('gemini') >= 0 || rateKeyLow.indexOf('gemini') >= 0) {
        return Number((((inUnits / 1000000) * 0.75) + ((outUnits / 1000000) * 3.75)).toFixed(6));
      }
    }

    return 0;
  } catch (e) {
    return 0;
  }
}

/**
 * Computes today's spend, daily cap, and remaining budget strictly scoped to one Client / Owner.
 */
function getClientBudgetSummary_(campaignId, explicitClientName, spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const targetClientName = resolveTrackerClientName_(ss, campaignId, explicitClientName);
  const targetClientKey = normalizeTrackerClientKey_(targetClientName);
  const isExternalClient = (targetClientKey !== 'GROWTHTARGET');
  const clientCapProp = isExternalClient ? (getTrackerClientPrefix_(targetClientName) + 'DAILY_SPEND_CAP_USD') : '';
  const normClientCapKey = clientCapProp.toUpperCase().replace(/[^A-Z0-9]/g, '');

  let defaultDailyCap = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails)
    ? Number(APP_CONFIG.guardrails.defaultDailySpendCapUsd)
    : 10.00;
  let clientDailyCap = null;
  let capSource = 'Default (Config)';

  if (ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const cleanKey = String(rows[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        const parsedVal = parseTrackerNumber_(rows[i][2]);
        if (isNaN(parsedVal) || parsedVal < 0) continue;

        if (cleanKey === 'DAILYSPENDCAPUSD') {
          defaultDailyCap = parsedVal;
          if (!isExternalClient) capSource = 'Agent Settings (DAILY_SPEND_CAP_USD)';
        } else if (isExternalClient && cleanKey === normClientCapKey) {
          clientDailyCap = parsedVal;
          capSource = 'Agent Settings (' + clientCapProp + ')';
        }
      }
    }
  }

  if (isExternalClient && clientDailyCap === null && clientCapProp) {
    const propVal = PropertiesService.getScriptProperties().getProperty(clientCapProp);
    const parsedProp = parseTrackerNumber_(propVal);
    if (!isNaN(parsedProp) && parsedProp >= 0) {
      clientDailyCap = parsedProp;
      capSource = 'Script Properties (' + clientCapProp + ')';
    }
  } else if (!isExternalClient && capSource === 'Default (Config)') {
    const localPropVal = PropertiesService.getScriptProperties().getProperty('DAILY_SPEND_CAP_USD');
    const parsedLocal = parseTrackerNumber_(localPropVal);
    if (!isNaN(parsedLocal) && parsedLocal >= 0) {
      defaultDailyCap = parsedLocal;
      capSource = 'Script Properties (DAILY_SPEND_CAP_USD)';
    }
  }

  const dailyCap = (isExternalClient && clientDailyCap !== null) ? clientDailyCap : defaultDailyCap;

  let todaySpend = 0;
  let callsToday = 0;
  const tz = (ss && ss.getSpreadsheetTimeZone) ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  const todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const logSheet = ss ? ss.getSheetByName('API_Usage_Log') : null;
  if (logSheet && logSheet.getLastRow() > 1) {
    const campaignMap = buildCampaignToClientKeyMap_(ss);
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();

    for (let r = 0; r < data.length; r++) {
      const rawDate = data[r][0];
      if (!rawDate) continue;
      const rowDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(rowDate.getTime())) continue;

      const rowIso = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
      if (rowIso !== todayIso) continue;

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
        callsToday++;
        const cost = parseFloat(data[r][7]) || 0;
        todaySpend += cost;
      }
    }
  }

  const remainingUsd = Math.max(0, dailyCap - todaySpend);
  const isOverCap = todaySpend >= dailyCap;

  return {
    clientName: targetClientName,
    isExternalClient: isExternalClient,
    dailyCap: Number(dailyCap.toFixed(2)),
    todaySpend: Number(todaySpend.toFixed(6)),
    remainingUsd: Number(remainingUsd.toFixed(4)),
    callsToday: callsToday,
    capSource: capSource,
    isOverCap: isOverCap
  };
}

/**
 * Pre-Flight Guardrail: Checks today's API spend scoped strictly to the target Client / Owner.
 * Enforces cap even when API_Usage_Log is empty (e.g. $0.00 pause cap).
 */
function checkApiBudgetGuardrails_(campaignId, explicitClientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  const summary = getClientBudgetSummary_(campaignId, explicitClientName, ss);
  Logger.log(
    '>>> Pre-Flight Budget Check [' + summary.clientName + ']: Today Spend = $' +
    summary.todaySpend.toFixed(4) + ' / Cap = $' + summary.dailyCap.toFixed(2) +
    ' (' + summary.capSource + ')'
  );

  if (summary.isOverCap) {
    throw new Error(
      '🛑 RUNAWAY LOOP ABORTED: Daily API Spend Hard Cap ($' + summary.dailyCap.toFixed(2) + ') reached for client "' + summary.clientName + '"! ' +
      'Today\'s spend for this client is $' + summary.todaySpend.toFixed(4) + '. Adjust cap in Agent Settings or wait until tomorrow.'
    );
  }
}

/**
 * Reads the configured minimum verifier credits threshold for a client.
 */
function getClientMinVerifierCredits_(ss, clientName, campaignId) {
  const targetClientName = resolveTrackerClientName_(ss, campaignId, clientName);
  const clientKey = normalizeTrackerClientKey_(targetClientName);
  const isExternalClient = (clientKey !== 'GROWTHTARGET');
  const clientMinProp = isExternalClient ? (getTrackerClientPrefix_(targetClientName) + 'MIN_VERIFIER_CREDITS') : '';
  const normClientMinKey = clientMinProp.toUpperCase().replace(/[^A-Z0-9]/g, '');

  let defaultMinCredits = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails)
    ? APP_CONFIG.guardrails.defaultMinVerifierCredits
    : 500;
  let clientMinCredits = null;
  let verifierProviderFromAgent = '';

  if (ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const cleanK = String(rows[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        const parsedVal = parseTrackerNumber_(rows[i][2]);

        if (cleanK === 'MINVERIFIERCREDITS' && !isNaN(parsedVal) && parsedVal >= 0) {
          defaultMinCredits = Math.round(parsedVal);
        } else if (isExternalClient && cleanK === normClientMinKey && !isNaN(parsedVal) && parsedVal >= 0) {
          clientMinCredits = Math.round(parsedVal);
        } else if (cleanK === 'VERIFIERPROVIDER') {
          const val = String(rows[i][2] || '').trim();
          if (val) verifierProviderFromAgent = val;
        }
      }
    }
  }

  if (isExternalClient && clientMinCredits === null && clientMinProp) {
    const propVal = PropertiesService.getScriptProperties().getProperty(clientMinProp);
    const parsedProp = parseTrackerNumber_(propVal);
    if (!isNaN(parsedProp) && parsedProp >= 0) {
      clientMinCredits = Math.round(parsedProp);
    }
  }

  return {
    minCredits: (isExternalClient && clientMinCredits !== null) ? clientMinCredits : defaultMinCredits,
    verifierProviderFromAgent: verifierProviderFromAgent
  };
}

/**
 * Performs a $0.00 read-only credit balance check for ZeroBounce, MillionVerifier, or NeverBounce.
 */
function fetchVerifierCreditBalance_(provider, apiKey) {
  const p = String(provider || '').trim().toLowerCase();
  if (!apiKey) {
    return { supported: false, ok: false, credits: null, unlimited: false, error: 'Missing API key.' };
  }

  if (p.indexOf('zerobounce') >= 0) {
    const url = 'https://api.zerobounce.net/v2/getcredits?api_key=' + encodeURIComponent(apiKey);
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = res.getResponseCode();
    const text = res.getContentText() || '';
    if (code < 200 || code >= 300) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: 'HTTP ' + code + ': ' + text.slice(0, 140) };
    }
    const json = JSON.parse(text || '{}');
    if (json.error) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: String(json.error) };
    }
    const credits = parseInt(json.Credits, 10);
    if (isNaN(credits) || credits < 0) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: 'Invalid ZeroBounce API key (Credits: ' + json.Credits + ').' };
    }
    return { supported: true, ok: true, credits: credits, unlimited: false, error: '' };
  }

  if (p.indexOf('millionverifier') >= 0) {
    const url = 'https://api.millionverifier.com/api/v3/credits?api=' + encodeURIComponent(apiKey);
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = res.getResponseCode();
    const text = res.getContentText() || '';
    if (code < 200 || code >= 300) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: 'HTTP ' + code + ': ' + text.slice(0, 140) };
    }
    const json = JSON.parse(text || '{}');
    if (json.error) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: String(json.error) };
    }
    const rawCredits = json.credits !== undefined ? json.credits : json.Credits;
    const credits = parseInt(rawCredits, 10);
    if (isNaN(credits) || credits < 0) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: 'Unexpected MillionVerifier response: ' + text.slice(0, 140) };
    }
    return { supported: true, ok: true, credits: credits, unlimited: false, error: '' };
  }

  if (p.indexOf('neverbounce') >= 0) {
    const url = 'https://api.neverbounce.com/v4/account/info?key=' + encodeURIComponent(apiKey);
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = res.getResponseCode();
    const text = res.getContentText() || '';
    if (code < 200 || code >= 300) {
      return { supported: true, ok: false, credits: null, unlimited: false, error: 'HTTP ' + code + ': ' + text.slice(0, 140) };
    }
    const json = JSON.parse(text || '{}');
    if (String(json.status || '').toLowerCase() !== 'success') {
      return { supported: true, ok: false, credits: null, unlimited: false, error: json.message || text.slice(0, 140) };
    }
    const info = json.credits_info || {};
    const billingType = String(json.billing_type || info.billing_type || '').toLowerCase();
    if (billingType === 'unlimited') {
      return { supported: true, ok: true, credits: null, unlimited: true, error: '' };
    }
    const freeCreds = parseInt(info.free_credits_remaining || 0, 10) || 0;
    const paidCreds = parseInt(info.paid_credits_remaining || 0, 10) || 0;
    return { supported: true, ok: true, credits: freeCreds + paidCreds, unlimited: false, error: '' };
  }

  return { supported: false, ok: true, credits: null, unlimited: false, error: '' };
}

/**
 * Checks verifier balance strictly for the specified campaign/client,
 * honoring per-client CLIENT_{NAME}_MIN_VERIFIER_CREDITS from Agent Settings.
 */
function checkVerifierBalanceAndAlert_(campaignId, explicitClientName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
    if (!ss) return;

    const activeCampId = campaignId || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : '');
    const clientName = resolveTrackerClientName_(ss, activeCampId, explicitClientName);
    const isExternalClient = (normalizeTrackerClientKey_(clientName) !== 'GROWTHTARGET');
    const minInfo = getClientMinVerifierCredits_(ss, clientName, activeCampId);
    const minCredits = minInfo.minCredits;
    let verifierProvider = minInfo.verifierProviderFromAgent || '';

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

    const providerLabel = verifierProvider || 'ZeroBounce';
    const pLow = providerLabel.toLowerCase();
    if (pLow === 'none' || pLow.indexOf('custom') >= 0) return;

    const candidateKeys = [];
    if (pLow.indexOf('zerobounce') >= 0) candidateKeys.push('ZEROBOUNCE_API_KEY');
    else if (pLow.indexOf('millionverifier') >= 0) candidateKeys.push('MILLIONVERIFIER_API_KEY');
    else if (pLow.indexOf('neverbounce') >= 0) candidateKeys.push('NEVERBOUNCE_API_KEY');
    candidateKeys.push('VERIFIER_API_KEY', 'EMAIL_VERIFICATION_API_KEY');

    let apiKey = '';
    if (typeof resolveClientScopedSecret_ === 'function') {
      const resolved = resolveClientScopedSecret_(
        candidateKeys,
        { campaignId: activeCampId, clientName: clientName },
        { throwOnMissing: false, allowAgentSettingsForLocal: !isExternalClient, spreadsheet: ss }
      );
      apiKey = resolved.key || '';
    }
    if (!apiKey) return;

    const bal = fetchVerifierCreditBalance_(providerLabel, apiKey);
    if (!bal.supported || !bal.ok || bal.unlimited || bal.credits === null) return;

    if (bal.credits < minCredits) {
      // Prevent duplicate alert spam within 6 hours for the same client + verifier
      const alertCooldownKey = 'LOW_VERIFIER_ALERT_' + normalizeTrackerClientKey_(clientName) + '_' + pLow.replace(/[^a-z0-9]/g, '');
      try {
        if (CacheService.getScriptCache().get(alertCooldownKey)) {
          return;
        }
      } catch (e) {}

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
        const subject = '⚠️ Low ' + providerLabel + ' Credits Alert [' + clientName + '] (' + bal.credits + ' remaining)';
        const body = 'Hello,\n\nYour AI Prospecting Studio automated check detected that the ' + providerLabel + ' email verification balance for client "' + clientName + '" (Campaign: ' + (activeCampId || 'Active') + ') has dropped to ' + bal.credits + ' credits (Threshold: ' + minCredits + ').\n\nPlease top up this account to prevent enrichment runs from pausing or failing mid-batch.\n\n- AI Prospecting Studio System';
        MailApp.sendEmail(alertEmail, subject, body);
        try { CacheService.getScriptCache().put(alertCooldownKey, '1', 21600); } catch (e) {}
        Logger.log('>>> Low balance email alert sent to ' + alertEmail + ' for client ' + clientName + ' (' + providerLabel + ': ' + bal.credits + ' / Min: ' + minCredits + ')');
      }
    }
  } catch (err) {
    Logger.log('>>> Warning in checkVerifierBalanceAndAlert_: ' + err.toString());
  }
}

/**
 * Returns live today spend vs cap for the Control Panel meter (fast, zero external API calls).
 */
function getClientBudgetSnapshot(clientName, campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const summary = getClientBudgetSummary_(campaignId || '', clientName || '', ss);
  return Object.assign({ success: true }, summary);
}

/**
 * Performs a $0.00 read-only Connection & Budget Audit for the active Client & Campaign.
 * Never spends AI tokens, search queries, or verification credits, never sends email, and never returns raw secrets.
 */
function runClientConnectionAndBudgetAudit(auditInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    return { success: false, overallStatus: 'FAIL', error: 'Spreadsheet not accessible.' };
  }

  const input = (auditInput && typeof auditInput === 'object') ? auditInput : {};
  let campaignId = String(input.campaignId || '').trim();
  if (!campaignId || campaignId === '__NEW_CAMPAIGN__') {
    campaignId = (typeof getStudioActiveCampaignId_ === 'function') ? getStudioActiveCampaignId_() : '';
  }

  const clientName = resolveTrackerClientName_(ss, campaignId, input.clientName || '');
  const isLocal = (normalizeTrackerClientKey_(clientName) === 'GROWTHTARGET');
  const prefix = isLocal ? '' : getTrackerClientPrefix_(clientName);

  // Fill any omitted config fields from saved campaign settings
  let savedCfg = {};
  if (typeof getSidebarInitialData === 'function' && campaignId) {
    try {
      const sData = getSidebarInitialData();
      if (sData && sData.configsMap && sData.configsMap[campaignId]) {
        savedCfg = sData.configsMap[campaignId];
      }
    } catch (e) {}
  }

  const aiProvider = String(input.aiProvider || savedCfg.aiProvider || 'GEMINI').trim().toUpperCase();
  const aiModel = String(input.aiModel || savedCfg.aiModel || (aiProvider.indexOf('OPENAI') >= 0 ? 'gpt-4o' : 'gemini-3.8-flash')).trim();
  const emailVerifier = String(input.emailVerifier || savedCfg.emailVerifier || 'ZeroBounce').trim();
  const emailEndpoint = String(input.emailEndpoint !== undefined ? input.emailEndpoint : (savedCfg.emailEndpoint || '')).trim();
  const senderProvider = String(input.emailSenderProvider || savedCfg.emailSenderProvider || 'Workspace Gmail (Drafts Only)').trim();
  const senderEndpoint = String(input.emailSenderEndpoint !== undefined ? input.emailSenderEndpoint : (savedCfg.emailSenderEndpoint || '')).trim();
  const senderEmail = String(input.senderEmail !== undefined ? input.senderEmail : (savedCfg.senderEmail || '')).trim();

  // 1. Budget Check
  const budgetSummary = getClientBudgetSummary_(campaignId, clientName, ss);
  const budgetCheck = {
    status: budgetSummary.isOverCap ? 'FAIL' : (budgetSummary.dailyCap > 0 && budgetSummary.todaySpend >= budgetSummary.dailyCap * 0.8 ? 'WARN' : 'PASS'),
    dailyCap: budgetSummary.dailyCap,
    todaySpend: budgetSummary.todaySpend,
    remainingUsd: budgetSummary.remainingUsd,
    callsToday: budgetSummary.callsToday,
    capSource: budgetSummary.capSource,
    message: budgetSummary.isOverCap
      ? ('Daily spend cap reached ($' + budgetSummary.todaySpend.toFixed(4) + ' / $' + budgetSummary.dailyCap.toFixed(2) + ')')
      : ('$' + budgetSummary.todaySpend.toFixed(4) + ' spent today / $' + budgetSummary.dailyCap.toFixed(2) + ' cap ($' + budgetSummary.remainingUsd.toFixed(2) + ' left)')
  };

  // 2. AI Provider Check ($0.00 /models GET)
  const aiCheck = auditAiConnection_(ss, clientName, campaignId, isLocal, aiProvider, aiModel);

  // 3. Email Verifier Check ($0.00 credit GET)
  const verifierCheck = auditVerifierConnection_(ss, clientName, campaignId, isLocal, emailVerifier, emailEndpoint);

  // 4. Email Dispatch Channel Check (Read-only auth/ping)
  const senderCheck = auditSenderConnection_(ss, clientName, campaignId, isLocal, prefix, senderProvider, senderEndpoint, senderEmail);

  // 5. Optional Serper Search Check ($0.00 Script Property presence check only)
  const serperCheck = auditSerperConnection_(ss, clientName, campaignId, isLocal);

  const allStatuses = [budgetCheck.status, aiCheck.status, verifierCheck.status, senderCheck.status];
  let overallStatus = 'PASS';
  if (allStatuses.indexOf('FAIL') >= 0) overallStatus = 'FAIL';
  else if (allStatuses.indexOf('WARN') >= 0) overallStatus = 'WARN';

  return {
    success: overallStatus !== 'FAIL',
    overallStatus: overallStatus,
    clientName: clientName,
    campaignId: campaignId || '(Unsaved)',
    budget: budgetCheck,
    ai: aiCheck,
    verifier: verifierCheck,
    sender: senderCheck,
    serper: serperCheck,
    checkedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss')
  };
}

function auditSerperConnection_(ss, clientName, campaignId, isLocal) {
  const baseProp = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.serperApiKeyProperty)
    ? APP_CONFIG.serperApiKeyProperty
    : 'SERPER_API_KEY';
  const prefix = isLocal ? '' : getTrackerClientPrefix_(clientName);
  let resolved = { key: '', propertyUsed: '', expectedProperty: prefix + baseProp };

  if (typeof resolveClientScopedSecret_ === 'function') {
    resolved = resolveClientScopedSecret_(
      [baseProp],
      { campaignId: campaignId, clientName: clientName },
      { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: ss }
    );
  } else if (isLocal) {
    const val = PropertiesService.getScriptProperties().getProperty(baseProp);
    if (val) resolved = { key: val, propertyUsed: baseProp, expectedProperty: baseProp };
  }

  return {
    status: resolved.key ? 'PASS' : 'INFO',
    configured: Boolean(resolved.key),
    propertyUsed: resolved.propertyUsed || '',
    expectedProperty: resolved.expectedProperty || (prefix + baseProp),
    message: resolved.key
      ? ('Configured (' + resolved.propertyUsed + ') — Google Search LinkedIn/Leadership pre-search enabled ($0.00 property check).')
      : ('Optional: Add ' + (resolved.expectedProperty || (prefix + baseProp)) + ' in Script Properties for live Google Search persona verification.')
  };
}

function auditAiConnection_(ss, clientName, campaignId, isLocal, aiProvider, aiModel) {
  const isOpenAi = aiProvider.indexOf('OPENAI') >= 0;
  const baseKey = isOpenAi ? 'OPENAI_API_KEY' : 'GEMINI_API_KEY';
  let resolved = { key: '', propertyUsed: '', expectedProperty: (isLocal ? baseKey : getTrackerClientPrefix_(clientName) + baseKey) };

  if (typeof resolveClientScopedSecret_ === 'function') {
    resolved = resolveClientScopedSecret_(
      [baseKey],
      { campaignId: campaignId, clientName: clientName },
      { throwOnMissing: false, allowAgentSettingsForLocal: isLocal, spreadsheet: ss }
    );
  }

  if (!resolved.key) {
    return {
      status: 'FAIL',
      provider: isOpenAi ? 'OPENAI' : 'GEMINI',
      model: aiModel,
      propertyUsed: '',
      message: 'Missing ' + resolved.expectedProperty + ' in Script Properties.'
    };
  }

  try {
    if (isOpenAi) {
      const res = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
        method: 'get',
        headers: { Authorization: 'Bearer ' + resolved.key },
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code < 200 || code >= 300) {
        return {
          status: 'FAIL',
          provider: 'OPENAI',
          model: aiModel,
          propertyUsed: resolved.propertyUsed,
          message: 'OpenAI auth failed (HTTP ' + code + ') via ' + resolved.propertyUsed
        };
      }
      const json = JSON.parse(res.getContentText() || '{}');
      const models = Array.isArray(json.data) ? json.data.map(function(m) { return String(m.id || ''); }) : [];
      const hasModel = !aiModel || models.indexOf(aiModel) >= 0;
      return {
        status: hasModel ? 'PASS' : 'WARN',
        provider: 'OPENAI',
        model: aiModel,
        propertyUsed: resolved.propertyUsed,
        message: hasModel
          ? ('Connected (' + resolved.propertyUsed + ') — model "' + aiModel + '" available ($0.00 check)')
          : ('Key valid (' + resolved.propertyUsed + '), but model "' + aiModel + '" was not listed in /v1/models.')
      };
    }

    const gemUrl = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(resolved.key);
    const res = UrlFetchApp.fetch(gemUrl, { method: 'get', muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      let errMsg = 'HTTP ' + code;
      try {
        const errJson = JSON.parse(res.getContentText() || '{}');
        if (errJson.error && errJson.error.message) errMsg += ': ' + errJson.error.message;
      } catch (e) {}
      return {
        status: 'FAIL',
        provider: 'GEMINI',
        model: aiModel,
        propertyUsed: resolved.propertyUsed,
        message: 'Gemini auth failed (' + errMsg + ') via ' + resolved.propertyUsed
      };
    }
    const json = JSON.parse(res.getContentText() || '{}');
    const modelNames = Array.isArray(json.models)
      ? json.models.map(function(m) { return String(m.name || '').replace(/^models\//, ''); })
      : [];
    const hasModel = !aiModel || modelNames.indexOf(aiModel) >= 0;
    return {
      status: hasModel ? 'PASS' : 'WARN',
      provider: 'GEMINI',
      model: aiModel,
      propertyUsed: resolved.propertyUsed,
      message: hasModel
        ? ('Connected (' + resolved.propertyUsed + ') — model "' + aiModel + '" available ($0.00 check)')
        : ('Key valid (' + resolved.propertyUsed + '), but "' + aiModel + '" wasn\'t in list (fallback cascade ready).')
    };
  } catch (err) {
    return {
      status: 'FAIL',
      provider: isOpenAi ? 'OPENAI' : 'GEMINI',
      model: aiModel,
      propertyUsed: resolved.propertyUsed,
      message: 'Connection error: ' + err.message
    };
  }
}

function auditVerifierConnection_(ss, clientName, campaignId, isLocal, emailVerifier, emailEndpoint) {
  const provider = String(emailVerifier || 'ZeroBounce').trim();
  const pLow = provider.toLowerCase();

  if (pLow === 'none') {
    return {
      status: 'WARN',
      provider: 'None',
      credits: null,
      minCredits: 0,
      propertyUsed: '',
      message: 'Verification disabled (None) — emails will not be verified before queuing.'
    };
  }

  const candidateKeys = [];
  if (pLow.indexOf('zerobounce') >= 0) candidateKeys.push('ZEROBOUNCE_API_KEY');
  else if (pLow.indexOf('millionverifier') >= 0) candidateKeys.push('MILLIONVERIFIER_API_KEY');
  else if (pLow.indexOf('neverbounce') >= 0) candidateKeys.push('NEVERBOUNCE_API_KEY');
  else if (pLow.indexOf('custom') >= 0) candidateKeys.push('CUSTOM_VERIFIER_API_KEY');
  candidateKeys.push('VERIFIER_API_KEY', 'EMAIL_VERIFICATION_API_KEY');

  let resolved = { key: '', propertyUsed: '', expectedProperty: (isLocal ? candidateKeys[0] : getTrackerClientPrefix_(clientName) + candidateKeys[0]) };
  if (typeof resolveClientScopedSecret_ === 'function') {
    resolved = resolveClientScopedSecret_(
      candidateKeys,
      { campaignId: campaignId, clientName: clientName },
      { throwOnMissing: false, allowAgentSettingsForLocal: isLocal, spreadsheet: ss }
    );
  }

  const minInfo = getClientMinVerifierCredits_(ss, clientName, campaignId);
  const minCredits = minInfo.minCredits;

  if (!resolved.key) {
    return {
      status: isLocal ? 'WARN' : 'FAIL',
      provider: provider,
      credits: null,
      minCredits: minCredits,
      propertyUsed: '',
      message: isLocal
        ? ('No ' + provider + ' key configured — local runs pass emails as unverified.')
        : ('Missing ' + resolved.expectedProperty + ' (or _VERIFIER_API_KEY) in Script Properties.')
    };
  }

  if (pLow.indexOf('custom') >= 0) {
    const validEndpoint = Boolean(emailEndpoint && /^https:\/\//i.test(emailEndpoint));
    return {
      status: validEndpoint ? 'PASS' : 'FAIL',
      provider: provider,
      credits: null,
      minCredits: minCredits,
      propertyUsed: resolved.propertyUsed,
      message: validEndpoint
        ? ('Key found (' + resolved.propertyUsed + ') & HTTPS endpoint configured.')
        : 'Custom HTTPS verifier requires a valid https:// endpoint URL.'
    };
  }

  try {
    const bal = fetchVerifierCreditBalance_(provider, resolved.key);
    if (!bal.ok) {
      return {
        status: 'FAIL',
        provider: provider,
        credits: null,
        minCredits: minCredits,
        propertyUsed: resolved.propertyUsed,
        message: provider + ' check failed: ' + bal.error
      };
    }
    if (bal.unlimited) {
      return {
        status: 'PASS',
        provider: provider,
        credits: null,
        minCredits: minCredits,
        propertyUsed: resolved.propertyUsed,
        message: 'Connected (' + resolved.propertyUsed + ') — Unlimited plan'
      };
    }
    const low = (bal.credits !== null && bal.credits < minCredits);
    return {
      status: low ? 'WARN' : 'PASS',
      provider: provider,
      credits: bal.credits,
      minCredits: minCredits,
      propertyUsed: resolved.propertyUsed,
      message: (low ? 'Low credits: ' : 'Connected (' + resolved.propertyUsed + ') — ') +
        bal.credits + ' credit(s) remaining (Min alert: ' + minCredits + ')'
    };
  } catch (err) {
    return {
      status: 'FAIL',
      provider: provider,
      credits: null,
      minCredits: minCredits,
      propertyUsed: resolved.propertyUsed,
      message: provider + ' error: ' + err.message
    };
  }
}

function auditSenderConnection_(ss, clientName, campaignId, isLocal, prefix, senderProvider, senderEndpoint, senderEmail) {
  const provLow = String(senderProvider || '').toLowerCase();

  // 1. Microsoft 365
  if (provLow.indexOf('microsoft') >= 0 || provLow.indexOf('m365') >= 0 || provLow.indexOf('exchange') >= 0) {
    if (!senderEmail || senderEmail.indexOf('@') < 1) {
      return {
        status: 'FAIL',
        provider: senderProvider,
        mailbox: senderEmail,
        message: 'Missing Sender Email / Mailbox (Exchange UPN) in Section 5.'
      };
    }
    try {
      const tenantId = (typeof getSenderApiKey_ === 'function') ? getSenderApiKey_('MS_TENANT_ID', ss, { campaignId: campaignId, clientName: clientName }) : '';
      const clientId = (typeof getSenderApiKey_ === 'function') ? getSenderApiKey_('MS_CLIENT_ID', ss, { campaignId: campaignId, clientName: clientName }) : '';
      const clientSecret = (typeof getSenderApiKey_ === 'function') ? getSenderApiKey_('MS_CLIENT_SECRET', ss, { campaignId: campaignId, clientName: clientName }) : '';

      if (!tenantId || !clientId || !clientSecret) {
        return {
          status: 'FAIL',
          provider: senderProvider,
          mailbox: senderEmail,
          message: 'Missing ' + prefix + 'MS_TENANT_ID / _MS_CLIENT_ID / _MS_CLIENT_SECRET in Script Properties.'
        };
      }
      if (typeof getMicrosoftGraphAccessToken_ !== 'function') {
        return { status: 'WARN', provider: senderProvider, mailbox: senderEmail, message: 'Credentials found; EmailSender helper not loaded.' };
      }

      const token = getMicrosoftGraphAccessToken_(tenantId, clientId, clientSecret, clientName, {});
      const boxUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(senderEmail) + '/mailFolders/inbox?$select=id,displayName';
      const boxRes = UrlFetchApp.fetch(boxUrl, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      const boxCode = boxRes.getResponseCode();
      if (boxCode >= 200 && boxCode < 300) {
        return {
          status: 'PASS',
          provider: senderProvider,
          mailbox: senderEmail,
          message: 'OAuth2 token & Exchange mailbox "' + senderEmail + '" verified ($0.00 read check).'
        };
      }
      const errMsg = (typeof parseMicrosoftGraphError_ === 'function')
        ? parseMicrosoftGraphError_(boxCode, boxRes.getContentText(), 'verify mailbox', senderEmail)
        : ('Graph HTTP ' + boxCode);
      return {
        status: (boxCode === 403 && provLow.indexOf('direct') >= 0) ? 'WARN' : 'FAIL',
        provider: senderProvider,
        mailbox: senderEmail,
        message: (boxCode === 403 && provLow.indexOf('direct') >= 0)
          ? ('OAuth2 succeeded, but Mail.ReadWrite is not granted (only /sendMail may work; bounce/reply polling disabled).')
          : errMsg
      };
    } catch (err) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: err.message };
    }
  }

  // 2. Google Workspace Gmail (Local or Client Relay)
  if (provLow.indexOf('gmail') >= 0 || provLow.indexOf('workspace') >= 0 || provLow.indexOf('draft') >= 0 || provLow.indexOf('direct') >= 0) {
    if (isLocal) {
      try {
        if (typeof ScriptApp !== 'undefined' && typeof ScriptApp.getAuthorizationInfo === 'function') {
          const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL, ['https://mail.google.com/']);
          if (authInfo && authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
            return {
              status: 'FAIL',
              provider: senderProvider,
              mailbox: senderEmail || 'Local Workspace',
              message: 'Missing Gmail permission (https://mail.google.com/) for current Chrome account.'
            };
          }
        }
        let activeEmail = '';
        try { activeEmail = Session.getActiveUser().getEmail() || ''; } catch (e) {}
        return {
          status: 'PASS',
          provider: senderProvider,
          mailbox: senderEmail || activeEmail || 'Local Workspace Gmail',
          message: 'Local Gmail permissions verified' + (activeEmail ? (' (Account: ' + activeEmail + ')') : '.')
        };
      } catch (err) {
        return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: err.message };
      }
    }

    // External Client Workspace Relay
    try {
      const urlRes = resolveClientScopedSecret_(['GMAIL_RELAY_URL'], { campaignId: campaignId, clientName: clientName }, { throwOnMissing: false, spreadsheet: ss });
      const secRes = resolveClientScopedSecret_(['GMAIL_RELAY_SECRET'], { campaignId: campaignId, clientName: clientName }, { throwOnMissing: false, spreadsheet: ss });

      if (!urlRes.key || !secRes.key) {
        return {
          status: 'FAIL',
          provider: senderProvider,
          mailbox: senderEmail,
          message: 'Missing ' + prefix + 'GMAIL_RELAY_URL or ' + prefix + 'GMAIL_RELAY_SECRET in Script Properties.'
        };
      }
      if (!/^https:\/\//i.test(urlRes.key)) {
        return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: prefix + 'GMAIL_RELAY_URL must start with https://.' };
      }

      const res = UrlFetchApp.fetch(urlRes.key, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          secret: secRes.key,
          action: 'pollEvents',
          lookbackDays: 1,
          trackedDrafts: [],
          trackedRecipients: [],
          clientName: clientName
        }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      const text = res.getContentText() || '';
      if (code >= 300) {
        return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: 'Relay HTTP ' + code + ': ' + text.slice(0, 140) };
      }
      let parsed = {};
      try { parsed = JSON.parse(text); } catch (e) {
        return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: 'Relay returned non-JSON (check Web App "Anyone" access).' };
      }
      if (parsed && parsed.success === true) {
        return {
          status: 'PASS',
          provider: senderProvider,
          mailbox: senderEmail || 'Client Workspace Relay',
          message: 'Client Gmail Relay v2.2+ authenticated cleanly.'
        };
      }
      const errStr = String((parsed && parsed.error) || text).trim();
      if (errStr.toLowerCase().indexOf('unsupported') >= 0 || errStr.toLowerCase().indexOf('unknown action') >= 0) {
        return {
          status: 'WARN',
          provider: senderProvider,
          mailbox: senderEmail,
          message: 'Relay authenticated (v2.1). Upgrade ClientGmailRelay.gs to v2.2 for bounce/reply polling.'
        };
      }
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: 'Relay error: ' + errStr.slice(0, 160) };
    } catch (err) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEmail, message: err.message };
    }
  }

  // 3. Smartlead
  if (provLow.indexOf('smartlead') >= 0) {
    const keyRes = resolveClientScopedSecret_(['SMARTLEAD_API_KEY'], { campaignId: campaignId, clientName: clientName }, { throwOnMissing: false, allowAgentSettingsForLocal: isLocal, spreadsheet: ss });
    if (!keyRes.key) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEndpoint, message: 'Missing ' + keyRes.expectedProperty + ' in Script Properties.' };
    }
    if (!senderEndpoint) {
      return { status: 'FAIL', provider: senderProvider, mailbox: '', message: 'Smartlead Campaign ID is blank in Email Sender Endpoint.' };
    }
    try {
      const url = 'https://server.smartlead.ai/api/v1/campaigns/' + encodeURIComponent(senderEndpoint) + '?api_key=' + encodeURIComponent(keyRes.key);
      const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = res.getResponseCode();
      return {
        status: (code >= 200 && code < 300) ? 'PASS' : 'FAIL',
        provider: senderProvider,
        mailbox: senderEndpoint,
        message: (code >= 200 && code < 300)
          ? ('Smartlead campaign #' + senderEndpoint + ' verified (' + keyRes.propertyUsed + ').')
          : ('Smartlead HTTP ' + code + ': ' + res.getContentText().slice(0, 140))
      };
    } catch (err) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEndpoint, message: err.message };
    }
  }

  // 4. Instantly
  if (provLow.indexOf('instantly') >= 0) {
    const keyRes = resolveClientScopedSecret_(['INSTANTLY_API_KEY'], { campaignId: campaignId, clientName: clientName }, { throwOnMissing: false, allowAgentSettingsForLocal: isLocal, spreadsheet: ss });
    if (!keyRes.key) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEndpoint, message: 'Missing ' + keyRes.expectedProperty + ' in Script Properties.' };
    }
    if (!senderEndpoint) {
      return { status: 'FAIL', provider: senderProvider, mailbox: '', message: 'Instantly Campaign ID is blank in Email Sender Endpoint.' };
    }
    try {
      const url = 'https://api.instantly.ai/api/v1/campaign/get?api_key=' + encodeURIComponent(keyRes.key) + '&campaign_id=' + encodeURIComponent(senderEndpoint);
      const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = res.getResponseCode();
      return {
        status: (code >= 200 && code < 300) ? 'PASS' : 'FAIL',
        provider: senderProvider,
        mailbox: senderEndpoint,
        message: (code >= 200 && code < 300)
          ? ('Instantly campaign verified (' + keyRes.propertyUsed + ').')
          : ('Instantly HTTP ' + code + ': ' + res.getContentText().slice(0, 140))
      };
    } catch (err) {
      return { status: 'FAIL', provider: senderProvider, mailbox: senderEndpoint, message: err.message };
    }
  }

  // 5. Custom Webhook
  if (provLow.indexOf('webhook') >= 0 || provLow.indexOf('custom') >= 0) {
    let webhookUrl = senderEndpoint;
    let propLabel = 'Campaign Endpoint';
    if (!webhookUrl) {
      const urlRes = resolveClientScopedSecret_(['WEBHOOK_URL', 'EMAIL_SENDER_WEBHOOK_URL'], { campaignId: campaignId, clientName: clientName }, { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: ss });
      webhookUrl = urlRes.key || '';
      propLabel = urlRes.propertyUsed || (prefix + 'WEBHOOK_URL');
    }
    if (!webhookUrl || !/^https:\/\//i.test(webhookUrl)) {
      return { status: 'FAIL', provider: senderProvider, mailbox: '', message: 'Valid https:// Webhook URL missing in Section 5 or ' + prefix + 'WEBHOOK_URL.' };
    }
    return {
      status: 'PASS',
      provider: senderProvider,
      mailbox: webhookUrl,
      message: 'HTTPS Webhook URL configured (' + propLabel + '; dry-run skips live POST).'
    };
  }

  return { status: 'WARN', provider: senderProvider, mailbox: senderEmail, message: 'Provider "' + senderProvider + '" has no automatic dry-run test.' };
}