/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * CENTRALIZED API USAGE TRACKER & BUDGET GUARDRAIL ENGINE
 * ============================================================================
 * File: ApiTracker.gs / ApiTracker.js
 * 
 * PURPOSE:
 * 1. Logs every external API call (LLMs, Verifiers) to "API_Usage_Log".
 * 2. Automatically calculates est. dollar cost based on token counts or unit pricing.
 * 3. Enforces pre-flight daily spend hard caps to prevent runaway loops.
 * 4. Dispatches automated low-balance email alerts via MailApp.
 * ============================================================================
 */

/**
 * Centralized logging function for any external API request.
 * @param {Object} data
 * @param {string} data.campaignId - Active campaign ID
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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    const serviceName = data.serviceName || 'Unknown';
    const operation = data.operation || 'API Call';
    const httpStatus = data.httpStatus || 200;
    const inputUnits = Number(data.inputUnits) || 0;
    const outputUnits = Number(data.outputUnits) || 0;
    const status = data.status || 'SUCCESS';
    const details = String(data.details || '');
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
 * Pre-Flight Guardrail: Checks today's total API spend against daily hard cap.
 * Throws an exception to abort execution if budget exceeded.
 */
function checkApiBudgetGuardrails_(campaignId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('API_Usage_Log');
    if (!sheet || sheet.getLastRow() <= 1) return;

    let dailyCap = 10.00;
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim().toUpperCase() === 'DAILY_SPEND_CAP_USD') {
          const val = parseFloat(rows[i][2]);
          if (!isNaN(val)) dailyCap = val;
          break;
        }
      }
    } else if (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.guardrails) {
      dailyCap = APP_CONFIG.guardrails.defaultDailySpendCapUsd;
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const todayStr = new Date().toDateString();
    let todaySpend = 0;

    for (let r = 0; r < data.length; r++) {
      const rowDate = new Date(data[r][0]);
      if (rowDate.toDateString() === todayStr) {
        const cost = parseFloat(data[r][7]) || 0;
        todaySpend += cost;
      }
    }

    Logger.log('>>> Pre-Flight Budget Check: Today Spend = $' + todaySpend.toFixed(4) + ' / Cap = $' + dailyCap.toFixed(2));

    if (todaySpend >= dailyCap) {
      throw new Error('🛑 RUNAWAY LOOP ABORTED: Daily API Spend Hard Cap ($' + dailyCap.toFixed(2) + ') reached! Today\'s spend is currently $' + todaySpend.toFixed(4) + '. Adjust cap in Agent Settings or wait until tomorrow.');
    }

  } catch (err) {
    if (err.toString().includes('RUNAWAY LOOP ABORTED')) {
      throw err;
    }
    Logger.log('>>> Warning in checkApiBudgetGuardrails_: ' + err.toString());
  }
}

/**
 * Checks email verifier balance (e.g., ZeroBounce) and sends email alert if below threshold.
 */
function checkVerifierBalanceAndAlert_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let minCredits = 500;
    let verifierProvider = 'ZeroBounce';
    const agentSheet = ss.getSheetByName('Agent Settings');
    
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const k = String(rows[i][1]).trim().toUpperCase();
        if (k === 'MIN_VERIFIER_CREDITS') {
          const val = parseInt(rows[i][2], 10);
          if (!isNaN(val)) minCredits = val;
        } else if (k === 'VERIFIER_PROVIDER') {
          const val = String(rows[i][2]).trim();
          if (val) verifierProvider = val;
        }
      }
    }

    if (verifierProvider.toLowerCase() !== 'zerobounce') {
      return;
    }

    const scriptProps = PropertiesService.getScriptProperties().getProperties();
    let apiKey = '';
    for (let k in scriptProps) {
      if (k.toUpperCase().includes('VERIFIER') || k.toUpperCase().includes('EMAIL_VERIFICATION')) {
        const val = String(scriptProps[k] || '').trim();
        if (val && !val.includes('PASTE_')) {
          apiKey = val;
          break;
        }
      }
    }

    if (!apiKey) return;

    const url = 'https://api.zerobounce.net/v2/getcredits?api_key=' + encodeURIComponent(apiKey);
    const options = { method: 'get', muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const json = JSON.parse(response.getContentText());
      const credits = parseInt(json.Credits, 10);
      
      if (!isNaN(credits) && credits < minCredits) {
        const userEmail = Session.getActiveUser().getEmail();
        if (userEmail) {
          const subject = '⚠️ Low Email Verification Credits Alert (' + credits + ' remaining)';
          const body = 'Hello,\n\nYour AI Prospecting Studio automated check detected that your ZeroBounce email verification balance has dropped to ' + credits + ' credits (Threshold: ' + minCredits + ').\n\nPlease top up your account to prevent enrichment runs from pausing or failing mid-batch.\n\n- AI Prospecting Studio System';
          MailApp.sendEmail(userEmail, subject, body);
          Logger.log('>>> Low balance email alert sent to ' + userEmail + ' (Credits: ' + credits + ')');
        }
      }
    }
  } catch (err) {
    Logger.log('>>> Warning in checkVerifierBalanceAndAlert_: ' + err.toString());
  }
}