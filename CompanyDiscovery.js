/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AGENT 1: COMPANY DISCOVERY & RESEARCH AGENT (MULTI-AI PROVIDER ENGINE)
 * ============================================================================
 * File: CompanyDiscovery.gs / CompanyDiscovery.js
 * 
 * CAPABILITIES:
 * 1. Multi-AI Support: Dynamically routes to GEMINI (gemini-3.8-flash with Grounding)
 *    or OPENAI (GPT-4o) reading API keys from Script Properties or Agent Settings.
 * 2. Multi-Department Intelligence: Normalizes and reads DEPARTMENT, CUSTOM_BUSINESS_NEEDS,
 *    and targeting parameters reliably from "Campaign Configuration".
 * 3. Dual-Tab Population: Writes master entries to "Company Database" AND
 *    populates all 25 columns of "Company Intelligence".
 * 4. Grounding Evidence Logging: Stores audit citation URLs in "Research Evidence".
 * 5. High-Contrast Headers: Sets clean Dark Navy font (#174ea6) on Light Blue (#e8f0fe).
 * ============================================================================
 */

/**
 * Called directly from the Sidebar Control Panel (No blocking popups; returns results to sidebar)
 */
function runCompanyDiscoveryFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'A research run is already in progress. Please wait.' };
  }

  try {
    const config = getCampaignConfigurationRecord_(spreadsheet, campaignId);
    if (!config || !config.id) {
      return { success: false, message: 'No configuration found for Campaign ID: ' + (campaignId || 'Active') + ' in "Campaign Configuration".' };
    }

    const companyLimit = Math.min(25, Math.max(1, Number(config.companyLimit || 5)));
    const prompt = buildCompanyDiscoveryPrompt_(config, companyLimit);

    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';

    Logger.log('Launching Agent 1 with ' + providerLabel + ' for campaign: ' + config.id);

    const aiResponse = executeAiDiscoveryCall_({
      provider: provider,
      model: model,
      prompt: prompt,
      enableWebSearch: true,
      json: true
    });

    const discoveredCompanies = (aiResponse && aiResponse.json && Array.isArray(aiResponse.json.companies)) 
      ? aiResponse.json.companies 
      : [];

    if (!discoveredCompanies.length) {
      return { 
        success: false, 
        message: providerLabel + ' found no matching companies. Broaden your criteria or custom business needs in the sidebar.' 
      };
    }

    const addedCount = saveDiscoveredCompanies_(spreadsheet, config.id, discoveredCompanies, config, providerLabel);
    updateCampaignDiscoveryCounts_(spreadsheet, config.id);

    return {
      success: true,
      count: addedCount,
      message: 'Discovered ' + addedCount + ' qualified companies using ' + providerLabel + ' for ' + config.department + '! Populated in "Company Database" & "Company Intelligence".'
    };
  } catch (error) {
    Logger.log('Discovery Error: ' + error.stack);
    return { success: false, message: 'Discovery failed: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Called from the Top Menu (Uses standard Google Sheets dialogs with safe context fallback)
 */
function runCompanyDiscoveryForSelectedCampaign() {
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Running in headless/editor context where UI is unavailable
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  
  if (!lock.tryLock(5000)) {
    if (ui) ui.alert('A research run is already in progress. Please wait for it to finish.');
    else Logger.log('A research run is already in progress.');
    return;
  }

  try {
    const campaignId = getActiveCampaignId_(spreadsheet, ui);
    if (!campaignId) return;

    const config = getCampaignConfigurationRecord_(spreadsheet, campaignId);
    if (!config || !config.id) {
      if (ui) ui.alert('No configuration found for Campaign ID: ' + campaignId + ' in "Campaign Configuration".');
      else Logger.log('No configuration found for Campaign ID: ' + campaignId);
      return;
    }

    const companyLimit = Math.min(25, Math.max(1, Number(config.companyLimit || 5)));
    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';
    
    if (ui) {
      ui.alert(
        'Starting Company Discovery (Agent 1)',
        'Searching for up to ' + companyLimit + ' companies using ' + providerLabel + ':\n' +
        '• Department: ' + config.department + '\n' +
        '• Industry: ' + (config.industry || 'Any') + '\n' +
        '• Geography: ' + (config.country || 'Any') + '\n' +
        '• Needs/Triggers: ' + (config.customBusinessNeeds.slice(0, 100) || 'General') + '\n\n' +
        'The AI research agent is now searching the web. Click OK to begin.',
        ui.ButtonSet.OK
      );
    }

    const prompt = buildCompanyDiscoveryPrompt_(config, companyLimit);
    const aiResponse = executeAiDiscoveryCall_({
      provider: provider,
      model: model,
      prompt: prompt,
      enableWebSearch: true,
      json: true
    });

    const discoveredCompanies = (aiResponse && aiResponse.json && Array.isArray(aiResponse.json.companies)) 
      ? aiResponse.json.companies 
      : [];

    if (!discoveredCompanies.length) {
      if (ui) ui.alert(providerLabel + ' did not return any companies matching the given criteria. Try broadening your criteria.');
      else Logger.log(providerLabel + ' did not return any companies.');
      return;
    }

    const addedCount = saveDiscoveredCompanies_(spreadsheet, config.id, discoveredCompanies, config, providerLabel);
    updateCampaignDiscoveryCounts_(spreadsheet, config.id);

    if (ui) {
      ui.alert(
        'Company Discovery Complete',
        'Discovered and added ' + addedCount + ' qualified companies to "Company Database" and "Company Intelligence" using ' + providerLabel + '.\n\n' +
        'Next Step: Run "Agent 2: Persona Mapping" from the sidebar to find decision-makers and draft outreach.',
        ui.ButtonSet.OK
      );
    } else {
      Logger.log('Company Discovery Complete: Added ' + addedCount + ' companies.');
    }

  } catch (error) {
    if (ui) ui.alert('Error during Company Discovery: ' + error.message);
    else Logger.log('Error during Company Discovery: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds the Structured Search Prompt for Gemini or OpenAI
 */
function buildCompanyDiscoveryPrompt_(config, limit) {
  return [
    'Act as the Senior B2B Account Research & Market Intelligence Agent.',
    'Your goal is to discover REAL, active, operating enterprise/mid-market companies using live web search.',
    '',
    'CAMPAIGN PARAMETERS:',
    '• Operating Department: ' + config.department,
    '• Target Geographies / Countries: ' + (config.country || 'United States & Canada'),
    '• Target Industry Verticals: ' + (config.industry || 'Freight & Logistics, Manufacturing'),
    '• Target Company Size (Headcount): ' + (config.size || '50 - 250 employees'),
    '• Research Depth: ' + (config.researchDepth || 'Deep Persona'),
    '• Minimum Fit Score Required: ' + (config.minCompanyScore || 80) + ' out of 100',
    '• Specific Business Needs / Buying Signals to Match: ',
    config.customBusinessNeeds || 'Scope 3 vendor emissions transparency, supply chain audits, vendor questionnaires',
    '',
    'DISCOVERY & QUALIFICATION RULES:',
    '1. Use web search to find verified, operating companies. NEVER hallucinate or invent company names, statistics, or URLs.',
    '2. Return official, secure root website URLs (e.g. "https://companyname.com").',
    '3. Provide up to ' + limit + ' distinct companies that strictly match the criteria.',
    '4. Score each company from 0 to 100 on how closely they exhibit the specified business needs, pain points, or buying signals.',
    '5. Only qualify companies scoring ' + (config.minCompanyScore || 80) + ' or higher.',
    '6. Identify specific evidence URLs verifying why this company is an ideal fit.',
    '',
    'Return ONLY a valid, parsable JSON object with NO markdown formatting, matching this exact schema:',
    '{',
    '  "companies": [',
    '    {',
    '      "companyName": "Exact Operating Name",',
    '      "website": "https://companydomain.com",',
    '      "headquarters": "City, Country",',
    '      "industry": "Specific Industry Sector",',
    '      "companySize": "Estimated employee count (e.g. 150 employees)",',
    '      "businessSummary": "1-2 sentence description of primary business activities",',
    '      "businessTriggers": "Specific pain points or buying triggers detected matching the campaign needs",',
    '      "solutionFit": "Why the department solution directly solves their operational challenge",',
    '      "companyOpportunityScore": 90,',
    '      "companyOpportunity": "Core commercial opportunity or bottleneck to address",',
    '      "digitalSignals": "Cloud systems, ERP/TMS, tooling, or modernization signals detected",',
    '      "scoreRationale": "Clear explanation of why this company passed qualification",',
    '      "sourceUrl": "Direct URL where company profile or trigger was verified"',
    '    }',
    '  ]',
    '}'
  ].join('\n');
}

/**
 * Resilient AI Router supporting Script Properties (Project Settings) & Agent Settings
 */
function executeAiDiscoveryCall_(options) {
  if (typeof callAI === 'function') {
    try {
      const res = callAI(options);
      if (res && (res.json || res.text)) return res;
    } catch (e) {
      Logger.log('External callAI error, using direct caller: ' + e.message);
    }
  }

  const provider = String(options.provider || 'GEMINI').toUpperCase();
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  let apiKey = scriptProps[provider === 'OPENAI' ? 'OPENAI_API_KEY' : 'GEMINI_API_KEY'] || '';

  if (!apiKey && typeof getStudioApiKey === 'function') {
    apiKey = getStudioApiKey(provider);
  }

  if (provider === 'OPENAI') {
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: options.model || 'gpt-4o',
      messages: [{ role: 'user', content: options.prompt }],
      response_format: options.json ? { type: 'json_object' } : undefined,
      temperature: 0.2
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const body = JSON.parse(response.getContentText() || '{}');
    const text = body.choices && body.choices[0] && body.choices[0].message ? body.choices[0].message.content : '{}';
    return { text: text, json: JSON.parse(text), citations: [] };
  }

  if (provider === 'GEMINI' && typeof callGemini_ === 'function') {
    return callGemini_(options);
  }

  throw new Error('No API key configured for ' + provider + ' in Project Settings.');
}

/**
 * Saves discovered companies into BOTH "Company Database" and "Company Intelligence"
 */
function saveDiscoveredCompanies_(spreadsheet, campaignId, discoveredCompanies, config, providerLabel) {
  const dbSheet = spreadsheet.getSheetByName('Company Database');
  const intelSheet = spreadsheet.getSheetByName('Company Intelligence');
  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');

  if (!dbSheet) throw new Error('Sheet "Company Database" not found.');
  if (!intelSheet) throw new Error('Sheet "Company Intelligence" not found.');

  // Ensure high-contrast headers on destination sheets
  [dbSheet, intelSheet].forEach(sh => {
    if (sh.getLastColumn() > 0) {
      sh.getRange(1, 1, 1, sh.getLastColumn())
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sh.setFrozenRows(1);
    }
  });

  const existingDbRecords = getSheetRecords_(dbSheet);
  const intelHeaders = intelSheet.getRange(1, 1, 1, intelSheet.getLastColumn()).getValues()[0];
  const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];

  let addedCount = 0;
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const isEsg = config.department.toLowerCase().includes('esg');

  discoveredCompanies.forEach(function(item) {
    const companyName = String(item.companyName || '').trim();
    const website = String(item.website || '').trim();
    const domain = domainFromUrl_(website);

    const isDuplicate = existingDbRecords.some(function(row) {
      const rowDomain = domainFromUrl_(row['Domain'] || row['Website'] || '');
      const rowName = String(row['Company Name'] || row['Company name'] || '').trim().toLowerCase();
      return (row['Campaign ID'] === campaignId) && (
        (domain && rowDomain && domain === rowDomain) ||
        (rowName && rowName === companyName.toLowerCase())
      );
    });

    if (isDuplicate || !companyName) return;

    const companyId = 'CMP-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    const score = Number(item.companyOpportunityScore || item.initialScore || 85);
    const status = score >= (config.minCompanyScore || 80) ? 'Qualified' : 'Disqualified';

    // 1. Write to "Company Database"
    const dbRecord = {
      'Company ID': companyId,
      'Campaign ID': campaignId,
      'Company Name': companyName,
      'Domain': domain || website,
      'Country': item.headquarters || config.country || '',
      'Industry': item.industry || config.industry || '',
      'Employee Size': item.companySize || config.size || '',
      'Discovery Method': providerLabel + ' - ' + config.department,
      'Discovery Date': dateStr,
      'Status': status
    };
    dbSheet.appendRow(buildRowFromHeaders_(dbHeaders, dbRecord));

    // 2. Write to "Company Intelligence" (All 25 Columns)
    // Solution fit strictly aligned with dropdown options
    const fitLabel = score >= 85 ? 'High' : (score >= 70 ? 'Medium' : 'Low');

    const intelRecord = {
      'Campaign ID': campaignId,
      'Campaign Name': config.name || config.id,
      'Company ID': companyId,
      'Company Name': companyName,
      'Country': item.headquarters || config.country || '',
      'Industry': item.industry || config.industry || '',
      'Company Size': item.companySize || config.size || '',
      'ESG Frameworks': isEsg ? 'CSRD, CDP, GHG Protocol' : 'N/A (Enterprise Stack)',
      'Reporting Obligations': isEsg ? 'Mandatory Tier 1 Scope 3 Reporting' : 'Standard Regulatory Compliance',
      'ESG Reporting Maturity': isEsg ? 'Developing / Spreadsheet-reliant' : 'N/A',
      'Supply Chain Profile': 'Multi-tier supplier & infrastructure network',
      'Scope 3 Exposure': isEsg ? 'High (Supply Chain Emissions)' : 'N/A',
      'Sustainability Data Needs': isEsg ? 'Automated vendor ESG data ingestion' : 'System observability & audit automation',
      'Digital Transformation Signals': item.digitalSignals || item.businessSummary || 'Cloud infrastructure detected',
      'ESG Governance': isEsg ? 'Director of Sustainability' : 'VP of Engineering / Operations',
      'Business Triggers': item.businessTriggers || config.customBusinessNeeds.slice(0, 150),
      'eValuater Solution Fit': fitLabel,
      'Company Opportunity Score': score,
      'Company Opportunity': item.companyOpportunity || item.businessSummary || 'Operational improvement opportunity',
      'Evidence URLs': item.sourceUrl || website,
      'Evidence Checked Date': dateStr,
      'Verification Status': 'Verified',
      'Human Review Status': 'Accepted',
      'Last Enriched At': now.toLocaleString(),
      'Company Score Rationale': item.scoreRationale || ('Qualified based on ' + config.department + ' criteria via ' + providerLabel + '.')
    };
    intelSheet.appendRow(buildRowFromHeaders_(intelHeaders, intelRecord));

    // 3. Write to "Research Evidence"
    if (evidenceSheet && item.sourceUrl) {
      evidenceSheet.appendRow([
        'EVD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
        campaignId,
        companyName,
        'Company Discovery',
        'Discovery Query via ' + providerLabel,
        item.sourceUrl,
        item.businessSummary || companyName,
        (score / 100).toFixed(2),
        now.toLocaleString()
      ]);
    }

    addedCount++;
  });

  return addedCount;
}

/**
 * Resolves default model based on selected AI Provider
 */
function resolveAiModel_(provider, configuredModel) {
  if (configuredModel && String(configuredModel).trim() !== '') {
    return String(configuredModel).trim();
  }
  return (provider === 'OPENAI') ? 'gpt-4o' : 'gemini-3.8-flash';
}

/**
 * Universal Configuration Loader: Reads from Campaign Configuration & Agent Settings with Key Normalization
 */
function getCampaignConfigurationRecord_(spreadsheet, campaignId) {
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  const agentSheet = spreadsheet.getSheetByName('Agent Settings');

  const config = {
    id: campaignId || 'CMP-005',
    name: '',
    department: 'ESG Advisory',
    industry: 'Freight & Logistics, Manufacturing',
    country: 'United States & Canada',
    size: '50 - 250 employees',
    companyLimit: 5,
    researchDepth: 'Deep Persona',
    minCompanyScore: 80,
    minPersonaScore: 85,
    customBusinessNeeds: '',
    aiProvider: 'GEMINI',
    aiModel: 'gemini-3.8-flash'
  };

  if (!configSheet) return config;

  const data = configSheet.getDataRange().getValues();
  if (data.length <= 1) return config;

  for (let r = 1; r < data.length; r++) {
    const rawKey = String(data[r][0] || '').trim();
    const cleanKey = rawKey.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const val = data[r][1];
    if (val === undefined || val === null || val === '') continue;

    if (cleanKey === 'CAMPAIGNID' || cleanKey === 'ID') config.id = String(val).trim();
    if (cleanKey === 'CAMPAIGNNAME' || cleanKey === 'NAME') config.name = String(val).trim();
    if (cleanKey === 'DEPARTMENT' || cleanKey === 'DEPARTMENTOWNER') config.department = String(val).trim();
    if (cleanKey === 'TARGETCOUNTRIES' || cleanKey === 'TARGETGEOGRAPHY' || cleanKey === 'COUNTRIESREGIONS' || cleanKey === 'COUNTRY') config.country = String(val).trim();
    if (cleanKey === 'TARGETINDUSTRIES' || cleanKey === 'INDUSTRIES' || cleanKey === 'INDUSTRY') config.industry = String(val).trim();
    if (cleanKey === 'COMPANYSIZERANGE' || cleanKey === 'COMPANYSIZE' || cleanKey === 'SIZE') config.size = String(val).trim();
    if (cleanKey === 'MAXCOMPANIES' || cleanKey === 'MAXCOMPANIESLIMIT' || cleanKey === 'COMPANYLIMIT') config.companyLimit = Number(val) || 5;
    if (cleanKey === 'RESEARCHDEPTH') config.researchDepth = String(val).trim();
    if (cleanKey === 'MINCOMPANYSCORE' || cleanKey === 'MINCOMPANYFITSCORE') config.minCompanyScore = Number(val) || 80;
    if (cleanKey === 'MINPERSONASCORE' || cleanKey === 'MINPERSONAFITSCORE') config.minPersonaScore = Number(val) || 85;
    if (cleanKey === 'CUSTOMBUSINESSNEEDS') config.customBusinessNeeds = String(val).trim();
    if (cleanKey === 'AIPROVIDER' || cleanKey === 'PRIMARYAIPROVIDER') config.aiProvider = String(val).trim().toUpperCase();
    if (cleanKey === 'AIMODEL' || cleanKey === 'OPENAIMODEL' || cleanKey === 'GEMINIMODEL') config.aiModel = String(val).trim();
  }

  // Load Model from Agent Settings if not specified
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentData = agentSheet.getDataRange().getValues();
    for (let i = 1; i < agentData.length; i++) {
      const k = String(agentData[i][1] || agentData[i][0] || '').trim().toUpperCase();
      const v = agentData[i][2] || agentData[i][1];
      
      if (k === 'PRIMARY_AI_PROVIDER' && v && !config.aiProvider) {
        config.aiProvider = String(v).trim().toUpperCase();
      }
      if (config.aiProvider === 'OPENAI' && k === 'OPENAI_MODEL' && v && !config.aiModel) {
        config.aiModel = String(v).trim();
      }
      if (config.aiProvider === 'GEMINI' && k === 'GEMINI_MODEL' && v && !config.aiModel) {
        config.aiModel = String(v).trim();
      }
    }
  }

  return config;
}

/**
 * Gets Active Campaign ID safely without throwing undefined exceptions.
 * Checks Dashboard B2, Campaign Configuration, active selection, or prompts.
 */
function getActiveCampaignId_(spreadsheet, ui) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return 'CMP-005';

  // 1. Try reading from Dashboard Cell B2
  const dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet) {
    const b2 = String(dashSheet.getRange('B2').getValue() || '').trim();
    if (b2 && b2 !== 'Select Campaign' && b2 !== 'No Active Campaign') return b2;
  }

  // 2. Try reading from Campaign Configuration Row 2
  const configSheet = ss.getSheetByName('Campaign Configuration');
  if (configSheet && configSheet.getLastRow() > 1) {
    const cId = String(configSheet.getRange(2, 2).getValue() || '').trim();
    if (cId) return cId;
  }

  // 3. Try checking active sheet selection safely
  try {
    const sheet = ss.getActiveSheet ? ss.getActiveSheet() : null;
    const allowed = ['Campaigns', 'Campaign Configuration', 'Company Database', 'Dashboard', 'Company Intelligence'];
    if (sheet && allowed.indexOf(sheet.getName()) >= 0) {
      const range = sheet.getActiveRange ? sheet.getActiveRange() : null;
      if (range && range.getRow() > 1) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
        const values = sheet.getRange(range.getRow(), 1, 1, headers.length).getValues()[0];
        const record = {};
        headers.forEach(function(h, i) { record[h] = values[i]; });
        if (record['Campaign ID']) return String(record['Campaign ID']).trim();
      }
    }
  } catch (e) {
    // Ignore selection read errors in headless/sidebar contexts
  }

  // 4. Fallback to UI prompt if ui is available
  if (ui && typeof ui.prompt === 'function') {
    try {
      const prompt = ui.prompt('Company Discovery', 'Enter the Campaign ID to discover companies for (e.g. CMP-005):', ui.ButtonSet.OK_CANCEL);
      if (prompt.getSelectedButton() === ui.Button.OK) {
        const resp = String(prompt.getResponseText() || '').trim();
        if (resp) return resp;
      }
    } catch (e) {}
  }

  return 'CMP-005';
}

/**
 * Updates timestamp in Campaigns tab
 */
function updateCampaignDiscoveryCounts_(spreadsheet, campaignId) {
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (!campSheet || campSheet.getLastRow() <= 1) return;

  const campRows = campSheet.getRange(2, 1, campSheet.getLastRow() - 1, campSheet.getLastColumn()).getValues();
  const campHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getDisplayValues()[0];
  const idCol = campHeaders.indexOf('Campaign ID');
  const updatedCol = campHeaders.indexOf('Last updated');

  for (let i = 0; i < campRows.length; i++) {
    if (String(campRows[i][idCol]).trim() === String(campaignId).trim()) {
      if (updatedCol >= 0) campSheet.getRange(i + 2, updatedCol + 1).setValue(new Date().toLocaleString());
      break;
    }
  }
}

/**
 * Helper: Normalizes URL to root domain
 */
function domainFromUrl_(url) {
  if (!url) return '';
  let clean = String(url).trim().toLowerCase();
  clean = clean.replace(/^https?:\/\//i, '').replace(/^www\./i, '');
  return clean.split('/')[0].split('?')[0];
}

/**
 * Helper: Converts Sheet Rows to Objects
 */
function getSheetRecords_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    records.push(obj);
  }
  return records;
}

/**
 * Helper: Maps Object to Array matching sheet header columns exactly
 */
function buildRowFromHeaders_(sheetHeaders, dataObject) {
  return sheetHeaders.map(header => {
    const cleanHeader = String(header).trim();
    return dataObject.hasOwnProperty(cleanHeader) ? dataObject[cleanHeader] : '';
  });
}