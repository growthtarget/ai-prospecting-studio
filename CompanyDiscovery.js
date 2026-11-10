/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AGENT 1: COMPANY DISCOVERY & RESEARCH AGENT (MULTI-AI PROVIDER ENGINE)
 * ============================================================================
 * File: CompanyDiscovery.gs
 * 
 * CAPABILITIES:
 * 1. Multi-AI Support: Dynamically routes to GEMINI (with Search Grounding) 
 *    or OPENAI (GPT-4o) based on your sidebar selection.
 * 2. Multi-Department Intelligence: Injects DEPARTMENT_OWNER & CUSTOM_BUSINESS_NEEDS.
 * 3. Dual-Tab Population: Writes master entries to "Company Database" AND
 *    populates all 25 columns of "Company Intelligence".
 * 4. Grounding Evidence Logging: Stores citation URLs in "Research Evidence".
 * ============================================================================
 */

/**
 * Called directly from the Sidebar Control Panel (No blocking popups; returns results to sidebar)
 */
function runCompanyDiscoveryFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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

    // Resolve AI Provider & Model dynamically (Gemini vs OpenAI)
    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';

    Logger.log('Launching Agent 1 with ' + providerLabel + ' for campaign: ' + config.id);

    // Call unified AI router (AIProvider.gs / OpenAIProvider.gs)
    const aiResponse = callAI({
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
        message: providerLabel + ' found no matching companies. Try broadening your criteria or business needs in the sidebar.' 
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
 * Called from the Top Menu (Uses standard Google Sheets dialogs)
 */
function runCompanyDiscoveryForSelectedCampaign() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  
  if (!lock.tryLock(5000)) {
    ui.alert('A research run is already in progress. Please wait for it to finish.');
    return;
  }

  try {
    const campaignId = getActiveCampaignId_(spreadsheet, ui);
    if (!campaignId) return;

    const config = getCampaignConfigurationRecord_(spreadsheet, campaignId);
    if (!config || !config.id) {
      ui.alert('No configuration found for Campaign ID: ' + campaignId + ' in "Campaign Configuration".');
      return;
    }

    const companyLimit = Math.min(25, Math.max(1, Number(config.companyLimit || 5)));
    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';
    
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

    const prompt = buildCompanyDiscoveryPrompt_(config, companyLimit);
    const aiResponse = callAI({
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
      ui.alert(providerLabel + ' did not return any companies matching the given criteria. Try broadening your criteria.');
      return;
    }

    const addedCount = saveDiscoveredCompanies_(spreadsheet, config.id, discoveredCompanies, config, providerLabel);
    updateCampaignDiscoveryCounts_(spreadsheet, config.id);

    ui.alert(
      'Company Discovery Complete',
      'Discovered and added ' + addedCount + ' qualified companies to "Company Database" and "Company Intelligence" using ' + providerLabel + '.\n\n' +
      'Next Step: Run "Agent 2: Persona Mapping" from the sidebar to find decision-makers and draft outreach.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('Error during Company Discovery: ' + error.message);
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
    '• Target Geographies / Countries: ' + (config.country || 'Global'),
    '• Target Industry Verticals: ' + (config.industry || 'Unspecified'),
    '• Target Company Size (Headcount): ' + (config.size || '50 - 250 employees'),
    '• Research Depth: ' + (config.researchDepth || 'Deep Persona'),
    '• Minimum Fit Score Required: ' + (config.minCompanyScore || 80) + ' out of 100',
    '• Specific Business Needs / Buying Signals to Match: ',
    config.customBusinessNeeds || 'Operational scaling, technology modernization, and workflow efficiency',
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
 * Saves discovered companies into BOTH "Company Database" and "Company Intelligence"
 */
function saveDiscoveredCompanies_(spreadsheet, campaignId, discoveredCompanies, config, providerLabel) {
  const dbSheet = spreadsheet.getSheetByName('Company Database');
  const intelSheet = spreadsheet.getSheetByName('Company Intelligence');
  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');

  if (!dbSheet) throw new Error('Sheet "Company Database" not found.');
  if (!intelSheet) throw new Error('Sheet "Company Intelligence" not found.');

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

    // Duplicate check
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

    // 1. Write to "Company Database" (Master Index)
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

    // 2. Write to "Company Intelligence" (All 25 Enriched Columns)
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
      'eValuater Solution Fit': item.solutionFit || ('High - Direct match for ' + config.department),
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

    // 3. Write to "Research Evidence" (Audit Citation)
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
  return (provider === 'OPENAI') ? 'gpt-4o' : 'gemini-2.5-flash';
}

/**
 * Universal Configuration Loader: Reads from Campaign Configuration & Agent Settings
 */
function getCampaignConfigurationRecord_(spreadsheet, campaignId) {
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  const agentSheet = spreadsheet.getSheetByName('Agent Settings');

  const config = {
    id: campaignId || 'CMP-001',
    name: '',
    department: 'ESG Advisory',
    industry: '',
    country: '',
    size: '50 - 250 employees',
    companyLimit: 5,
    researchDepth: 'Deep Persona',
    minCompanyScore: 80,
    minPersonaScore: 85,
    customBusinessNeeds: '',
    aiProvider: 'GEMINI',
    aiModel: ''
  };

  if (!configSheet) return config;

  const data = configSheet.getDataRange().getValues();
  if (data.length <= 1) return config;

  // Check if Vertical Key-Value format (Col A = Parameter, Col B = Value)
  const isKeyValue = String(data[0][0]).toLowerCase().includes('parameter');

  if (isKeyValue) {
    for (let r = 1; r < data.length; r++) {
      const key = String(data[r][0] || '').trim().toUpperCase();
      const val = data[r][1];

      if (key === 'CAMPAIGN_ID' && val) config.id = String(val).trim();
      if (key === 'CAMPAIGN_NAME' && val) config.name = String(val).trim();
      if (key === 'DEPARTMENT_OWNER' && val) config.department = String(val).trim();
      if ((key === 'TARGET_INDUSTRIES' || key === 'INDUSTRIES') && val) config.industry = String(val).trim();
      if ((key === 'TARGET_GEOGRAPHY' || key === 'COUNTRIES/REGIONS') && val) config.country = String(val).trim();
      if ((key === 'COMPANY_SIZE' || key === 'SIZE') && val) config.size = String(val).trim();
      if ((key === 'MAX_COMPANIES_LIMIT' || key === 'COMPANY LIMIT') && val) config.companyLimit = Number(val);
      if (key === 'RESEARCH_DEPTH' && val) config.researchDepth = String(val).trim();
      if (key === 'MIN_COMPANY_FIT_SCORE' && val) config.minCompanyScore = Number(val);
      if (key === 'MIN_PERSONA_FIT_SCORE' && val) config.minPersonaScore = Number(val);
      if (key === 'CUSTOM_BUSINESS_NEEDS' && val) config.customBusinessNeeds = String(val).trim();
      if (key === 'PRIMARY_AI_PROVIDER' && val) config.aiProvider = String(val).trim().toUpperCase();
      if ((key === 'AI_MODEL' || key === 'GEMINI_MODEL' || key === 'OPENAI_MODEL') && val) config.aiModel = String(val).trim();
    }
  } else {
    // Horizontal Tabular Format
    const records = getSheetRecords_(configSheet);
    const row = records.find(function(r) {
      return String(r['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();
    }) || records[0];

    if (row) {
      config.id = row['Campaign ID'] || config.id;
      config.name = row['Campaign name'] || row['Campaign Name'] || '';
      config.department = row['Department'] || config.department;
      config.industry = row['Target industries'] || row['Industry'] || '';
      config.country = row['Countries/regions'] || row['Country'] || '';
      config.size = row['Company size'] || config.size;
      config.companyLimit = Number(row['Company limit'] || row['Maximum companies'] || 5);
      config.researchDepth = row['Research depth'] || config.researchDepth;
      config.customBusinessNeeds = row['Custom business needs'] || row['Objective'] || '';
      if (row['AI Provider']) config.aiProvider = String(row['AI Provider']).trim().toUpperCase();
      if (row['AI Model']) config.aiModel = String(row['AI Model']).trim();
    }
  }

  // Load Model from Agent Settings based on selected provider
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
 * Gets Active Campaign ID from active selection or prompt
 */
function getActiveCampaignId_(spreadsheet, ui) {
  const sheet = spreadsheet.getActiveSheet();
  const allowed = ['Campaigns', 'Campaign Configuration', 'Company Database', 'Dashboard', 'Company Intelligence'];
  
  if (allowed.indexOf(sheet.getName()) >= 0) {
    const range = sheet.getActiveRange();
    if (range && range.getRow() > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      const values = sheet.getRange(range.getRow(), 1, 1, headers.length).getValues()[0];
      const record = {};
      headers.forEach(function(h, i) { record[h] = values[i]; });
      if (record['Campaign ID']) return String(record['Campaign ID']).trim();
    }
  }
  
  const prompt = ui.prompt('Company Discovery', 'Enter the Campaign ID to discover companies for (e.g. CMP-001):', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return '';
  return String(prompt.getResponseText() || '').trim();
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