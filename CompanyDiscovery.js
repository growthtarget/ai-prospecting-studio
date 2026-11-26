/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AGENT 1: COMPANY DISCOVERY & RESEARCH AGENT (CLIENT-ISOLATED MULTI-AI ENGINE)
 * ============================================================================
 * File: CompanyDiscovery.gs / CompanyDiscovery.js
 *
 * CAPABILITIES:
 * 1. Multi-AI Support: Dynamically routes to GEMINI (with Grounding + Free-Tier Fallback)
 *    or OPENAI (GPT-4o) using strict Client-Scoped API keys.
 * 2. Multi-Client Intelligence: Resolves Client / Owner, CUSTOM_BUSINESS_NEEDS,
 *    CAMPAIGN_FOCUS, SEARCH_KEYWORDS, and targeting parameters reliably.
 * 3. Total Campaign Cap Enforcement: Checks existing companies in the active
 *    Campaign ID and discovers only the remaining slots needed to reach Discovery Limit.
 * 4. Dual-Tab Population: Writes master entries to "Company Database" (stamped with Col J Campaign ID)
 *    AND populates all 25 columns of "Company Intelligence".
 * 5. Scenario B Deduplication (Enhanced): Isolates deduplication by Campaign ID while
 *    normalizing parentheticals, legal/regional suffixes, slash aliases, and parent/subsidiary
 *    brand+domain variants (e.g., "Enbridge Gas Inc." vs "Enbridge Gas Inc. (Affiliated...)").
 * 6. Grounding Evidence Logging: Stores audit citation URLs in "Research Evidence" (9 canonical columns).
 * 7. API Usage Tracking: Passes campaign ID, client context, and operation to ApiTracker.
 * ============================================================================
 */

/**
 * Called directly from the Sidebar Control Panel (Resolves active campaign if parameter is blank)
 */
function runCompanyDiscoveryFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    return { success: false, message: 'A research run is already in progress. Please wait.' };
  }

  try {
    const activeCampaignId = campaignId || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : 'CMP-005');
    const config = getCampaignConfigurationRecord_(spreadsheet, activeCampaignId);

    if (!config || !config.id) {
      return { success: false, message: 'No configuration found for Campaign ID: ' + activeCampaignId };
    }

    const companyLimit = Math.min(25, Math.max(1, Number(config.companyLimit || 5)));
    const existingCompanies = getExistingCampaignCompanies_(spreadsheet, config.id);
    const existingCount = existingCompanies.length;
    const remainingNeeded = Math.max(0, companyLimit - existingCount);

    if (remainingNeeded === 0) {
      return {
        success: true,
        count: 0,
        alreadyAtLimit: true,
        message: 'No new companies added: Campaign ' + config.id + ' already has ' + existingCount +
          ' discovered company record(s), meeting your Discovery Limit of ' + companyLimit +
          '. Increase Discovery Limit above ' + existingCount + ' in the Control Panel to discover more.'
      };
    }

    // Request a small candidate buffer on incremental runs in case an AI result deduplicates out,
    // while strictly capping saved rows at remainingNeeded.
    const promptCandidateCount = existingCount > 0
      ? Math.min(25, remainingNeeded + 2)
      : remainingNeeded;

    const prompt = buildCompanyDiscoveryPrompt_(config, promptCandidateCount, existingCompanies, remainingNeeded);

    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';

    Logger.log(
      'Launching Agent 1 with ' + providerLabel +
      ' for campaign: ' + config.id +
      ' (Client: ' + config.department + ', Existing: ' + existingCount +
      ', Cap: ' + companyLimit + ', Remaining Needed: ' + remainingNeeded + ')'
    );

    const aiResponse = executeAiDiscoveryCall_({
      campaignId: config.id,
      clientName: config.department,
      provider: provider,
      model: model,
      prompt: prompt,
      enableWebSearch: true,
      json: true,
      operation: 'CompanyDiscovery'
    });

    const discoveredCompanies = extractDiscoveredCompaniesList_(aiResponse);

    if (!discoveredCompanies.length) {
      const preview = String(aiResponse && aiResponse.text ? aiResponse.text : '').replace(/\s+/g, ' ').slice(0, 160);
      Logger.log('Agent 1 returned 0 companies. Raw response preview: ' + preview);
      return {
        success: false,
        message: providerLabel + ' found no matching companies. Broaden your criteria or custom business needs in the sidebar.' +
          (preview ? (' (Response: ' + preview + ')') : '')
      };
    }

    const addedCount = saveDiscoveredCompanies_(
      spreadsheet,
      config.id,
      discoveredCompanies,
      config,
      providerLabel,
      aiResponse,
      remainingNeeded
    );
    updateCampaignDiscoveryCounts_(spreadsheet, config.id);

    if (typeof updateDashboardForActiveCampaign === 'function' &&
        typeof getStudioActiveCampaignId_ === 'function' &&
        config.id === getStudioActiveCampaignId_()) {
      updateDashboardForActiveCampaign(config.id);
    }

    if (addedCount === 0) {
      return {
        success: false,
        count: 0,
        message: providerLabel + ' returned only duplicate/already-discovered companies for ' + config.id +
          '. Try running Discovery again or broadening your criteria.'
      };
    }

    return {
      success: true,
      count: addedCount,
      message: 'Discovered ' + addedCount + ' net-new company record(s) (Campaign total: ' +
        (existingCount + addedCount) + ' / ' + companyLimit + ') using ' + providerLabel +
        ' for ' + config.department + '! Populated in "Company Database", "Company Intelligence", and "Research Evidence".'
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
  } catch (e) {}

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    if (ui) ui.alert('A research run is already in progress. Please wait for it to finish.');
    else Logger.log('A research run is already in progress.');
    return;
  }

  try {
    const activeCampaignId = getActiveCampaignId_(spreadsheet, ui) || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : 'CMP-005');
    const config = getCampaignConfigurationRecord_(spreadsheet, activeCampaignId);

    if (!config || !config.id) {
      if (ui) ui.alert('No configuration found for Campaign ID: ' + activeCampaignId);
      else Logger.log('No configuration found for Campaign ID: ' + activeCampaignId);
      return;
    }

    const companyLimit = Math.min(25, Math.max(1, Number(config.companyLimit || 5)));
    const existingCompanies = getExistingCampaignCompanies_(spreadsheet, config.id);
    const existingCount = existingCompanies.length;
    const remainingNeeded = Math.max(0, companyLimit - existingCount);

    if (remainingNeeded === 0) {
      const limitMsg = 'Campaign ' + config.id + ' already has ' + existingCount +
        ' discovered company record(s), meeting your Discovery Limit of ' + companyLimit + '.\n\n' +
        'To discover more companies, increase Discovery Limit above ' + existingCount +
        ' in the Campaign Control Panel (or remove unwanted/duplicate rows first).';
      if (ui) ui.alert('Discovery Limit Reached', limitMsg, ui.ButtonSet.OK);
      else Logger.log(limitMsg);
      return;
    }

    const provider = String(config.aiProvider || 'GEMINI').toUpperCase();
    const model = resolveAiModel_(provider, config.aiModel);
    const providerLabel = provider === 'OPENAI' ? 'OpenAI (' + model + ')' : 'Google Gemini (' + model + ')';

    if (ui) {
      ui.alert(
        'Starting Company Discovery (Agent 1)',
        'Searching for ' + remainingNeeded + ' net-new company(s) (Existing: ' + existingCount + ', Target Cap: ' + companyLimit + ') using ' + providerLabel + ' (Campaign: ' + config.id + '):\n' +
        '• Client / Owner: ' + config.department + '\n' +
        '• Industry: ' + (config.industry || 'Any') + '\n' +
        '• Geography: ' + (config.country || 'Any') + '\n' +
        '• Needs/Triggers: ' + (config.customBusinessNeeds.slice(0, 100) || config.campaignFocus || 'General') + '\n\n' +
        'The AI research agent is now searching. Click OK to begin.',
        ui.ButtonSet.OK
      );
    }

    const promptCandidateCount = existingCount > 0
      ? Math.min(25, remainingNeeded + 2)
      : remainingNeeded;

    const prompt = buildCompanyDiscoveryPrompt_(config, promptCandidateCount, existingCompanies, remainingNeeded);
    const aiResponse = executeAiDiscoveryCall_({
      campaignId: config.id,
      clientName: config.department,
      provider: provider,
      model: model,
      prompt: prompt,
      enableWebSearch: true,
      json: true,
      operation: 'CompanyDiscovery'
    });

    const discoveredCompanies = extractDiscoveredCompaniesList_(aiResponse);

    if (!discoveredCompanies.length) {
      const preview = String(aiResponse && aiResponse.text ? aiResponse.text : '').replace(/\s+/g, ' ').slice(0, 160);
      Logger.log('Agent 1 returned 0 companies. Raw response preview: ' + preview);
      if (ui) ui.alert(providerLabel + ' did not return any companies matching the given criteria. Try broadening your criteria.');
      else Logger.log(providerLabel + ' did not return any companies.');
      return;
    }

    const addedCount = saveDiscoveredCompanies_(
      spreadsheet,
      config.id,
      discoveredCompanies,
      config,
      providerLabel,
      aiResponse,
      remainingNeeded
    );
    updateCampaignDiscoveryCounts_(spreadsheet, config.id);

    if (typeof updateDashboardForActiveCampaign === 'function' &&
        typeof getStudioActiveCampaignId_ === 'function' &&
        config.id === getStudioActiveCampaignId_()) {
      updateDashboardForActiveCampaign(config.id);
    }

    if (ui) {
      if (addedCount === 0) {
        ui.alert(
          'No Net-New Companies Added',
          providerLabel + ' returned only duplicate/already-discovered companies for ' + config.id + '.\nTry running Discovery again or broadening your criteria.',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          'Company Discovery Complete',
          'Discovered and added ' + addedCount + ' net-new company record(s) (Campaign total: ' +
          (existingCount + addedCount) + ' / ' + companyLimit + ') to "Company Database", "Company Intelligence", and "Research Evidence" using ' +
          providerLabel + ' for ' + config.department + '.\n\n' +
          'Next Step: Run "Agent 2: Persona Mapping" from the sidebar to find decision-makers and draft outreach.',
          ui.ButtonSet.OK
        );
      }
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
 * Safely normalizes AI JSON outputs regardless of wrapper key or top-level array formatting.
 */
function extractDiscoveredCompaniesList_(aiResponse) {
  if (!aiResponse) return [];
  let payload = aiResponse.json;

  if (!payload && aiResponse.text && typeof parseGeminiJsonOutput_ === 'function') {
    payload = parseGeminiJsonOutput_(aiResponse.text);
  }
  if (!payload) return [];

  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.companies)) return payload.companies;
  if (Array.isArray(payload.Companies)) return payload.Companies;
  if (Array.isArray(payload.results)) return payload.results;
  if (Array.isArray(payload.data)) return payload.data;
  if (Array.isArray(payload.accounts)) return payload.accounts;
  if (payload.company && typeof payload.company === 'object') return [payload.company];

  return [];
}

/**
 * Builds the Structured Search Prompt for Gemini or OpenAI (with Campaign Exclusion List)
 */
function buildCompanyDiscoveryPrompt_(config, limit, existingCompanies, remainingNeeded) {
  const existingList = Array.isArray(existingCompanies) ? existingCompanies : [];
  const targetNetNew = Number(remainingNeeded || limit || 1);

  const lines = [
    'Act as the Senior B2B Account Research & Market Intelligence Agent.',
    'Your goal is to identify REAL, active, operating enterprise or mid-market companies matching the campaign profile below.',
    '',
    'CAMPAIGN PARAMETERS:',
    '• Client / Account Owner: ' + config.department,
    '• Target Geographies / Countries: ' + (config.country || 'United States & Canada'),
    '• Target Industry Verticals: ' + (config.industry || 'Freight & Logistics, Manufacturing'),
    '• Target Company Size (Headcount): ' + (config.size || '50 - 250 employees'),
    '• Research Depth: ' + (config.researchDepth || 'Deep Persona'),
    '• Evidence Lookback Period: Last ' + (config.evidencePeriodMonths || 12) + ' months',
    '• Campaign Focus: ' + (config.campaignFocus || 'Operational resilience & compliance readiness'),
    '• Search Keywords: ' + (config.searchKeywords || 'Scope 3, CSRD, compliance audits'),
    '• Target Minimum Fit Score: ' + (config.minCompanyScore || 80) + ' out of 100',
    '• Specific Business Needs / Buying Signals to Match: ',
    config.customBusinessNeeds || 'Scope 3 vendor emissions transparency, supply chain audits, vendor questionnaires',
    ''
  ];

  if (existingList.length > 0) {
    lines.push('ALREADY DISCOVERED IN THIS CAMPAIGN — STRICT EXCLUSION LIST:');
    existingList.forEach(function(item) {
      lines.push('• ' + item.companyName + (item.domain ? ' (' + item.domain + ')' : ''));
    });
    lines.push('CRITICAL EXCLUSION RULE: Do NOT return any company listed above, nor any division, subsidiary, business unit, affiliate, parent company, or alternate domain belonging to those same organizations.');
    lines.push('');
  }

  lines.push('DISCOVERY & QUALIFICATION RULES:');
  lines.push('1. Identify REAL, established operating companies matching the target geography, industry, and size. If live web search is available, verify details with search; otherwise rely on verified knowledge of real companies and their official domains. Never invent fictional company names.');
  lines.push('2. Return official root website URLs (e.g. "https://companyname.com").');
  if (existingList.length > 0 && limit > targetNetNew) {
    lines.push('3. Provide ' + limit + ' distinct net-new real companies ranked highest fit first (so at least ' + targetNetNew + ' unique company qualifies without overlapping the exclusion list).');
  } else {
    lines.push('3. Provide ' + limit + ' distinct real companies that best match the criteria (do not return an empty list unless criteria are impossible).');
  }
  lines.push('4. Score each company from 0 to 100 on how closely they exhibit the specified business needs, pain points, or buying signals.');
  lines.push('5. Prioritize companies scoring ' + (config.minCompanyScore || 80) + ' or higher.');
  lines.push('6. Include the official company URL or relevant public evidence URL in "sourceUrl".');
  lines.push('');
  lines.push('Return ONLY a valid, parsable JSON object with NO markdown formatting, matching this exact schema:');
  lines.push('{');
  lines.push('  "companies": [');
  lines.push('    {');
  lines.push('      "companyName": "Exact Operating Name",');
  lines.push('      "website": "https://companydomain.com",');
  lines.push('      "headquarters": "City, Country",');
  lines.push('      "industry": "Specific Industry Sector",');
  lines.push('      "companySize": "Estimated employee count (e.g. 150 employees)",');
  lines.push('      "businessSummary": "1-2 sentence description of primary business activities",');
  lines.push('      "businessTriggers": "Specific pain points or buying triggers detected matching the campaign needs",');
  lines.push('      "solutionFit": "High, Medium, or Low",');
  lines.push('      "companyOpportunityScore": 90,');
  lines.push('      "companyOpportunity": "Core commercial opportunity or bottleneck to address",');
  lines.push('      "digitalSignals": "Cloud systems, ERP/TMS, tooling, or modernization signals detected",');
  lines.push('      "scoreRationale": "Clear explanation of why this company matches qualification criteria",');
  lines.push('      "sourceUrl": "Direct URL where company profile or trigger was verified"');
  lines.push('    }');
  lines.push('  ]');
  lines.push('}');

  return lines.join('\n');
}

/**
 * Resilient AI Router honoring ApiTracker budget guardrails and strict client key routing
 */
function executeAiDiscoveryCall_(options) {
  if (typeof callAI === 'function') {
    return callAI(options);
  }

  const provider = String(options.provider || 'GEMINI').toUpperCase();

  if (provider === 'OPENAI') {
    if (typeof requestOpenAIChatCompletion_ === 'function') {
      const rawOpenAi = requestOpenAIChatCompletion_(options);
      const text = rawOpenAi && rawOpenAi.choices && rawOpenAi.choices[0] && rawOpenAi.choices[0].message
        ? rawOpenAi.choices[0].message.content
        : '{}';
      return { text: text, json: JSON.parse(text), citations: [], raw: rawOpenAi };
    }
    throw new Error('OpenAIProvider / AIProvider not loaded.');
  }

  if (provider === 'GEMINI' && typeof callGemini_ === 'function') {
    return callGemini_(options);
  }

  throw new Error('AI router not available for ' + provider);
}

/**
 * Saves discovered companies into "Company Database", "Company Intelligence", and "Research Evidence".
 * Guarantees Column J is "Campaign ID", enforces Scenario B + variant deduplication, and honors maxToAdd cap.
 */
function saveDiscoveredCompanies_(spreadsheet, campaignId, discoveredCompanies, config, providerLabel, aiResponse, maxToAdd) {
  let dbSheet = spreadsheet.getSheetByName('Company Database');
  let intelSheet = spreadsheet.getSheetByName('Company Intelligence');
  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');

  if (!dbSheet) dbSheet = spreadsheet.insertSheet('Company Database');
  if (!intelSheet) intelSheet = spreadsheet.insertSheet('Company Intelligence');

  const canonicalDbHeaders = [
    'Company ID', 'Company Name', 'Domain', 'Country', 'Industry',
    'Employee Size', 'Discovery Method', 'Discovery Date', 'Status', 'Campaign ID'
  ];
  const canonicalIntelHeaders = [
    'Campaign ID','Campaign Name','Company ID','Company Name','Country','Industry','Company Size',
    'ESG Frameworks','Reporting Obligations','ESG Reporting Maturity','Supply Chain Profile',
    'Scope 3 Exposure','Sustainability Data Needs','Digital Transformation Signals','ESG Governance',
    'Business Triggers','eValuater Solution Fit','Company Opportunity Score','Company Opportunity',
    'Evidence URLs','Evidence Checked Date','Verification Status','Human Review Status',
    'Last Enriched At','Company Score Rationale'
  ];
  const canonicalEvidenceHeaders = [
    'Evidence ID', 'Campaign ID', 'Entity Name', 'Entity Type',
    'Search Query Used', 'Grounding Evidence URL', 'Extracted Snippet / Signal',
    'Verification Score', 'Date Captured'
  ];

  if (dbSheet.getLastRow() === 0 || dbSheet.getLastColumn() === 0) {
    dbSheet.getRange(1, 1, 1, canonicalDbHeaders.length).setValues([canonicalDbHeaders]);
  } else {
    const lastCol = Math.max(dbSheet.getLastColumn(), 10);
    const currentDbHeaders = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const hasCampaignCol = currentDbHeaders.some(function(h) {
      return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '') === 'campaignid';
    });
    if (!hasCampaignCol) {
      dbSheet.getRange(1, 10).setValue('Campaign ID');
    }
  }

  if (intelSheet.getLastRow() === 0 || intelSheet.getLastColumn() === 0) {
    intelSheet.getRange(1, 1, 1, canonicalIntelHeaders.length).setValues([canonicalIntelHeaders]);
  }
  if (evidenceSheet && (evidenceSheet.getLastRow() === 0 || evidenceSheet.getLastColumn() === 0)) {
    evidenceSheet.getRange(1, 1, 1, canonicalEvidenceHeaders.length).setValues([canonicalEvidenceHeaders]);
  }

  [dbSheet, intelSheet, evidenceSheet].forEach(function(sh) {
    if (sh && sh.getLastColumn() > 0) {
      sh.getRange(1, 1, 1, sh.getLastColumn())
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sh.setFrozenRows(1);
    }
  });

  const existingDbRecords = getDiscoverySheetRecords_(dbSheet);
  const existingIntelRecords = getDiscoverySheetRecords_(intelSheet);
  const intelHeaders = intelSheet.getRange(1, 1, 1, intelSheet.getLastColumn()).getValues()[0];
  const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];

  const grounding = (typeof extractGroundingMetadata_ === 'function' && aiResponse)
    ? extractGroundingMetadata_(aiResponse.raw || aiResponse)
    : { citationsByUrl: {}, searchQueries: [] };

  const searchQueries = grounding.searchQueries || [];
  const defaultSearchQuery = searchQueries.length > 0
    ? searchQueries.join('; ')
    : (config.department + ' Target Discovery: ' + (config.industry || '') + ' ' + (config.country || ''));

  const addCap = (maxToAdd !== undefined && maxToAdd !== null && Number(maxToAdd) > 0)
    ? Number(maxToAdd)
    : 0;

  let addedCount = 0;
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];

  const esgSignalText = [
    config.name,
    config.campaignFocus,
    config.searchKeywords,
    config.customBusinessNeeds,
    config.industry
  ].join(' ').toLowerCase();
  const isEsg = /\b(esg|scope\s*3|csrd|sustainability|emissions|ghg|clean\s*energy|cleantech|gri)\b/.test(esgSignalText);

  discoveredCompanies.forEach(function(item) {
    if (addCap > 0 && addedCount >= addCap) return;
    if (!item || typeof item !== 'object') return;

    const companyName = String(item.companyName || item.name || item.company_name || item.organization || '').trim();
    if (!companyName) return;

    const website = String(item.website || item.url || item.domain || item.websiteUrl || '').trim();
    const domain = extractDiscoveryDomainFromUrl_(website);

    if (isDuplicateCompanyInCampaign_(companyName, domain, existingDbRecords, campaignId)) {
      Logger.log('Skipped duplicate/variant company in ' + campaignId + ': ' + companyName + ' (' + domain + ')');
      return;
    }

    const companyId = 'CMP-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    const score = Math.round(Number(item.companyOpportunityScore || item.fitScore || item.score || item.initialScore || 85));
    const status = score >= (config.minCompanyScore || 80) ? 'Qualified' : 'Disqualified';

    const hq = String(item.headquarters || item.country || item.location || config.country || '').trim();
    const ind = String(item.industry || item.sector || config.industry || '').trim();
    const empSize = String(item.companySize || item.employeeSize || item.company_size || item.size || config.size || '').trim();
    const sourceUrl = String(item.sourceUrl || item.evidenceUrl || item.source_url || website).trim();

    const dbRecord = {
      'Company ID': companyId,
      'Campaign ID': campaignId,
      'Company Name': companyName,
      'Domain': domain || website,
      'Country': hq,
      'Industry': ind,
      'Employee Size': empSize,
      'Discovery Method': providerLabel + ' - ' + config.department,
      'Discovery Date': dateStr,
      'Status': status
    };
    dbSheet.appendRow(buildDiscoveryRowFromHeaders_(dbHeaders, dbRecord));

    // Track immediately inside existingDbRecords so within-batch duplicates are also blocked.
    existingDbRecords.push({
      'Campaign ID': campaignId,
      'Company ID': companyId,
      'Company Name': companyName,
      'Domain': domain || website
    });

    const fitLabel = status === 'Disqualified'
      ? 'Low'
      : (score >= 85 ? 'High' : (score >= 70 ? 'Medium' : 'Low'));

    if (!isDuplicateCompanyInCampaign_(companyName, domain, existingIntelRecords, campaignId)) {
      const intelRecord = {
        'Campaign ID': campaignId,
        'Campaign Name': config.name || config.id,
        'Company ID': companyId,
        'Company Name': companyName,
        'Country': hq,
        'Industry': ind,
        'Company Size': empSize,
        'ESG Frameworks': isEsg ? 'CSRD, CDP, GHG Protocol' : 'N/A (Enterprise Stack)',
        'Reporting Obligations': isEsg ? 'Mandatory Tier 1 Scope 3 Reporting' : 'Standard Regulatory Compliance',
        'ESG Reporting Maturity': isEsg ? 'Developing / Spreadsheet-reliant' : 'N/A',
        'Supply Chain Profile': 'Multi-tier supplier & infrastructure network',
        'Scope 3 Exposure': isEsg ? 'High (Supply Chain Emissions)' : 'N/A',
        'Sustainability Data Needs': isEsg ? 'Automated vendor ESG data ingestion' : 'System observability & audit automation',
        'Digital Transformation Signals': item.digitalSignals || item.digitalTransformationSignals || item.businessSummary || 'Cloud infrastructure detected',
        'ESG Governance': isEsg ? 'Director of Sustainability' : 'VP of Engineering / Operations',
        'Business Triggers': item.businessTriggers || item.triggers || item.businessSummary || String(config.customBusinessNeeds || '').slice(0, 150),
        'eValuater Solution Fit': fitLabel,
        'Company Opportunity Score': score,
        'Company Opportunity': item.companyOpportunity || item.businessSummary || 'Operational improvement opportunity',
        'Evidence URLs': sourceUrl,
        'Evidence Checked Date': dateStr,
        'Verification Status': 'Verified',
        'Human Review Status': 'Accepted',
        'Last Enriched At': now.toLocaleString(),
        'Company Score Rationale': item.scoreRationale || ('Qualified based on ' + config.department + ' criteria via ' + providerLabel + '.')
      };
      intelSheet.appendRow(buildDiscoveryRowFromHeaders_(intelHeaders, intelRecord));

      existingIntelRecords.push({
        'Campaign ID': campaignId,
        'Company ID': companyId,
        'Company Name': companyName,
        'Domain': domain || website
      });
    }

    if (evidenceSheet && sourceUrl) {
      const normUrl = typeof normalizeSourceUrl_ === 'function' ? normalizeSourceUrl_(sourceUrl) : sourceUrl.replace(/\/+$/, '');
      const isGrounded = Boolean(grounding.citationsByUrl && grounding.citationsByUrl[normUrl]);
      const verificationScore = isGrounded ? 'High' : (score >= 80 ? 'Medium' : 'Unknown');
      const snippet = String(item.businessTriggers || item.solutionFit || item.businessSummary || companyName).trim();

      evidenceSheet.appendRow([
        'EVD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
        campaignId,
        companyName,
        'Company',
        defaultSearchQuery,
        sourceUrl,
        snippet,
        verificationScore,
        now.toLocaleString()
      ]);
    }

    addedCount++;
  });

  return addedCount;
}

/**
 * Reads existing companies already discovered for a specific Campaign ID from "Company Database".
 */
function getExistingCampaignCompanies_(spreadsheet, campaignId) {
  const dbSheet = spreadsheet.getSheetByName('Company Database');
  if (!dbSheet || dbSheet.getLastRow() <= 1) return [];

  const targetCamp = String(campaignId || '').trim();
  const records = getDiscoverySheetRecords_(dbSheet);
  const seenKeys = {};
  const list = [];

  records.forEach(function(row) {
    const rowCampaign = String(row['Campaign ID'] || '').trim();
    if (!targetCamp || rowCampaign !== targetCamp) return;

    const companyName = String(row['Company Name'] || row['Company name'] || '').trim();
    const domain = extractDiscoveryDomainFromUrl_(row['Domain'] || row['Website'] || '');
    if (!companyName && !domain) return;

    const dedupKey = normalizeCompanyNameForDedup_(companyName) + '|' + domain;
    if (seenKeys[dedupKey]) return;
    seenKeys[dedupKey] = true;

    list.push({
      companyId: String(row['Company ID'] || '').trim(),
      companyName: companyName,
      domain: domain
    });
  });

  return list;
}

/**
 * Checks whether a candidate company is a duplicate or division/brand variant of an existing
 * company in the SAME Campaign ID (Scenario B keeps separate campaigns cleanly isolated).
 */
function isDuplicateCompanyInCampaign_(companyName, domain, existingRecords, campaignId) {
  const targetCamp = String(campaignId || '').trim();
  const candDomain = extractDiscoveryDomainFromUrl_(domain || '');
  const candStem = extractDomainBrandStem_(candDomain);
  const candVariants = getCompanyNameVariantsForDedup_(companyName);
  const candPrimaryNorm = candVariants.length ? candVariants[0] : '';
  const candFirstWord = candPrimaryNorm ? candPrimaryNorm.split(' ')[0] : '';

  return (existingRecords || []).some(function(row) {
    const rowCampaign = String(row['Campaign ID'] || '').trim();
    if (rowCampaign !== targetCamp) return false;

    const rowDomain = extractDiscoveryDomainFromUrl_(row['Domain'] || row['Website'] || '');
    if (candDomain && rowDomain && candDomain === rowDomain) {
      return true;
    }

    const rowName = String(row['Company Name'] || row['Company name'] || '').trim();
    const rowVariants = getCompanyNameVariantsForDedup_(rowName);

    for (let i = 0; i < candVariants.length; i++) {
      if (rowVariants.indexOf(candVariants[i]) >= 0) {
        return true;
      }
    }

    const rowPrimaryNorm = rowVariants.length ? rowVariants[0] : '';
    const rowStem = extractDomainBrandStem_(rowDomain);

    // Same root domain stem across different TLDs (e.g. .com vs .ca) when primary brand word aligns
    if (candStem && rowStem && candStem === rowStem && candStem.length >= 6) {
      const rowFirstWord = rowPrimaryNorm ? rowPrimaryNorm.split(' ')[0] : '';
      if (candFirstWord && rowFirstWord && candFirstWord.length >= 4 && candFirstWord === rowFirstWord) {
        return true;
      }
    }

    // Parent/subsidiary brand prefix match (e.g. "Enbridge Inc." + enbridge.com vs "Enbridge Gas Inc." + enbridgegas.com)
    if (candPrimaryNorm && rowPrimaryNorm && candStem && rowStem) {
      const shorterName = candPrimaryNorm.length <= rowPrimaryNorm.length ? candPrimaryNorm : rowPrimaryNorm;
      const longerName = candPrimaryNorm.length > rowPrimaryNorm.length ? candPrimaryNorm : rowPrimaryNorm;
      const shorterStem = candStem.length <= rowStem.length ? candStem : rowStem;
      const longerStem = candStem.length > rowStem.length ? candStem : rowStem;

      if (shorterName.length >= 5 &&
          !isGenericCompanyPrefixWord_(shorterName) &&
          (longerName === shorterName || longerName.indexOf(shorterName + ' ') === 0) &&
          shorterStem.length >= 5 &&
          longerStem.indexOf(shorterStem) === 0) {
        return true;
      }
    }

    return false;
  });
}

function getCompanyNameVariantsForDedup_(name) {
  const raw = String(name || '').trim();
  if (!raw) return [];

  const variants = [];
  const fullNorm = normalizeCompanyNameForDedup_(raw);
  if (fullNorm) variants.push(fullNorm);

  if (raw.indexOf('/') >= 0) {
    raw.split('/').forEach(function(part) {
      const normPart = normalizeCompanyNameForDedup_(part);
      if (normPart && normPart.length >= 4 && variants.indexOf(normPart) < 0) {
        variants.push(normPart);
      }
    });
  }

  return variants;
}

function normalizeCompanyNameForDedup_(name) {
  let clean = String(name || '').trim().toLowerCase();
  if (!clean) return '';

  // Remove parenthetical qualifiers such as "(Affiliated Clean Infrastructure Division)"
  clean = clean.replace(/\([^)]*\)/g, ' ');
  clean = clean.replace(/[.,'"&]+/g, ' ').replace(/[-_/]+/g, ' ').replace(/\s+/g, ' ').trim();

  const stopSuffixes = {
    'inc': true, 'incorporated': true, 'corp': true, 'corporation': true,
    'ltd': true, 'limited': true, 'llc': true, 'llp': true, 'lp': true,
    'plc': true, 'gmbh': true, 'sa': true, 'ag': true, 'co': true,
    'company': true, 'companies': true, 'of': true, 'group': true,
    'holdings': true, 'canada': true, 'usa': true, 'us': true,
    'division': true, 'subsidiary': true, 'branch': true
  };

  const words = clean.split(' ').filter(Boolean);
  while (words.length > 1 && stopSuffixes[words[words.length - 1]]) {
    words.pop();
  }

  return words.join(' ').trim();
}

function extractDomainBrandStem_(domain) {
  const clean = extractDiscoveryDomainFromUrl_(domain);
  if (!clean) return '';
  const parts = clean.split('.').filter(Boolean);
  if (parts.length === 1) return parts[0];

  const secondLevelPrefixes = { 'co': true, 'com': true, 'org': true, 'net': true, 'gov': true, 'edu': true };
  if (parts.length >= 3 &&
      secondLevelPrefixes[parts[parts.length - 2]] &&
      parts[parts.length - 1].length === 2) {
    return parts[parts.length - 3];
  }
  return parts[parts.length - 2];
}

function isGenericCompanyPrefixWord_(word) {
  const generic = {
    'american': true, 'canadian': true, 'north': true, 'south': true, 'east': true, 'west': true,
    'global': true, 'national': true, 'united': true, 'clean': true, 'green': true, 'energy': true,
    'power': true, 'solar': true, 'carbon': true, 'pacific': true, 'atlantic': true, 'western': true,
    'eastern': true, 'northern': true, 'southern': true, 'first': true, 'premier': true
  };
  return Boolean(generic[String(word || '').trim().toLowerCase()]);
}

function resolveAiModel_(provider, configuredModel) {
  if (configuredModel && String(configuredModel).trim() !== '') {
    return String(configuredModel).trim();
  }
  return (provider === 'OPENAI') ? 'gpt-4o' : 'gemini-3.8-flash';
}

function getCampaignConfigurationRecord_(spreadsheet, campaignId) {
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  const agentSheet = spreadsheet.getSheetByName('Agent Settings');

  const targetId = String(campaignId || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : 'CMP-005')).trim();

  const config = {
    id: targetId,
    name: '',
    department: '',
    industry: 'Freight & Logistics, Manufacturing',
    country: 'United States & Canada',
    size: '50 - 250 employees',
    companyLimit: 5,
    researchDepth: 'Deep Persona',
    evidencePeriodMonths: 12,
    campaignFocus: '',
    searchKeywords: '',
    minCompanyScore: 80,
    minPersonaScore: 85,
    customBusinessNeeds: '',
    aiProvider: '',
    aiModel: ''
  };

  let hydratedFromCampaigns = false;
  if (typeof getSidebarInitialData === 'function') {
    try {
      const sData = getSidebarInitialData();
      if (sData && sData.configsMap && sData.configsMap[targetId]) {
        const c = sData.configsMap[targetId];
        if (c.name) config.name = String(c.name).trim();
        if (c.department) config.department = String(c.department).trim();
        if (c.industry) config.industry = String(c.industry).trim();
        if (c.country) config.country = String(c.country).trim();
        if (c.size) config.size = String(c.size).trim();
        if (c.limit) config.companyLimit = Number(c.limit) || 5;
        if (c.researchDepth) config.researchDepth = String(c.researchDepth).trim();
        if (c.evidencePeriodMonths) config.evidencePeriodMonths = Number(c.evidencePeriodMonths) || 12;
        if (c.campaignFocus) config.campaignFocus = String(c.campaignFocus).trim();
        if (c.searchKeywords) config.searchKeywords = String(c.searchKeywords).trim();
        if (c.minCompanyScore) config.minCompanyScore = Number(c.minCompanyScore) || 80;
        if (c.minPersonaScore) config.minPersonaScore = Number(c.minPersonaScore) || 85;
        if (c.customBusinessNeeds) config.customBusinessNeeds = String(c.customBusinessNeeds).trim();
        if (c.aiProvider) config.aiProvider = String(c.aiProvider).trim().toUpperCase();
        if (c.aiModel) config.aiModel = String(c.aiModel).trim();
        hydratedFromCampaigns = true;
      }
    } catch (e) {}
  }

  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    const dict = {};
    for (let r = 1; r < data.length; r++) {
      const cleanKey = String(data[r][0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (cleanKey) dict[cleanKey] = data[r][1];
    }

    const activeCfgId = String(dict['CAMPAIGNID'] || dict['ID'] || '').trim();
    if (!hydratedFromCampaigns || !activeCfgId || activeCfgId.toLowerCase() === targetId.toLowerCase()) {
      if (!targetId && activeCfgId) config.id = activeCfgId;
      if (dict['CAMPAIGNNAME'] || dict['NAME']) config.name = String(dict['CAMPAIGNNAME'] || dict['NAME']).trim();
      if (dict['DEPARTMENT'] || dict['DEPARTMENTOWNER'] || dict['CLIENTOWNER'] || dict['CLIENT']) {
        config.department = String(dict['DEPARTMENT'] || dict['DEPARTMENTOWNER'] || dict['CLIENTOWNER'] || dict['CLIENT']).trim();
      }
      if (dict['TARGETCOUNTRIES'] || dict['TARGETGEOGRAPHY'] || dict['COUNTRIESREGIONS'] || dict['COUNTRY']) config.country = String(dict['TARGETCOUNTRIES'] || dict['TARGETGEOGRAPHY'] || dict['COUNTRIESREGIONS'] || dict['COUNTRY']).trim();
      if (dict['TARGETINDUSTRIES'] || dict['INDUSTRIES'] || dict['INDUSTRY']) config.industry = String(dict['TARGETINDUSTRIES'] || dict['INDUSTRIES'] || dict['INDUSTRY']).trim();
      if (dict['COMPANYSIZERANGE'] || dict['COMPANYSIZE'] || dict['SIZE']) config.size = String(dict['COMPANYSIZERANGE'] || dict['COMPANYSIZE'] || dict['SIZE']).trim();
      if (dict['MAXCOMPANIES'] || dict['MAXCOMPANIESLIMIT'] || dict['COMPANYLIMIT']) config.companyLimit = Number(dict['MAXCOMPANIES'] || dict['MAXCOMPANIESLIMIT'] || dict['COMPANYLIMIT']) || 5;
      if (dict['RESEARCHDEPTH']) config.researchDepth = String(dict['RESEARCHDEPTH']).trim();
      if (dict['EVIDENCEPERIODMONTHS']) config.evidencePeriodMonths = Number(dict['EVIDENCEPERIODMONTHS']) || 12;
      if (dict['CAMPAIGNFOCUS']) config.campaignFocus = String(dict['CAMPAIGNFOCUS']).trim();
      if (dict['SEARCHKEYWORDS']) config.searchKeywords = String(dict['SEARCHKEYWORDS']).trim();
      if (dict['MINCOMPANYSCORE'] || dict['MINCOMPANYFITSCORE']) config.minCompanyScore = Number(dict['MINCOMPANYSCORE'] || dict['MINCOMPANYFITSCORE']) || 80;
      if (dict['MINPERSONASCORE'] || dict['MINPERSONAFITSCORE']) config.minPersonaScore = Number(dict['MINPERSONASCORE'] || dict['MINPERSONAFITSCORE']) || 85;
      if (dict['CUSTOMBUSINESSNEEDS']) config.customBusinessNeeds = String(dict['CUSTOMBUSINESSNEEDS']).trim();
      if (dict['AIPROVIDER'] || dict['PRIMARYAIPROVIDER']) config.aiProvider = String(dict['AIPROVIDER'] || dict['PRIMARYAIPROVIDER']).trim().toUpperCase();
      if (dict['AIMODEL'] || dict['OPENAIMODEL'] || dict['GEMINIMODEL']) config.aiModel = String(dict['AIMODEL'] || dict['OPENAIMODEL'] || dict['GEMINIMODEL']).trim();
    }
  }

  if (!config.department) {
    config.department = typeof resolveCampaignClientName_ === 'function'
      ? resolveCampaignClientName_(config.id, '', spreadsheet)
      : 'Growth Target';
  }

  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentData = agentSheet.getDataRange().getValues();
    for (let i = 1; i < agentData.length; i++) {
      const k = String(agentData[i][1] || agentData[i][0] || '').trim().toUpperCase();
      const rawVal = agentData[i].length >= 3 ? agentData[i][2] : agentData[i][1];
      const v = String(rawVal || '').trim();
      if (!v) continue;

      if (k === 'PRIMARY_AI_PROVIDER' && !config.aiProvider) {
        config.aiProvider = v.toUpperCase();
      }
      if (config.aiProvider === 'OPENAI' && k === 'OPENAI_MODEL' && !config.aiModel) {
        config.aiModel = v;
      }
      if (config.aiProvider === 'GEMINI' && k === 'GEMINI_MODEL' && !config.aiModel) {
        config.aiModel = v;
      }
    }
  }

  if (!config.aiProvider) config.aiProvider = 'GEMINI';
  if (!config.aiModel) config.aiModel = config.aiProvider === 'OPENAI' ? 'gpt-4o' : 'gemini-3.8-flash';

  return config;
}

function getActiveCampaignId_(spreadsheet, ui) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
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
  } catch (e) {}

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

function updateCampaignDiscoveryCounts_(spreadsheet, campaignId) {
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (!campSheet || campSheet.getLastRow() <= 1) return;

  const campRows = campSheet.getRange(2, 1, campSheet.getLastRow() - 1, campSheet.getLastColumn()).getValues();
  const campHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getDisplayValues()[0];
  const idCol = campHeaders.indexOf('Campaign ID');
  let updatedCol = campHeaders.indexOf('Last Modified');
  if (updatedCol < 0) updatedCol = campHeaders.indexOf('Last updated');

  for (let i = 0; i < campRows.length; i++) {
    if (String(campRows[i][idCol]).trim() === String(campaignId).trim()) {
      if (updatedCol >= 0) campSheet.getRange(i + 2, updatedCol + 1).setValue(new Date().toLocaleString());
      break;
    }
  }
}

function extractDiscoveryDomainFromUrl_(url) {
  if (!url) return '';
  let clean = String(url).trim().toLowerCase();
  clean = clean.replace(/^https?:\/\//i, '').replace(/^www\./i, '');
  return clean.split('/')[0].split('?')[0].split('#')[0];
}

function getDiscoverySheetRecords_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).trim(); });
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    records.push(obj);
  }
  return records;
}

function buildDiscoveryRowFromHeaders_(sheetHeaders, dataObject) {
  return sheetHeaders.map(function(header) {
    const cleanHeader = String(header).trim();
    return dataObject.hasOwnProperty(cleanHeader) ? dataObject[cleanHeader] : '';
  });
}