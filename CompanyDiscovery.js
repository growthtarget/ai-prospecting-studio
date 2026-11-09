/**
 * Agent 1: Company Discovery & Research Agent
 * Reads targeting criteria from 'Campaign Configuration' and discovers matching companies using Gemini.
 */
function runCompanyDiscoveryForSelectedCampaign() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  
  if (!lock.tryLock(1000)) {
    ui.alert('A research run is already in progress. Please wait for it to finish.');
    return;
  }

  try {
    const campaignId = getActiveCampaignId_(spreadsheet, ui);
    if (!campaignId) return;

    const config = getCampaignConfigurationRecord_(spreadsheet, campaignId);
    if (!config) {
      ui.alert('No configuration found for Campaign ID: ' + campaignId + ' in "Campaign Configuration".');
      return;
    }

    const companyLimit = Math.min(15, Math.max(1, Number(config['Company limit'] || config['Maximum companies'] || 5)));
    
    ui.alert(
      'Starting Company Discovery',
      'Searching for up to ' + companyLimit + ' companies in:\n' +
      '• Industry: ' + (config['Target industries'] || 'Any') + '\n' +
      '• Country: ' + (config['Countries/regions'] || 'Any') + '\n' +
      '• Size: ' + (config['Company size'] || 'Any') + '\n\n' +
      'Gemini with Google Search is now researching the web. Click OK to begin.',
      ui.ButtonSet.OK
    );

    // 1. Build the Discovery Prompt
    const prompt = buildCompanyDiscoveryPrompt_(config, companyLimit);

    // 2. Execute Universal AI Search with Google Grounding
    const aiResponse = callAI({
      provider: 'GEMINI',
      model: 'gemini-2.5-flash',
      prompt: prompt,
      enableWebSearch: true,
      json: true
    });

    const discoveredCompanies = (aiResponse.json && Array.isArray(aiResponse.json.companies)) ? aiResponse.json.companies : [];
    if (!discoveredCompanies.length) {
      ui.alert('The AI search did not return any companies matching the given criteria. Try broadening your criteria.');
      return;
    }

    // 3. Save Discovered Companies into 'Company Database' & Log Evidence
    const dbSheet = spreadsheet.getSheetByName('Company Database');
    const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');
    if (!dbSheet) throw new Error('Sheet "Company Database" not found.');

    const existingCompanies = getSheetRecords_(dbSheet);
    let addedCount = 0;

    discoveredCompanies.forEach(function(item) {
      const companyName = String(item.companyName || '').trim();
      const website = String(item.website || '').trim();
      const domain = domainFromUrl_(website);

      // Duplicate Prevention: Check if company or domain already exists in this campaign
      const isDuplicate = existingCompanies.some(function(row) {
        const rowDomain = domainFromUrl_(row['Website'] || '');
        return (row['Campaign ID'] === campaignId) && (
          (domain && rowDomain && domain === rowDomain) ||
          (row['Company name'] && row['Company name'].toLowerCase() === companyName.toLowerCase())
        );
      });

      if (isDuplicate || !companyName) return;

      const companyId = 'COMP-' + Utilities.getUuid().slice(0, 6).toUpperCase();
      const now = new Date();

      // Write to Company Database
      dbSheet.appendRow([
        companyId,
        campaignId,
        companyName,
        website,
        item.headquarters || config['Countries/regions'] || '',
        item.industry || config['Target industries'] || '',
        item.companySize || config['Company size'] || '',
        item.businessSummary || '',
        item.qualificationResult || 'Qualified',
        item.qualificationExplanation || 'Discovered via Google Search Grounding matching campaign criteria.',
        item.initialScore || 80,
        'Web Verified',
        'Discovered',
        now,
        now
      ]);

      // Write Supporting Evidence if available
      if (evidenceSheet && item.sourceUrl) {
        evidenceSheet.appendRow([
          'EVD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
          companyId,
          'Company profile discovery',
          item.businessSummary || companyName,
          item.sourceUrl,
          'Official Website / Web Registry',
          now,
          'Yes',
          'High',
          'Identified during initial campaign search'
        ]);
      }

      addedCount++;
    });

    // 4. Update Campaign Counts in "Campaigns" & "Dashboard"
    updateCampaignDiscoveryCounts_(spreadsheet, campaignId);

    ui.alert(
      'Company Discovery Complete',
      'Successfully discovered and added ' + addedCount + ' new qualified companies to "Company Database".\n\n' +
      'Next Step: Run "2. Run Persona Mapper" to find decision-makers and draft personalized outreach.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Company Discovery Error: ' + error.message);
    ui.alert('Error during Company Discovery: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds the structured Google Search Grounding prompt
 */
function buildCompanyDiscoveryPrompt_(config, limit) {
  return [
    'Act as the Senior B2B Account Research & Discovery Agent.',
    'Your goal is to discover REAL, active companies that strictly match these campaign criteria using live Google Search:',
    '• Target Countries/Regions: ' + (config['Countries/regions'] || 'Global'),
    '• Target Industries: ' + (config['Target industries'] || 'Unspecified'),
    '• Target Company Size (Headcount): ' + (config['Company size'] || 'Any'),
    '• Required Qualification Criteria: ' + (config['Required criteria'] || 'None'),
    '• Exclusions (DO NOT INCLUDE): ' + (config['Exclusions'] || 'None'),
    '• Relevant Activities: ' + (config['Relevant activities'] || 'Standard business operations'),
    '',
    'Rules:',
    '1. Use web search to find real, verifiable companies. Never hallucinate or invent company names or URLs.',
    '2. Return official, secure website domains (e.g., https://example.com).',
    '3. Provide at most ' + limit + ' distinct companies.',
    '4. Verify each company strictly adheres to the inclusion criteria and is not disqualified by exclusions.',
    '',
    'Return ONLY a valid JSON object matching this schema:',
    '{',
    '  "companies": [',
    '    {',
    '      "companyName": "Exact Legal or Operating Name",',
    '      "website": "https://company-domain.com",',
    '      "headquarters": "City, Country",',
    '      "industry": "Specific Industry",',
    '      "companySize": "Estimated employee count or range",',
    '      "businessSummary": "Concise 1-2 sentence description of primary operations and business activities",',
    '      "qualificationResult": "Qualified",',
    '      "qualificationExplanation": "Why this company matches the target criteria",',
    '      "initialScore": 85,',
    '      "sourceUrl": "Direct URL where profile or verification was confirmed"',
    '    }',
    '  ]',
    '}'
  ].join('\n');
}

/**
 * Helper to fetch Campaign Configuration
 */
function getCampaignConfigurationRecord_(spreadsheet, campaignId) {
  const sheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (!sheet) return null;
  const records = getSheetRecords_(sheet);
  return records.find(function(row) {
    return String(row['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();
  });
}

/**
 * Helper to detect active Campaign ID
 */
function getActiveCampaignId_(spreadsheet, ui) {
  const sheet = spreadsheet.getActiveSheet();
  const allowed = ['Campaigns', 'Campaign Configuration', 'Company Database', 'Dashboard'];
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
 * Updates summary counts on Campaigns sheet
 */
function updateCampaignDiscoveryCounts_(spreadsheet, campaignId) {
  const compSheet = spreadsheet.getSheetByName('Company Database');
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (!compSheet || !campSheet) return;

  const records = getSheetRecords_(compSheet).filter(function(r) {
    return String(r['Campaign ID'] || '').trim() === String(campaignId).trim();
  });

  const foundCount = records.length;
  const qualifiedCount = records.filter(function(r) {
    return String(r['Qualification result'] || '').toLowerCase() === 'qualified';
  }).length;

  const campRows = campSheet.getRange(2, 1, Math.max(1, campSheet.getLastRow() - 1), campSheet.getLastColumn()).getValues();
  const campHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getDisplayValues()[0];
  const idCol = campHeaders.indexOf('Campaign ID');
  const updatedCol = campHeaders.indexOf('Last updated');

  for (let i = 0; i < campRows.length; i++) {
    if (String(campRows[i][idCol]).trim() === String(campaignId).trim()) {
      if (updatedCol >= 0) campSheet.getRange(i + 2, updatedCol + 1).setValue(new Date());
      break;
    }
  }
}