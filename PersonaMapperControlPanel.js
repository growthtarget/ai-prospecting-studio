/**
 * Opens the comprehensive Prospecting Control Panel in the sidebar.
 */
function showPersonaMapperControlPanel() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  
  // Pre-load all campaigns and criteria server-side
  const configsMap = getAllCampaignConfigsMap_();
  template.initialCampaigns = Object.keys(configsMap).map(function(id) {
    return { id: id, name: configsMap[id].name || id };
  });
  template.initialConfigsMap = configsMap;

  const html = template.evaluate()
    .setTitle('AI Prospecting Studio')
    .setWidth(340);
    
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Saves and OVERWRITES all unique criteria entered in the sidebar across all sheets.
 */
function saveCampaignFullConfig(payload) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensurePersonaMapperSheets_(spreadsheet);

  const cid = String(payload.id).trim();
  const cname = String(payload.name).trim();

  // 1. Update 'Campaigns' sheet
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet) {
    upsertSimpleRow_(campSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Campaign name': cname,
      'Objective': 'Active Lead Generation',
      'Status': 'Active',
      'Company limit': payload.limit || 5,
      'Research depth': 'Standard',
      'Last updated': new Date()
    });
  }

  // 2. Overwrite 'Campaign Configuration' sheet with unique criteria
  const confSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (confSheet) {
    upsertSimpleRow_(confSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Target industries': payload.industry,
      'Countries/regions': payload.country,
      'Company size': payload.size,
      'Relevant activities': 'Commercial operations and corporate reporting',
      'Required criteria': 'Must match specified industry and country filters',
      'Exclusions': 'Consulting agencies, pure SaaS software',
      'Research instructions': 'Identify verified corporate headquarters and executive team',
      'Evidence expectations': 'Official corporate website and public registries',
      'Scoring criteria and weights': 'Size fit: 40%, Industry fit: 60%',
      'Required output fields': 'Company Name, Website, Country, Size',
      'Selected model': payload.aiModel || 'gemini-2.5-flash'
    });
  }

  // 3. Overwrite 'Persona Mapper Settings' sheet with unique criteria
  const setSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (setSheet) {
    upsertSimpleRow_(setSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Campaign Name': cname,
      'Persona Mapper Enabled': 'TRUE',
      'AI Provider': payload.aiProvider || 'GEMINI',
      'AI Model': payload.aiModel || 'gemini-2.5-flash',
      'Persona Titles': payload.titles,
      'Search Keywords': 'CSRD, Scope 3, supply chain, emissions',
      'Campaign Focus': 'Supply chain compliance and data automation',
      'Evidence Period Months': 12,
      'Research Depth': 'Standard',
      'Include ESG Reporting Obligations': payload.incEsgObligations ? 'TRUE' : 'FALSE',
      'Include ESG Reporting Maturity': payload.incEsgObligations ? 'TRUE' : 'FALSE',
      'Include Supply Chain Complexity': payload.incSupplyChain ? 'TRUE' : 'FALSE',
      'Include Scope 3 Exposure': payload.incScope3 ? 'TRUE' : 'FALSE',
      'Include Sustainability Data Challenges': payload.incDataChallenges ? 'TRUE' : 'FALSE',
      'Include Digital Transformation': payload.incDigital ? 'TRUE' : 'FALSE',
      'Include ESG Governance': 'TRUE',
      'Include Business Triggers': payload.incTriggers ? 'TRUE' : 'FALSE',
      'Include eValuater Solution Fit': payload.incEvaluater ? 'TRUE' : 'FALSE',
      'Weight ESG Reporting Relevance': 20,
      'Weight Supply Chain Complexity': 20,
      'Weight Sustainability Data Needs': 20,
      'Weight Digital Transformation Relevance': 15,
      'Weight Business Triggers': 15,
      'Weight eValuater Solution Fit': 10,
      'Minimum Company Opportunity Score': 70,
      'Minimum Persona Confidence': 'Medium',
      'Minimum Email Priority Score': 70,
      'Human Approval Required': 'TRUE',
      'Auto Send Enabled': 'FALSE',
      'Email Priority Weight Company Score': 50,
      'Email Priority Weight Persona Fit': 30,
      'Email Priority Weight Buying Signal': 20,
      'Notes': 'Saved via Sidebar Control Panel',
      'Max Companies Per Run': payload.limit || 5,
      'Email Verification Provider': payload.emailVerifier || 'ZeroBounce',
      'Email Verification Endpoint': payload.emailEndpoint || 'https://api.zerobounce.net/v2/validate'
    });
  }

  return { success: true, message: 'Settings saved and updated for ' + cid + '!' };
}

/**
 * Reads all campaign data across all sheets into a single server-side bundle.
 */
function getAllCampaignConfigsMap_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configMap = {};

  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet) {
    const records = getSheetRecords_(campSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || r['Campaign id'] || '').trim();
      const name = String(r['Campaign name'] || r['Campaign Name'] || '').trim();
      if (id) {
        if (!configMap[id]) configMap[id] = { id: id };
        configMap[id].name = name;
      }
    });
  }

  const confSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (confSheet) {
    const records = getSheetRecords_(confSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || r['Campaign id'] || '').trim();
      if (id) {
        if (!configMap[id]) configMap[id] = { id: id };
        configMap[id].name = configMap[id].name || r['Campaign Name'] || r['Campaign name'] || '';
        configMap[id].country = r['Countries/regions'] || '';
        configMap[id].industry = r['Target industries'] || '';
        configMap[id].size = r['Company size'] || '';
        configMap[id].limit = r['Company limit'] || 5;
      }
    });
  }

  const setSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (setSheet) {
    const records = getSheetRecords_(setSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || '').trim();
      if (id) {
        if (!configMap[id]) configMap[id] = { id: id };
        configMap[id].name = configMap[id].name || r['Campaign Name'] || '';
        configMap[id].titles = r['Persona Titles'] || '';
        configMap[id].aiProvider = r['AI Provider'] || 'GEMINI';
        configMap[id].aiModel = r['AI Model'] || 'gemini-2.5-flash';
        configMap[id].emailVerifier = r['Email Verification Provider'] || 'ZeroBounce';
        configMap[id].emailEndpoint = r['Email Verification Endpoint'] || '';
        configMap[id].incEsgObligations = toBoolean_(r['Include ESG Reporting Obligations']);
        configMap[id].incScope3 = toBoolean_(r['Include Scope 3 Exposure']);
        configMap[id].incSupplyChain = toBoolean_(r['Include Supply Chain Complexity']);
        configMap[id].incDataChallenges = toBoolean_(r['Include Sustainability Data Challenges']);
        configMap[id].incDigital = toBoolean_(r['Include Digital Transformation']);
        configMap[id].incTriggers = toBoolean_(r['Include Business Triggers']);
        configMap[id].incEvaluater = toBoolean_(r['Include eValuater Solution Fit']);
      }
    });
  }

  return configMap;
}

function getAvailableCampaignsList() {
  const configsMap = getAllCampaignConfigsMap_();
  return Object.keys(configsMap).map(function(id) {
    return { id: id, name: configsMap[id].name || id };
  });
}

function navigateToReviewQueueSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensurePersonaMapperSheets_(spreadsheet);

  let sheet = spreadsheet.getSheetByName('Prospect Review Queue') || 
              spreadsheet.getSheetByName('Email Review Queue');

  if (sheet) {
    spreadsheet.setActiveSheet(sheet);
    sheet.activate();
    SpreadsheetApp.flush();
    return { success: true, message: 'Switched to Prospect Review Queue.' };
  }
  return { success: false, message: 'Prospect Review Queue sheet not found.' };
}

function upsertSimpleRow_(sheet, keyHeader, keyValue, record) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  const keyIndex = headers.indexOf(keyHeader);
  if (keyIndex < 0) return;

  const dataRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  let targetRow = 0;

  for (let i = 0; i < dataRows.length; i++) {
    if (String(dataRows[i][keyIndex] || '').trim().toLowerCase() === String(keyValue).trim().toLowerCase()) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow > 0) {
    const existing = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
    headers.forEach(function(header, idx) {
      if (record[header] !== undefined) existing[idx] = record[header];
    });
    sheet.getRange(targetRow, 1, 1, lastCol).setValues([existing]);
  } else {
    const newRow = headers.map(function(header) {
      return record[header] !== undefined ? record[header] : '';
    });
    sheet.appendRow(newRow);
  }
}