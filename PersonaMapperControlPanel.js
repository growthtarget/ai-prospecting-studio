/**
 * Opens the comprehensive Prospecting Control Panel in the sidebar.
 * Pre-loads campaign data on the server to prevent PERMISSION_DENIED storage errors.
 */
function showPersonaMapperControlPanel() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  
  // Pre-fetch campaign list directly within the spreadsheet context
  template.initialCampaigns = getAvailableCampaignsList();

  const html = template.evaluate()
    .setTitle('AI Prospecting Studio')
    .setWidth(340);
    
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns available campaigns for the dropdown from ALL possible tabs (Campaigns, Config, Settings).
 * Handles casing differences and eliminates duplicates.
 */
function getAvailableCampaignsList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const campaignMap = {};

  // Check 'Campaigns' sheet
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet) {
    const records = getSheetRecords_(campSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || r['Campaign id'] || r['campaign id'] || '').trim();
      const name = String(r['Campaign name'] || r['Campaign Name'] || r['campaign name'] || '').trim();
      if (id) campaignMap[id] = name || campaignMap[id] || '';
    });
  }

  // Check 'Campaign Configuration' sheet
  const confSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (confSheet) {
    const records = getSheetRecords_(confSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || r['Campaign id'] || '').trim();
      const name = String(r['Campaign Name'] || r['Campaign name'] || '').trim();
      if (id) campaignMap[id] = name || campaignMap[id] || '';
    });
  }

  // Check 'Persona Mapper Settings' sheet
  const setSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (setSheet) {
    const records = getSheetRecords_(setSheet);
    records.forEach(function(r) {
      const id = String(r['Campaign ID'] || '').trim();
      const name = String(r['Campaign Name'] || '').trim();
      if (id) campaignMap[id] = name || campaignMap[id] || '';
    });
  }

  return Object.keys(campaignMap).map(function(id) {
    return { id: id, name: campaignMap[id] };
  });
}

/**
 * Loads all targeting, ESG signals, AI model, and verification settings for a campaign.
 */
function getCampaignFullConfig(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const cleanId = String(campaignId || '').trim();
  const result = { id: cleanId };

  // Read from Campaign Configuration
  const confSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (confSheet) {
    const match = getSheetRecords_(confSheet).find(function(r) {
      return String(r['Campaign ID'] || '').trim() === cleanId;
    });
    if (match) {
      result.name = match['Campaign Name'] || match['Campaign name'] || '';
      result.country = match['Countries/regions'] || '';
      result.industry = match['Target industries'] || '';
      result.size = match['Company size'] || '';
      result.limit = match['Company limit'] || 5;
    }
  }

  // Read from Persona Mapper Settings
  const setSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (setSheet) {
    const match = getSheetRecords_(setSheet).find(function(r) {
      return String(r['Campaign ID'] || '').trim() === cleanId;
    });
    if (match) {
      result.name = result.name || match['Campaign Name'] || '';
      result.titles = match['Persona Titles'] || '';
      result.aiProvider = match['AI Provider'] || 'GEMINI';
      result.aiModel = match['AI Model'] || 'gemini-2.5-flash';
      result.emailVerifier = match['Email Verification Provider'] || 'ZeroBounce';
      result.emailEndpoint = match['Email Verification Endpoint'] || '';
      result.incEsgObligations = toBoolean_(match['Include ESG Reporting Obligations']);
      result.incScope3 = toBoolean_(match['Include Scope 3 Exposure']);
      result.incSupplyChain = toBoolean_(match['Include Supply Chain Complexity']);
      result.incDataChallenges = toBoolean_(match['Include Sustainability Data Challenges']);
      result.incDigital = toBoolean_(match['Include Digital Transformation']);
      result.incTriggers = toBoolean_(match['Include Business Triggers']);
      result.incEvaluater = toBoolean_(match['Include eValuater Solution Fit']);
    }
  }

  return result;
}

/**
 * Saves or updates settings across Campaign Configuration & Persona Mapper Settings simultaneously.
 */
function saveCampaignFullConfig(payload) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensurePersonaMapperSheets_(spreadsheet);

  const cid = String(payload.id).trim();
  const cname = String(payload.name).trim();

  // 1. Update 'Campaigns' summary sheet
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet) {
    upsertSimpleRow_(campSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Campaign name': cname,
      'Status': 'Active',
      'Last updated': new Date()
    });
  }

  // 2. Update 'Campaign Configuration' sheet
  const confSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (confSheet) {
    upsertSimpleRow_(confSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Campaign Name': cname,
      'Countries/regions': payload.country,
      'Target industries': payload.industry,
      'Company size': payload.size,
      'Company limit': payload.limit
    });
  }

  // 3. Update 'Persona Mapper Settings' sheet
  const setSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (setSheet) {
    upsertSimpleRow_(setSheet, 'Campaign ID', cid, {
      'Campaign ID': cid,
      'Campaign Name': cname,
      'Persona Mapper Enabled': 'TRUE',
      'AI Provider': payload.aiProvider,
      'AI Model': payload.aiModel,
      'Persona Titles': payload.titles,
      'Include ESG Reporting Obligations': payload.incEsgObligations ? 'TRUE' : 'FALSE',
      'Include Scope 3 Exposure': payload.incScope3 ? 'TRUE' : 'FALSE',
      'Include Supply Chain Complexity': payload.incSupplyChain ? 'TRUE' : 'FALSE',
      'Include Sustainability Data Challenges': payload.incDataChallenges ? 'TRUE' : 'FALSE',
      'Include Digital Transformation': payload.incDigital ? 'TRUE' : 'FALSE',
      'Include Business Triggers': payload.incTriggers ? 'TRUE' : 'FALSE',
      'Include eValuater Solution Fit': payload.incEvaluater ? 'TRUE' : 'FALSE',
      'Email Verification Provider': payload.emailVerifier,
      'Email Verification Endpoint': payload.emailEndpoint
    });
  }

  return { success: true, message: 'Settings saved successfully for ' + cid + '!' };
}

/**
 * Switches the user's active view to the Prospect Review Queue tab.
 */
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

/**
 * Helper to update or append a row in a sheet by key
 */
function upsertSimpleRow_(sheet, keyHeader, keyValue, record) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

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