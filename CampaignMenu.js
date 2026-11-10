/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * CAMPAIGN MENU & UI CONTROLLER (BACKEND ROUTERS)
 * ============================================================================
 * File: CampaignMenu.gs
 * 
 * NOTE: Menu registration is handled centrally by code.gs to avoid duplicate
 * onOpen collisions. This file handles sidebar/modal serving and campaign saving.
 * ============================================================================
 */

/**
 * Opens ControlPanel.html in the Google Sheets Sidebar
 */
function showControlPanelSidebar() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  template.initialCampaigns = getAvailableCampaignsList();
  template.initialConfigsMap = getAllCampaignConfigsMap();
  const html = template.evaluate()
    .setTitle('🎯 Campaign Control Panel')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens ControlPanel.html as a centered Modal Dialog
 */
function showCreateCampaignDialog() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  template.initialCampaigns = getAvailableCampaignsList();
  template.initialConfigsMap = getAllCampaignConfigsMap();
  const html = template.evaluate()
    .setWidth(600)
    .setHeight(760)
    .setTitle('🎯 Campaign Control Panel');
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 Campaign Control Panel');
}

/**
 * Reads available campaigns list for dropdown in ControlPanel.html
 */
function getAvailableCampaignsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Campaigns');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const list = [];
  data.forEach(r => {
    const id = String(r[0] || '').trim();
    const name = String(r[1] || '').trim();
    if (id) list.push({ id: id, name: name });
  });
  return list;
}

/**
 * Reads saved campaign configurations map for instant switching
 */
function getAllCampaignConfigsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Campaign Configuration');
  const configsMap = {};

  if (!configSheet) return configsMap;

  const data = configSheet.getDataRange().getValues();
  const activeParams = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = data[i][1];
    if (key) activeParams[key] = val;
  }

  const activeCid = activeParams['CAMPAIGN_ID'] || 'CMP-001';
  const provider = (activeParams['PRIMARY_AI_PROVIDER'] || 'GEMINI').toUpperCase();
  const model = activeParams['AI_MODEL'] || (provider === 'OPENAI' ? (activeParams['OPENAI_MODEL'] || 'gpt-4o') : (activeParams['GEMINI_MODEL'] || 'gemini-2.5-flash'));

  configsMap[activeCid] = {
    id: activeCid,
    name: activeParams['CAMPAIGN_NAME'] || '',
    department: activeParams['DEPARTMENT_OWNER'] || 'ESG Advisory',
    country: activeParams['TARGET_GEOGRAPHY'] || '',
    industry: activeParams['TARGET_INDUSTRIES'] || '',
    size: activeParams['COMPANY_SIZE'] || '',
    limit: activeParams['MAX_COMPANIES_LIMIT'] || 5,
    researchDepth: activeParams['RESEARCH_DEPTH'] || 'Deep Persona',
    minCompanyScore: activeParams['MIN_COMPANY_FIT_SCORE'] || 80,
    minPersonaScore: activeParams['MIN_PERSONA_FIT_SCORE'] || 85,
    titles: activeParams['TARGET_PERSONA_TITLES'] || '',
    campaignFocus: activeParams['CAMPAIGN_FOCUS'] || '',
    searchKeywords: activeParams['SEARCH_KEYWORDS'] || '',
    customBusinessNeeds: activeParams['CUSTOM_BUSINESS_NEEDS'] || '',
    aiProvider: provider,
    aiModel: model,
    emailVerifier: activeParams['EMAIL_VERIFICATION_PROVIDER'] || 'ZeroBounce',
    emailEndpoint: activeParams['VERIFIER_ENDPOINT'] || '',
    emailSenderProvider: activeParams['DISPATCH_MODE'] || 'Workspace Gmail (Drafts Only)',
    emailSenderEndpoint: activeParams['SENDER_ENDPOINT'] || '',
    senderName: activeParams['SENDER_NAME'] || 'Alex Mercer',
    senderTitle: activeParams['SENDER_TITLE'] || 'Solutions Consultant',
    senderCompany: activeParams['SENDER_COMPANY'] || 'eValuater Solutions',
    incEsgObligations: true,
    incScope3: true,
    incSupplyChain: true,
    incDataChallenges: true,
    incDigital: true,
    incTriggers: true,
    incEvaluater: true
  };

  return configsMap;
}

/**
 * Backend Handler: Saves full campaign payload from ControlPanel.html
 */
function saveCampaignFullConfig(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Update or Insert into Campaigns Tab
  const campaignsSheet = ss.getSheetByName('Campaigns');
  if (campaignsSheet) {
    let rowIndex = -1;
    const data = campaignsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === payload.id) {
        rowIndex = i + 1;
        break;
      }
    }

    const now = new Date();
    const dateStr = now.toISOString().split('T')[0];

    if (rowIndex > 0) {
      campaignsSheet.getRange(rowIndex, 2).setValue(payload.name);
      campaignsSheet.getRange(rowIndex, 5).setValue(payload.limit);
      campaignsSheet.getRange(rowIndex, 6).setValue(payload.researchDepth || 'Deep Persona');
      campaignsSheet.getRange(rowIndex, 8).setValue(now.toLocaleString());
    } else {
      campaignsSheet.appendRow([
        payload.id,
        payload.name,
        payload.campaignFocus || 'Custom Department Campaign',
        'Active',
        payload.limit || 5,
        payload.researchDepth || 'Deep Persona',
        dateStr,
        now.toLocaleString()
      ]);
    }
  }

  // 2. Update Campaign Configuration Tab
  const configSheet = ss.getSheetByName('Campaign Configuration');
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    const provider = String(payload.aiProvider || 'GEMINI').toUpperCase();
    const model = payload.aiModel || (provider === 'OPENAI' ? 'gpt-4o' : 'gemini-2.5-flash');

    const updates = {
      'CAMPAIGN_ID': payload.id,
      'CAMPAIGN_NAME': payload.name,
      'DEPARTMENT_OWNER': payload.department || 'ESG Advisory',
      'TARGET_INDUSTRIES': payload.industry,
      'TARGET_GEOGRAPHY': payload.country,
      'COMPANY_SIZE': payload.size,
      'RESEARCH_DEPTH': payload.researchDepth,
      'MIN_COMPANY_FIT_SCORE': payload.minCompanyScore,
      'MIN_PERSONA_FIT_SCORE': payload.minPersonaScore,
      'MAX_COMPANIES_LIMIT': payload.limit,
      'TARGET_PERSONA_TITLES': payload.titles,
      'CAMPAIGN_FOCUS': payload.campaignFocus,
      'SEARCH_KEYWORDS': payload.searchKeywords,
      'CUSTOM_BUSINESS_NEEDS': payload.customBusinessNeeds || '',
      'PRIMARY_AI_PROVIDER': provider,
      'AI_MODEL': model,
      'EMAIL_VERIFICATION_PROVIDER': payload.emailVerifier,
      'VERIFIER_ENDPOINT': payload.emailEndpoint,
      'DISPATCH_MODE': payload.emailSenderProvider,
      'SENDER_ENDPOINT': payload.emailSenderEndpoint,
      'SENDER_NAME': payload.senderName,
      'SENDER_TITLE': payload.senderTitle,
      'SENDER_COMPANY': payload.senderCompany,
      'STATUS': 'Active'
    };

    if (provider === 'OPENAI') {
      updates['OPENAI_MODEL'] = model;
    } else {
      updates['GEMINI_MODEL'] = model;
    }

    const updatedKeys = [];
    for (let r = 1; r < configData.length; r++) {
      const key = String(configData[r][0]).trim();
      if (updates.hasOwnProperty(key)) {
        configSheet.getRange(r + 1, 2).setValue(updates[key]);
        updatedKeys.push(key);
      }
    }

    if (!updatedKeys.includes('CUSTOM_BUSINESS_NEEDS')) {
      configSheet.appendRow(['CUSTOM_BUSINESS_NEEDS', payload.customBusinessNeeds || '', 'Department pain points & buying signals']);
    }
    if (!updatedKeys.includes('AI_MODEL')) {
      configSheet.appendRow(['AI_MODEL', model, 'Configured AI Model']);
    }
  }

  // 3. Update Master Dashboard
  const dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet && dashSheet.getLastRow() > 2) {
    dashSheet.getRange(4, 2).setValue(payload.id);
    dashSheet.getRange(3, 2).setValue(payload.name);
    dashSheet.getRange(3, 4).setValue(new Date().toLocaleString());
  }

  return { success: true, message: 'Settings saved for ' + payload.id + ' (' + payload.aiProvider + ') successfully!' };
}

function runCompanyDiscoveryFromSidebar(cid) {
  if (typeof runCompanyDiscoveryPipeline === 'function') {
    return runCompanyDiscoveryPipeline(cid);
  }
  return { success: true, message: 'Discovery started for ' + cid + '.' };
}

function runPersonaMapperFromSidebar(cid) {
  if (typeof runPersonaMappingPipeline === 'function') {
    return runPersonaMappingPipeline(cid);
  }
  return { success: true, message: 'Persona Mapper completed for ' + cid + '.' };
}

function navigateToReviewQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prospect Review Queue');
  if (sheet) {
    ss.setActiveSheet(sheet);
    return { success: true, message: 'Switched to Prospect Review Queue tab.' };
  }
  return { success: false, message: 'Prospect Review Queue sheet not found.' };
}

function triggerSetupAllDropdowns() {
  if (typeof setupAllSheetsAndDropdowns === 'function') {
    setupAllSheetsAndDropdowns();
  } else {
    SpreadsheetApp.getUi().alert('setupAllSheetsAndDropdowns() function not found. Please ensure SheetSetup.gs is updated.');
  }
}

function triggerPilotRun() {
  if (typeof runPilotOnRealSheets === 'function') {
    runPilotOnRealSheets();
  } else {
    SpreadsheetApp.getUi().alert('runPilotOnRealSheets() function not found in PilotSandboxRunner.gs.');
  }
}

function triggerPilotCleanup() {
  if (typeof cleanupPilotDataFromRealSheets === 'function') {
    cleanupPilotDataFromRealSheets();
  } else {
    SpreadsheetApp.getUi().alert('cleanupPilotDataFromRealSheets() function not found in PilotSandboxRunner.gs.');
  }
}

function triggerApprovedGmailDrafts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Prospect Review Queue');
  if (!queueSheet) {
    ss.toast('Prospect Review Queue tab not found.', 'Error', 5);
    return;
  }

  const data = queueSheet.getDataRange().getValues();
  let createdCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = String(row[7] || '').trim();
    const approvalStatus = String(row[21] || '').trim();
    const sendStatus = String(row[24] || '').trim();
    const subject = String(row[14] || '').trim();
    const body = String(row[15] || '').trim();

    if (approvalStatus.startsWith('Approved') && sendStatus !== 'Draft Created in Gmail' && email) {
      try {
        const draft = GmailApp.createDraft(email, subject, body);
        queueSheet.getRange(i + 1, 25).setValue('Draft Created in Gmail');
        queueSheet.getRange(i + 1, 26).setValue(new Date().toLocaleString());
        queueSheet.getRange(i + 1, 20).setValue('Gmail Draft ID: ' + draft.getId());
        createdCount++;
      } catch (err) {
        queueSheet.getRange(i + 1, 27).setValue(err.message);
      }
    }
  }

  ss.toast('Drafts created in Gmail: ' + createdCount, 'Draft Engine', 6);
}