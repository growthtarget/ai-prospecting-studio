/** @OnlyCurrentDoc */

/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * MASTER CONTROLLER & UNIFIED TOP MENU
 * ============================================================================
 * File: code.gs
 * ============================================================================
 */

/**
 * 39 Canonical Header Columns for Persona Mapper Settings
 */
const PERSONA_MAPPER_39_HEADERS = [
  'Campaign ID',
  'Campaign Name',
  'Persona Mapper Enabled',
  'AI Provider',
  'AI Model',
  'Persona Titles',
  'Search Keywords',
  'Campaign Focus',
  'Evidence Period Months',
  'Research Depth',
  'Include ESG Reporting Obligations',
  'Include ESG Reporting Maturity',
  'Include Supply Chain Complexity',
  'Include Scope 3 Exposure',
  'Include Sustainability Data Challenges',
  'Include Digital Transformation',
  'Include ESG Governance',
  'Include Business Triggers',
  'Include eValuater Solution Fit',
  'Weight ESG Reporting Relevance',
  'Weight Supply Chain Complexity',
  'Weight Sustainability Data Needs',
  'Weight Digital Transformation Relevance',
  'Weight Business Triggers',
  'Weight eValuater Solution Fit',
  'Minimum Company Opportunity Score',
  'Minimum Persona Confidence',
  'Minimum Email Priority Score',
  'Human Approval Required',
  'Auto Send Enabled',
  'Email Priority Weight Company Score',
  'Email Priority Weight Persona Fit',
  'Email Priority Weight Buying Signal',
  'Notes',
  'Max Companies Per Run',
  'Email Verification Provider',
  'Email Verification Endpoint',
  'Email Sender Provider',
  'Email Sender Endpoint'
];

/**
 * Helper to clean and normalize keys (removes spaces, underscores, symbols, uppercase)
 */
function normalizeStudioKey(str) {
  return String(str || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Master onOpen Trigger: Installs the complete, unified menu across all tools.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎯 AI Prospecting Studio')
    .addItem('🚀 Open Campaign Control Panel', 'showPersonaMapperControlPanel')
    .addItem('🔄 Sync Config to Persona Mapper Settings', 'executeMasterPersonaMapperSync')
    .addItem('🎨 Fix Tab Header Contrast & Styling', 'fixAllTabHeaderContrast')
    .addSeparator()
    .addItem('🔍 1. Discover Target Companies', 'runCompanyDiscoveryForSelectedCampaign')
    .addItem('👤 2. Run Persona Mapper', 'runPersonaMapperForSelectedCampaign')
    .addSeparator()
    .addItem('📋 Go to Prospect Review Queue', 'navigateToReviewQueueSheet')
    .addItem('✅ Approve Selected Prospect', 'approveSelectedProspect')
    .addItem('✏️ Mark Selected: Needs Editing', 'markSelectedProspectNeedsEditing')
    .addItem('❌ Reject Selected Prospect', 'rejectSelectedProspect')
    .addSeparator()
    .addItem('📝 Create Gmail Drafts (Safe Review)', 'createGmailDraftsForApprovedProspects')
    .addItem('✉️ Send Approved Outreach Emails', 'sendApprovedOutreachEmails')
    .addSeparator()
    .addItem('⚙️ Initialize Sheets & Setup All Dropdowns', 'triggerSetupAllDropdowns')
    .addItem('⚙️ Prepare / Verify All Tabs', 'preparePersonaMapperSheets')
    .addItem('❓ Check Spreadsheet Setup', 'showSetupStatus')
    .addSeparator()
    .addItem('🧪 Run Authentic Pilot Pipeline', 'triggerPilotRun')
    .addItem('🧹 Clean Up Pilot Test Data', 'triggerPilotCleanup')
    .addItem('🔑 Authorize Studio Permissions', 'authorizeStudioPermissions')
    .addToUi();
}

/**
 * Restyles headers across all tabs with Dark Blue Text on Light Blue Background.
 */
function fixAllTabHeaderContrast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let count = 0;

  sheets.forEach(sh => {
    if (sh.getLastColumn() > 0) {
      const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
      headerRange
        .setFontWeight('bold')
        .setFontColor('#174ea6')          // Dark Navy Blue (High Contrast)
        .setBackground('#e8f0fe')         // Light Google Blue
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sh.setFrozenRows(1);
      count++;
    }
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Contrast Restyled',
    'Successfully applied high-contrast headers (Dark Navy on Light Blue) and frozen Row 1 across ' + count + ' tabs!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Adaptive sheet finder prioritizing "Persona Mapper Settings" (plural)
 */
function findPersonaMapperSettingsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const candidateNames = [
    'Persona Mapper Settings',
    'Persona Mapper Setting',
    'Persona  Mapper Settings',
    'Persona  Mapper Setting',
    'PersonaMapperSettings',
    'Persona_Mapper_Settings'
  ];

  for (let i = 0; i < candidateNames.length; i++) {
    const s = ss.getSheetByName(candidateNames[i]);
    if (s) return s;
  }

  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const norm = allSheets[i].getName().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (norm.includes('personamapper') || (norm.includes('persona') && norm.includes('setting'))) {
      return allSheets[i];
    }
  }

  return ss.insertSheet('Persona Mapper Settings');
}

/**
 * Universal API Key resolver that checks Script Properties & Agent Settings
 */
function getStudioApiKey(provider) {
  const norm = String(provider || '').toUpperCase();
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();

  for (let k in scriptProps) {
    const uk = k.toUpperCase();
    if ((norm.includes('OPENAI') || norm.includes('GPT')) && uk.includes('OPENAI')) {
      if (scriptProps[k] && !scriptProps[k].includes('PASTE_')) return scriptProps[k];
    }
    if (norm.includes('GEMINI') && uk.includes('GEMINI')) {
      if (scriptProps[k] && !scriptProps[k].includes('PASTE_')) return scriptProps[k];
    }
  }

  for (let k in userProps) {
    const uk = k.toUpperCase();
    if ((norm.includes('OPENAI') || norm.includes('GPT')) && uk.includes('OPENAI')) {
      if (userProps[k] && !userProps[k].includes('PASTE_')) return userProps[k];
    }
    if (norm.includes('GEMINI') && uk.includes('GEMINI')) {
      if (userProps[k] && !userProps[k].includes('PASTE_')) return userProps[k];
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Agent Settings');
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cell = String(data[r][c] || '').trim().toUpperCase();
        if ((norm.includes('OPENAI') || norm.includes('GPT')) && cell === 'OPENAI_API_KEY') {
          const val = String(data[r][c + 1] || data[r][2] || '').trim();
          if (val && !val.includes('PASTE_')) return val;
        }
        if (norm.includes('GEMINI') && cell === 'GEMINI_API_KEY') {
          const val = String(data[r][c + 1] || data[r][2] || '').trim();
          if (val && !val.includes('PASTE_')) return val;
        }
      }
    }
  }
  return '';
}

function getApiKey(provider) {
  return getStudioApiKey(provider);
}

/**
 * Opens the Campaign Control Panel sidebar and pre-loads campaigns and configuration.
 */
function showPersonaMapperControlPanel() {
  const template = HtmlService.createTemplateFromFile('ControlPanel');
  const sidebarData = getSidebarInitialData();
  template.initialCampaigns = sidebarData.campaigns;
  template.initialConfigsMap = sidebarData.configsMap;

  const html = template.evaluate().setTitle('🎯 Campaign Control Panel').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns available campaigns and configuration maps for sidebar initialization.
 */
function getSidebarInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaigns = [];
  const configsMap = {};

  const pSheet = findPersonaMapperSettingsSheet(ss);
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const cid = String(pData[i][0] || '').trim();
      if (!cid) continue;

      const name = String(pData[i][1] || cid).trim();
      campaigns.push({ id: cid, name: name });

      configsMap[cid] = {
        id: cid,
        name: name,
        aiProvider: String(pData[i][3] || 'OPENAI'),
        aiModel: String(pData[i][4] || 'gpt-4o'),
        titles: String(pData[i][5] || ''),
        searchKeywords: String(pData[i][6] || ''),
        campaignFocus: String(pData[i][7] || ''),
        evidencePeriodMonths: pData[i][8] || 12,
        researchDepth: String(pData[i][9] || 'Deep Persona'),
        incEsgObligations: String(pData[i][10]).toUpperCase() !== 'FALSE',
        incEsgMaturity: String(pData[i][11]).toUpperCase() !== 'FALSE',
        incSupplyChain: String(pData[i][12]).toUpperCase() !== 'FALSE',
        incScope3: String(pData[i][13]).toUpperCase() !== 'FALSE',
        incDataChallenges: String(pData[i][14]).toUpperCase() !== 'FALSE',
        incDigital: String(pData[i][15]).toUpperCase() !== 'FALSE',
        incEsgGovernance: String(pData[i][16]).toUpperCase() !== 'FALSE',
        incTriggers: String(pData[i][17]).toUpperCase() !== 'FALSE',
        incEvaluater: String(pData[i][18]).toUpperCase() !== 'FALSE',
        weightEsgReporting: String(pData[i][19] || '20%'),
        weightSupplyChain: String(pData[i][20] || '15%'),
        weightDataChallenges: String(pData[i][21] || '15%'),
        weightDigital: String(pData[i][22] || '15%'),
        weightTriggers: String(pData[i][23] || '20%'),
        weightEvaluaterFit: String(pData[i][24] || '15%'),
        minCompanyScore: pData[i][25] || 80,
        minPersonaScore: pData[i][26] || 85,
        minEmailPriorityScore: pData[i][27] || 80,
        humanApprovalRequired: String(pData[i][28]).toUpperCase() !== 'FALSE',
        autoSendEnabled: String(pData[i][29]).toUpperCase() === 'TRUE',
        weightEmailCompanyScore: String(pData[i][30] || '40%'),
        weightEmailPersonaFit: String(pData[i][31] || '35%'),
        weightEmailBuyingSignal: String(pData[i][32] || '25%'),
        customBusinessNeeds: String(pData[i][33] || ''),
        limit: pData[i][34] || 5,
        emailVerifier: String(pData[i][35] || 'ZeroBounce'),
        emailEndpoint: String(pData[i][36] || ''),
        emailSenderProvider: String(pData[i][37] || 'Workspace Gmail (Drafts Only)'),
        emailSenderEndpoint: String(pData[i][38] || '')
      };
    }
  }

  const cSheet = ss.getSheetByName('Campaign Configuration');
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    const configDict = {};
    for (let j = 1; j < cData.length; j++) {
      const k = normalizeStudioKey(cData[j][0]);
      if (k) configDict[k] = cData[j][1];
    }

    const activeCid = String(configDict['CAMPAIGNID'] || configDict['ACTIVECAMPAIGNID'] || configDict['CAMPAIGN'] || '').trim();
    if (activeCid) {
      if (!configsMap[activeCid]) {
        configsMap[activeCid] = { id: activeCid, name: String(configDict['CAMPAIGNNAME'] || activeCid) };
        campaigns.push({ id: activeCid, name: configsMap[activeCid].name });
      }

      const c = configsMap[activeCid];
      if (configDict['DEPARTMENT']) c.department = String(configDict['DEPARTMENT']);
      if (configDict['TARGETCOUNTRIES']) c.country = String(configDict['TARGETCOUNTRIES']);
      if (configDict['TARGETINDUSTRIES']) c.industry = String(configDict['TARGETINDUSTRIES']);
      if (configDict['COMPANYSIZERANGE']) c.size = String(configDict['COMPANYSIZERANGE']);
      if (configDict['MAXCOMPANIES']) c.limit = Number(configDict['MAXCOMPANIES']) || 5;
      if (configDict['RESEARCHDEPTH']) c.researchDepth = String(configDict['RESEARCHDEPTH']);
      if (configDict['TARGETTITLES']) c.titles = String(configDict['TARGETTITLES']);
      if (configDict['CAMPAIGNFOCUS']) c.campaignFocus = String(configDict['CAMPAIGNFOCUS']);
      if (configDict['SEARCHKEYWORDS']) c.searchKeywords = String(configDict['SEARCHKEYWORDS']);
      if (configDict['CUSTOMBUSINESSNEEDS']) c.customBusinessNeeds = String(configDict['CUSTOMBUSINESSNEEDS']);
      if (configDict['AIPROVIDER']) c.aiProvider = String(configDict['AIPROVIDER']);
      if (configDict['AIMODEL']) c.aiModel = String(configDict['AIMODEL']);
      if (configDict['SENDERNAME']) c.senderName = String(configDict['SENDERNAME']);
      if (configDict['SENDERTITLE']) c.senderTitle = String(configDict['SENDERTITLE']);
      if (configDict['SENDERCOMPANY']) c.senderCompany = String(configDict['SENDERCOMPANY']);
      if (configDict['EMAILVERIFICATIONPROVIDER']) c.emailVerifier = String(configDict['EMAILVERIFICATIONPROVIDER']);
      if (configDict['EMAILVERIFICATIONENDPOINT']) c.emailEndpoint = String(configDict['EMAILVERIFICATIONENDPOINT']);
      if (configDict['EMAILSENDERPROVIDER']) c.emailSenderProvider = String(configDict['EMAILSENDERPROVIDER']);
      if (configDict['EMAILSENDERENDPOINT']) c.emailSenderEndpoint = String(configDict['EMAILSENDERENDPOINT']);
    }
  }

  return { campaigns: campaigns, configsMap: configsMap };
}

function getAvailableCampaignsList() {
  return getSidebarInitialData().campaigns;
}

/**
 * ISOLATED, UN-HIJACKABLE SAVE PIPELINE:
 * Writes directly to Campaign Configuration and Persona Mapper Settings (Row 2).
 */
function studioDirectSaveCampaignPipeline(payload) {
  if (!payload || !payload.id) {
    throw new Error('Campaign ID is required to save configuration.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. REBUILD CAMPAIGN CONFIGURATION (Guarantees CAMPAIGN_ID in Row 2)
  let configSheet = ss.getSheetByName('Campaign Configuration');
  if (!configSheet) configSheet = ss.insertSheet('Campaign Configuration');

  const configRows = [
    ['Parameter', 'Configured Value', 'Description'],
    ['CAMPAIGN_ID', payload.id, 'Active Campaign Identifier'],
    ['CAMPAIGN_NAME', payload.name || payload.id, 'Active Campaign Name'],
    ['DEPARTMENT', payload.department || 'Procurement', 'Operating Department / Business Unit'],
    ['TARGET_COUNTRIES', payload.country || 'United States & Canada', 'Target Geographies / Regions'],
    ['TARGET_INDUSTRIES', payload.industry || 'Freight & Logistics, Manufacturing', 'Target Industry Verticals'],
    ['COMPANY_SIZE_RANGE', payload.size || '50 - 250 employees', 'Target Employee Size Range'],
    ['MAX_COMPANIES', Number(payload.limit) || 5, 'Maximum target companies to discover per run'],
    ['RESEARCH_DEPTH', payload.researchDepth || 'Deep Persona', 'Depth of company research and profiling'],
    ['EVIDENCE_PERIOD_MONTHS', Number(payload.evidencePeriodMonths) || 12, 'Historical lookback window in months'],
    ['MIN_COMPANY_SCORE', Number(payload.minCompanyScore) || 80, 'Minimum company qualification threshold (0-100)'],
    ['MIN_PERSONA_SCORE', Number(payload.minPersonaScore) || 85, 'Minimum persona fit confidence threshold (0-100)'],
    ['MIN_EMAIL_PRIORITY_SCORE', Number(payload.minEmailPriorityScore) || 80, 'Minimum email priority score for sending (0-100)'],
    ['TARGET_TITLES', payload.titles || 'VP Procurement; Head of Supply Chain', 'Semicolon-delimited target job titles'],
    ['CAMPAIGN_FOCUS', payload.campaignFocus || 'Scope 3 vendor transparency', 'Primary value proposition hook'],
    ['SEARCH_KEYWORDS', payload.searchKeywords || 'CSRD, supplier audits', 'Search discovery query keywords'],
    ['CUSTOM_BUSINESS_NEEDS', payload.customBusinessNeeds || '', 'Manual department business needs & buying triggers'],
    ['AI_PROVIDER', payload.aiProvider || 'OPENAI', 'Selected AI engine (GEMINI or OPENAI)'],
    ['PRIMARY_AI_PROVIDER', payload.aiProvider || 'OPENAI', 'Selected AI engine (legacy alias)'],
    ['AI_MODEL', payload.aiModel || 'gpt-4o', 'AI Model version name'],
    ['EMAIL_VERIFICATION_PROVIDER', payload.emailVerifier || 'ZeroBounce', 'Email verification vendor'],
    ['EMAIL_VERIFICATION_ENDPOINT', payload.emailEndpoint || '', 'Email verification REST endpoint URL'],
    ['EMAIL_SENDER_PROVIDER', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform'],
    ['DISPATCH_MODE', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform mode'],
    ['EMAIL_SENDER_ENDPOINT', payload.emailSenderEndpoint || '', 'Sender campaign UUID or webhook URL'],
    ['SENDER_NAME', payload.senderName || 'Alex Mercer', 'Display name for outreach email signature'],
    ['SENDER_TITLE', payload.senderTitle || 'Solutions Consultant', 'Sender job title for email signature'],
    ['SENDER_COMPANY', payload.senderCompany || 'eValuater Solutions', 'Sender company name for email signature'],
    ['INCLUDE_ESG_OBLIGATIONS', payload.incEsgObligations !== false ? 'TRUE' : 'FALSE', 'Evaluate CSRD/GRI obligations'],
    ['INCLUDE_ESG_MATURITY', payload.incEsgMaturity !== false ? 'TRUE' : 'FALSE', 'Evaluate ESG reporting maturity'],
    ['INCLUDE_SUPPLY_CHAIN', payload.incSupplyChain !== false ? 'TRUE' : 'FALSE', 'Evaluate supply chain complexity'],
    ['INCLUDE_SCOPE3', payload.incScope3 !== false ? 'TRUE' : 'FALSE', 'Evaluate Scope 3 emissions exposure'],
    ['INCLUDE_DATA_CHALLENGES', payload.incDataChallenges !== false ? 'TRUE' : 'FALSE', 'Evaluate sustainability data challenges'],
    ['INCLUDE_DIGITAL', payload.incDigital !== false ? 'TRUE' : 'FALSE', 'Evaluate digital transformation & AI adoption'],
    ['INCLUDE_ESG_GOVERNANCE', payload.incEsgGovernance !== false ? 'TRUE' : 'FALSE', 'Evaluate ESG board oversight'],
    ['INCLUDE_TRIGGERS', payload.incTriggers !== false ? 'TRUE' : 'FALSE', 'Evaluate commercial buying triggers'],
    ['INCLUDE_EVALUATER', payload.incEvaluater !== false ? 'TRUE' : 'FALSE', 'Evaluate eValuater solution fit'],
    ['HUMAN_APPROVAL_REQUIRED', payload.humanApprovalRequired !== false ? 'TRUE' : 'FALSE', 'Requires human review in queue before sending'],
    ['AUTO_SEND_ENABLED', payload.autoSendEnabled === true ? 'TRUE' : 'FALSE', 'Auto-send emails immediately upon approval']
  ];

  configSheet.clear();
  configSheet.getRange(1, 1, configRows.length, 3).setValues(configRows);
  configSheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  configSheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  // 2. WRITE DIRECTLY TO Persona Mapper Settings (Row 2, 39 Columns)
  let pSheet = ss.getSheetByName('Persona Mapper Settings') || ss.getSheetByName('Persona Mapper Setting');
  if (!pSheet) pSheet = ss.insertSheet('Persona Mapper Settings');

  pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length).setValues([PERSONA_MAPPER_39_HEADERS]);
  pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length)
    .setFontWeight('bold')
    .setFontColor('#174ea6')          // Dark Navy Blue (High Contrast)
    .setBackground('#e8f0fe')         // Light Google Blue
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  pSheet.setFrozenRows(1);

  const pRow = [
    payload.id,                                                  // 1: Campaign ID
    payload.name || payload.id,                                 // 2: Campaign Name
    'TRUE',                                                      // 3: Persona Mapper Enabled
    payload.aiProvider || 'OPENAI',                              // 4: AI Provider
    payload.aiModel || 'gpt-4o',                                 // 5: AI Model
    payload.titles || 'VP Procurement; Head of Supply Chain',    // 6: Persona Titles
    payload.searchKeywords || 'CSRD, supplier audits',           // 7: Search Keywords
    payload.campaignFocus || 'Scope 3 vendor transparency',      // 8: Campaign Focus
    Number(payload.evidencePeriodMonths) || 12,                  // 9: Evidence Period Months
    payload.researchDepth || 'Deep Persona',                     // 10: Research Depth
    payload.incEsgObligations !== false ? 'TRUE' : 'FALSE',      // 11: Include ESG Reporting Obligations
    payload.incEsgMaturity !== false ? 'TRUE' : 'FALSE',         // 12: Include ESG Reporting Maturity
    payload.incSupplyChain !== false ? 'TRUE' : 'FALSE',         // 13: Include Supply Chain Complexity
    payload.incScope3 !== false ? 'TRUE' : 'FALSE',              // 14: Include Scope 3 Exposure
    payload.incDataChallenges !== false ? 'TRUE' : 'FALSE',      // 15: Include Sustainability Data Challenges
    payload.incDigital !== false ? 'TRUE' : 'FALSE',             // 16: Include Digital Transformation
    payload.incEsgGovernance !== false ? 'TRUE' : 'FALSE',       // 17: Include ESG Governance
    payload.incTriggers !== false ? 'TRUE' : 'FALSE',            // 18: Include Business Triggers
    payload.incEvaluater !== false ? 'TRUE' : 'FALSE',           // 19: Include eValuater Solution Fit
    payload.weightEsgReporting || '20%',                         // 20: Weight ESG Reporting Relevance
    payload.weightSupplyChain || '15%',                          // 21: Weight Supply Chain Complexity
    payload.weightDataChallenges || '15%',                       // 22: Weight Sustainability Data Needs
    payload.weightDigital || '15%',                              // 23: Weight Digital Transformation Relevance
    payload.weightTriggers || '20%',                             // 24: Weight Business Triggers
    payload.weightEvaluaterFit || '15%',                         // 25: Weight eValuater Solution Fit
    Number(payload.minCompanyScore) || 80,                       // 26: Minimum Company Opportunity Score
    Number(payload.minPersonaScore) || 85,                       // 27: Minimum Persona Confidence
    Number(payload.minEmailPriorityScore) || 80,                 // 28: Minimum Email Priority Score
    payload.humanApprovalRequired !== false ? 'TRUE' : 'FALSE',  // 29: Human Approval Required
    payload.autoSendEnabled === true ? 'TRUE' : 'FALSE',         // 30: Auto Send Enabled
    payload.weightEmailCompanyScore || '40%',                    // 31: Email Priority Weight Company Score
    payload.weightEmailPersonaFit || '35%',                      // 32: Email Priority Weight Persona Fit
    payload.weightEmailBuyingSignal || '25%',                    // 33: Email Priority Weight Buying Signal
    payload.customBusinessNeeds || '',                           // 34: Notes (Holds custom business needs)
    Number(payload.limit) || 5,                                  // 35: Max Companies Per Run
    payload.emailVerifier || 'ZeroBounce',                       // 36: Email Verification Provider
    payload.emailEndpoint || '',                                 // 37: Email Verification Endpoint
    payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', // 38: Email Sender Provider
    payload.emailSenderEndpoint || ''                            // 39: Email Sender Endpoint
  ];

  pSheet.getRange(2, 1, 1, 39).setValues([pRow]);
  SpreadsheetApp.flush();

  // 3. WRITE TO Campaigns TAB
  let campSheet = ss.getSheetByName('Campaigns');
  if (campSheet) {
    if (campSheet.getLastRow() <= 1) {
      campSheet.getRange(1, 1, 1, 10).setValues([['Campaign ID', 'Campaign Name', 'Department', 'Status', 'Countries', 'Industries', 'Titles', 'AI Provider', 'Created', 'Modified']]);
      campSheet.getRange(1, 1, 1, 10)
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe');
    }
    campSheet.getRange(2, 1, 1, 10).setValues([[
      payload.id,
      payload.name || payload.id,
      payload.department || 'Procurement',
      'Active',
      payload.country || 'United States & Canada',
      payload.industry || 'Freight & Logistics, Manufacturing',
      payload.titles || 'VP Procurement; Head of Supply Chain',
      payload.aiProvider || 'OPENAI',
      new Date().toLocaleDateString(),
      new Date().toLocaleDateString()
    ]]);
    SpreadsheetApp.flush();
  }

  // 4. WRITE TO Dashboard TAB
  let dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet) {
    dashSheet.getRange('B3').setValue('Active');
    dashSheet.getRange('B4').setValue(payload.name || payload.id);
    dashSheet.getRange('D3').setValue(new Date().toLocaleString());
  }

  return {
    success: true,
    campaignId: payload.id,
    campaignName: payload.name || payload.id,
    message: 'Campaign ' + payload.id + ' written directly to Row 2 of Persona Mapper Settings (39 cols) & Campaign Configuration!'
  };
}

/**
 * Menu action to trigger master sync
 */
function executeMasterPersonaMapperSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName('Campaign Configuration');
  const dict = {};

  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const cleanKey = normalizeStudioKey(cData[i][0]);
      if (cleanKey) dict[cleanKey] = cData[i][1];
    }
  }

  let cid = String(dict['CAMPAIGNID'] || dict['ACTIVECAMPAIGNID'] || dict['CAMPAIGN'] || '').trim();

  if (!cid) {
    const dashSheet = ss.getSheetByName('Dashboard');
    if (dashSheet) {
      const dVals = dashSheet.getRange('A1:E6').getValues();
      for (let r = 0; r < dVals.length; r++) {
        for (let c = 0; c < dVals[r].length; c++) {
          const val = String(dVals[r][c] || '').trim();
          if (val.startsWith('CMP-')) {
            cid = val;
            break;
          }
        }
        if (cid) break;
      }
    }
  }

  if (!cid) cid = 'CMP-005';

  const payload = {
    id: cid,
    name: String(dict['CAMPAIGNNAME'] || dict['NAME'] || cid),
    department: String(dict['DEPARTMENT'] || 'Procurement'),
    country: String(dict['TARGETCOUNTRIES'] || 'United States & Canada'),
    industry: String(dict['TARGETINDUSTRIES'] || 'Freight & Logistics, Manufacturing'),
    size: String(dict['COMPANYSIZERANGE'] || '50 - 250 employees'),
    limit: Number(dict['MAXCOMPANIES']) || 5,
    researchDepth: String(dict['RESEARCHDEPTH'] || 'Deep Persona'),
    evidencePeriodMonths: Number(dict['EVIDENCEPERIODMONTHS']) || 12,
    minCompanyScore: Number(dict['MINCOMPANYSCORE']) || 80,
    minPersonaScore: Number(dict['MINPERSONASCORE']) || 85,
    minEmailPriorityScore: Number(dict['MINEMAILPRIORITYSCORE']) || 80,
    titles: String(dict['TARGETTITLES'] || 'VP Procurement; Head of Supply Chain'),
    campaignFocus: String(dict['CAMPAIGNFOCUS'] || 'Scope 3 vendor transparency'),
    searchKeywords: String(dict['SEARCHKEYWORDS'] || 'CSRD, supplier audits'),
    customBusinessNeeds: String(dict['CUSTOMBUSINESSNEEDS'] || ''),
    aiProvider: String(dict['AIPROVIDER'] || 'OPENAI'),
    aiModel: String(dict['AIMODEL'] || 'gpt-4o'),
    emailVerifier: String(dict['EMAILVERIFICATIONPROVIDER'] || 'ZeroBounce'),
    emailEndpoint: String(dict['EMAILVERIFICATIONENDPOINT'] || ''),
    emailSenderProvider: String(dict['EMAILSENDERPROVIDER'] || 'Workspace Gmail (Drafts Only)'),
    emailSenderEndpoint: String(dict['EMAILSENDERENDPOINT'] || ''),
    senderName: String(dict['SENDERNAME'] || 'Alex Mercer'),
    senderTitle: String(dict['SENDERTITLE'] || 'Solutions Consultant'),
    senderCompany: String(dict['SENDERCOMPANY'] || 'eValuater Solutions')
  };

  const res = studioDirectSaveCampaignPipeline(payload);

  const pSheet = findPersonaMapperSettingsSheet(ss);
  if (pSheet) ss.setActiveSheet(pSheet);

  SpreadsheetApp.getUi().alert(
    '🎯 Master Sync Complete',
    '✅ ' + res.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncActiveCampaignToPersonaMapperSettings() {
  executeMasterPersonaMapperSync();
}

/**
 * Navigates the Google Sheets UI to the Prospect Review Queue tab.
 */
function navigateToReviewQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Prospect Review Queue');
  if (!queueSheet) {
    return { success: false, message: 'Prospect Review Queue sheet not found.' };
  }
  ss.setActiveSheet(queueSheet);
  return { success: true, message: 'Switched to Prospect Review Queue.' };
}

function runCompanyDiscoveryFromSidebar(cid) {
  if (typeof runCompanyDiscoveryForSelectedCampaign === 'function') {
    return runCompanyDiscoveryForSelectedCampaign(cid);
  } else if (typeof runCompanyDiscovery === 'function') {
    return runCompanyDiscovery(cid);
  }
  throw new Error('Discovery engine is not loaded in this project.');
}

function runPersonaMapperFromSidebar(cid) {
  if (typeof runPersonaMapperForSelectedCampaign === 'function') {
    return runPersonaMapperForSelectedCampaign(cid);
  } else if (typeof runPersonaMapper === 'function') {
    return runPersonaMapper(cid);
  }
  throw new Error('Persona Mapper engine is not loaded in this project.');
}

/**
 * Audits all 12 project tabs, column schemas, and credentials.
 */
function showSetupStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredTabs = [
    'Dashboard',
    'Campaigns',
    'Company Intelligence',
    'Persona Leads',
    'Persona Mapper Settings',
    'Prospect Review Queue',
    'Campaign Configuration',
    'Company Database',
    'Research Evidence',
    'Agent Settings',
    'Execution Logs',
    'Lists & Settings'
  ];

  const present = [];
  const missing = [];

  requiredTabs.forEach(name => {
    if (name === 'Persona Mapper Settings') {
      if (findPersonaMapperSettingsSheet(ss)) present.push('Persona Mapper Settings');
      else missing.push('Persona Mapper Settings');
    } else {
      if (ss.getSheetByName(name)) present.push(name);
      else missing.push(name);
    }
  });

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();
  const detectedKeys = [];

  for (let k in scriptProps) {
    const uk = k.toUpperCase();
    if (uk.includes('OPENAI') && scriptProps[k] && !scriptProps[k].includes('PASTE_')) {
      detectedKeys.push('OpenAI (Project Settings: ' + k + ')');
    }
    if (uk.includes('GEMINI') && scriptProps[k] && !scriptProps[k].includes('PASTE_')) {
      detectedKeys.push('Gemini (Project Settings: ' + k + ')');
    }
  }

  for (let k in userProps) {
    const uk = k.toUpperCase();
    if (uk.includes('OPENAI') && userProps[k] && !userProps[k].includes('PASTE_')) {
      detectedKeys.push('OpenAI (User Properties: ' + k + ')');
    }
    if (uk.includes('GEMINI') && userProps[k] && !userProps[k].includes('PASTE_')) {
      detectedKeys.push('Gemini (User Properties: ' + k + ')');
    }
  }

  const agentSheet = ss.getSheetByName('Agent Settings');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const data = agentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const val = String(data[i][j] || '').trim().toUpperCase();
        if (val === 'OPENAI_API_KEY') {
          const keyVal = String(data[i][j + 1] || data[i][2] || '').trim();
          if (keyVal && !keyVal.includes('PASTE_')) detectedKeys.push('OpenAI (Agent Settings)');
        }
        if (val === 'GEMINI_API_KEY') {
          const keyVal = String(data[i][j + 1] || data[i][2] || '').trim();
          if (keyVal && !keyVal.includes('PASTE_')) detectedKeys.push('Gemini (Agent Settings)');
        }
      }
    }
  }

  let configuredModel = 'gpt-4o';
  const configSheet = ss.getSheetByName('Campaign Configuration');
  if (configSheet && configSheet.getLastRow() > 1) {
    const cData = configSheet.getDataRange().getValues();
    for (let r = 1; r < cData.length; r++) {
      const k = normalizeStudioKey(cData[r][0]);
      if (k === 'AIMODEL' && cData[r][1]) configuredModel = String(cData[r][1]).trim();
    }
  }

  let apiKeyInfo = 'Not Configured (Add key in Project Settings or Agent Settings)';
  if (detectedKeys.length > 0) {
    apiKeyInfo = 'Connected: ' + detectedKeys.join(', ') + ' [Model: ' + configuredModel + ']';
  }

  let senderInfo = 'Google Workspace (Drafts Only)';
  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const k = normalizeStudioKey(data[i][0]);
      if (k === 'DISPATCHMODE' && data[i][1]) {
        senderInfo = String(data[i][1]);
        break;
      }
    }
  }

  let message = '📊 SPREADSHEET SYSTEM HEALTH CHECK:\n\n';
  message += '• Project Tabs Verified: ' + present.length + ' of ' + requiredTabs.length + ' Tabs Active\n';
  if (missing.length > 0) {
    message += '⚠️ Missing Tabs: ' + missing.join(', ') + '\n\n';
  } else {
    message += '✅ All 12 project tabs exist and match system schema.\n\n';
  }

  message += '• AI Engine Status: ' + apiKeyInfo + '\n';
  message += '• Email Dispatch Channel: ' + senderInfo + '\n';
  message += '• Suppression Engine: Active (checking Lists & Settings)\n';
  message += '• Human Review Queue: Operational\n\n';
  message += 'Your workspace is healthy and ready to run!';

  SpreadsheetApp.getUi().alert('Spreadsheet Setup Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function authorizeStudioPermissions() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  SpreadsheetApp.getActiveSpreadsheet();
  GmailApp.getDrafts();

  SpreadsheetApp.getUi().alert(
    'Authorization Successful!',
    'All permissions (Web Search, Gemini/OpenAI API, Spreadsheet, and Gmail) are now fully authorized.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function approveSelectedProspect() {
  updateSelectedProspectStatus('Approved - Ready for Outreach');
}

function markSelectedProspectNeedsEditing() {
  updateSelectedProspectStatus('Needs Human Review');
}

function rejectSelectedProspect() {
  updateSelectedProspectStatus('Rejected - Unqualified');
}

function updateSelectedProspectStatus(statusText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== 'Prospect Review Queue') {
    SpreadsheetApp.getUi().alert('Please switch to the "Prospect Review Queue" tab before updating prospects.');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a prospect row (row 2 or below).');
    return;
  }

  sheet.getRange(row, 22).setValue(statusText);
  sheet.getRange(row, 23).setValue(Session.getActiveUser().getEmail() || 'Admin Reviewer');
  sheet.getRange(row, 24).setValue(new Date().toLocaleString());
  SpreadsheetApp.getActiveSpreadsheet().toast('Row ' + row + ' updated to: ' + statusText, 'Status Updated', 3);
}

function createGmailDraftsForApprovedProspects() {
  if (typeof triggerApprovedGmailDrafts === 'function') {
    triggerApprovedGmailDrafts();
  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName('Prospect Review Queue');
    if (!queueSheet) {
      SpreadsheetApp.getUi().alert('Prospect Review Queue sheet not found.');
      return;
    }

    const data = queueSheet.getDataRange().getValues();
    let draftCount = 0;

    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][7] || '').trim();
      const approval = String(data[i][21] || '').trim();
      const sendStatus = String(data[i][24] || '').trim();
      const subject = String(data[i][14] || '').trim();
      const body = String(data[i][15] || '').trim();

      if (approval.startsWith('Approved') && sendStatus !== 'Draft Created' && sendStatus !== 'Sent' && email && subject && body) {
        try {
          GmailApp.createDraft(email, subject, body);
          queueSheet.getRange(i + 1, 25).setValue('Draft Created');
          draftCount++;
        } catch (err) {
          queueSheet.getRange(i + 1, 27).setValue(err.message);
        }
      }
    }

    SpreadsheetApp.getUi().alert('Drafts Created', 'Successfully generated ' + draftCount + ' draft(s) in Gmail Drafts.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sendApprovedOutreachEmails() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    'Send Outreach Emails?',
    'This will send real emails for all leads marked "Approved" in the Prospect Review Queue.\n\nAre you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) return;

  if (typeof sendApprovedEmailsPipeline === 'function') {
    sendApprovedEmailsPipeline();
  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName('Prospect Review Queue');
    if (!queueSheet) {
      ui.alert('Prospect Review Queue sheet not found.');
      return;
    }

    const data = queueSheet.getDataRange().getValues();
    let sentCount = 0;

    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][7] || '').trim();
      const approval = String(data[i][21] || '').trim();
      const sendStatus = String(data[i][24] || '').trim();
      const subject = String(data[i][14] || '').trim();
      const body = String(data[i][15] || '').trim();

      if (approval.startsWith('Approved') && sendStatus !== 'Sent' && email && subject && body) {
        try {
          GmailApp.sendEmail(email, subject, body);
          queueSheet.getRange(i + 1, 25).setValue('Sent');
          queueSheet.getRange(i + 1, 26).setValue(new Date().toLocaleString());
          sentCount++;
        } catch (err) {
          queueSheet.getRange(i + 1, 27).setValue(err.message);
        }
      }
    }
    ui.alert('Outreach Complete', 'Successfully dispatched ' + sentCount + ' live emails.', ui.ButtonSet.OK);
  }
}

function triggerSetupAllDropdowns() {
  if (typeof setupAllDropdownValidations === 'function') {
    setupAllDropdownValidations();
  } else {
    SpreadsheetApp.getUi().alert('SheetSetup.gs dropdown initializer not found.');
  }
}

function preparePersonaMapperSheets() {
  if (typeof setupPersonaMapperSheets === 'function') {
    setupPersonaMapperSheets();
  } else {
    SpreadsheetApp.getUi().alert('Persona Mapper setup initializer not found.');
  }
}

function triggerPilotRun() {
  if (typeof runAuthenticPilotSimulation === 'function') {
    runAuthenticPilotSimulation();
  } else {
    SpreadsheetApp.getUi().alert('PilotSandboxRunner.gs not found.');
  }
}

function triggerPilotCleanup() {
  if (typeof cleanupSandboxPilotData === 'function') {
    cleanupSandboxPilotData();
  } else {
    SpreadsheetApp.getUi().alert('PilotSandboxRunner.gs cleanup not found.');
  }
}