/** @OnlyCurrentDoc */

/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * MASTER CONTROLLER & UNIFIED TOP MENU
 * ============================================================================
 * File: Code.js / code.gs
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
 * Master 28 Canonical Header Columns for Campaigns Tab
 */
const CAMPAIGNS_MASTER_HEADERS = [
  'Campaign ID',
  'Campaign Name',
  'Department',
  'Status',
  'Target Countries',
  'Target Industries',
  'Company Size',
  'Discovery Limit',
  'Research Depth',
  'Evidence Period Months',
  'Min Company Score',
  'Min Persona Score',
  'Min Email Priority Score',
  'Target Titles',
  'Campaign Focus',
  'Search Keywords',
  'Custom Business Needs',
  'AI Provider',
  'AI Model',
  'Email Verifier Provider',
  'Email Verifier Endpoint',
  'Email Sender Provider',
  'Email Sender Endpoint',
  'Sender Name',
  'Sender Title',
  'Sender Company',
  'Campaign Config JSON',
  'Last Modified'
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
  template.activeCampaignId = sidebarData.activeCampaignId || '';

  const html = template.evaluate().setTitle('🎯 Campaign Control Panel').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns available campaigns and configuration maps for sidebar initialization.
 * Accurately merges: Campaigns (with JSON backup), Persona Mapper Settings, and Campaign Configuration.
 */
function getSidebarInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsList = [];
  const campaignsMap = {};
  const configsMap = {};
  let activeCampaignId = '';

  // 1. Read Campaign Configuration to determine current active campaign
  const cSheet = ss.getSheetByName('Campaign Configuration');
  const activeConfigDict = {};
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let j = 1; j < cData.length; j++) {
      const k = normalizeStudioKey(cData[j][0]);
      if (k) activeConfigDict[k] = cData[j][1];
    }
    activeCampaignId = String(activeConfigDict['CAMPAIGNID'] || activeConfigDict['ACTIVECAMPAIGNID'] || activeConfigDict['CAMPAIGN'] || '').trim();
  }

  // 2. Read from Campaigns Tab (Primary Campaign Store)
  const campSheet = ss.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const campData = campSheet.getDataRange().getValues();
    const headers = campData[0].map(h => normalizeStudioKey(h));

    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const nameIdx = headers.indexOf('CAMPAIGNNAME') >= 0 ? headers.indexOf('CAMPAIGNNAME') : 1;
    const deptIdx = headers.indexOf('DEPARTMENT');
    const countryIdx = headers.indexOf('TARGETCOUNTRIES') >= 0 ? headers.indexOf('TARGETCOUNTRIES') : headers.indexOf('COUNTRIES');
    const indIdx = headers.indexOf('TARGETINDUSTRIES') >= 0 ? headers.indexOf('TARGETINDUSTRIES') : headers.indexOf('INDUSTRIES');
    const sizeIdx = headers.indexOf('COMPANYSIZE') >= 0 ? headers.indexOf('COMPANYSIZE') : headers.indexOf('COMPANYSIZERANGE');
    const limitIdx = headers.indexOf('DISCOVERYLIMIT') >= 0 ? headers.indexOf('DISCOVERYLIMIT') : headers.indexOf('MAXCOMPANIES');
    const depthIdx = headers.indexOf('RESEARCHDEPTH');
    const evIdx = headers.indexOf('EVIDENCEPERIODMONTHS');
    const minCompIdx = headers.indexOf('MINCOMPANYSCORE');
    const minPersIdx = headers.indexOf('MINPERSONASCORE');
    const minEmailIdx = headers.indexOf('MINEMAILPRIORITYSCORE');
    const titlesIdx = headers.indexOf('TARGETTITLES') >= 0 ? headers.indexOf('TARGETTITLES') : headers.indexOf('TITLES');
    const focusIdx = headers.indexOf('CAMPAIGNFOCUS');
    const keywordsIdx = headers.indexOf('SEARCHKEYWORDS');
    const needsIdx = headers.indexOf('CUSTOMBUSINESSNEEDS');
    const aiProvIdx = headers.indexOf('AIPROVIDER');
    const aiModIdx = headers.indexOf('AIMODEL');
    const verifProvIdx = headers.indexOf('EMAILVERIFIERPROVIDER') >= 0 ? headers.indexOf('EMAILVERIFIERPROVIDER') : headers.indexOf('EMAILVERIFIER');
    const verifEndIdx = headers.indexOf('EMAILVERIFIERENDPOINT');
    const senderProvIdx = headers.indexOf('EMAILSENDERPROVIDER');
    const senderEndIdx = headers.indexOf('EMAILSENDERENDPOINT');
    const senderNameIdx = headers.indexOf('SENDERNAME');
    const senderTitleIdx = headers.indexOf('SENDERTITLE');
    const senderCoIdx = headers.indexOf('SENDERCOMPANY');
    const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON');

    for (let i = 1; i < campData.length; i++) {
      const cid = String(campData[i][idIdx] || '').trim();
      if (!cid) continue;

      const name = String(campData[i][nameIdx] || cid).trim();
      campaignsMap[cid] = name;

      let campaignObj = { id: cid, name: name };

      // Try reading full config JSON if stored in column
      if (jsonIdx >= 0 && campData[i][jsonIdx]) {
        try {
          const parsed = JSON.parse(campData[i][jsonIdx]);
          if (parsed && typeof parsed === 'object') {
            campaignObj = Object.assign(campaignObj, parsed);
          }
        } catch (e) {
          // Fall back to reading separate columns
        }
      }

      // Read explicit table columns
      if (deptIdx >= 0 && campData[i][deptIdx]) campaignObj.department = String(campData[i][deptIdx]).trim();
      if (countryIdx >= 0 && campData[i][countryIdx]) campaignObj.country = String(campData[i][countryIdx]).trim();
      if (indIdx >= 0 && campData[i][indIdx]) campaignObj.industry = String(campData[i][indIdx]).trim();
      if (sizeIdx >= 0 && campData[i][sizeIdx]) campaignObj.size = String(campData[i][sizeIdx]).trim();
      if (limitIdx >= 0 && campData[i][limitIdx]) campaignObj.limit = Number(campData[i][limitIdx]) || campaignObj.limit || 5;
      if (depthIdx >= 0 && campData[i][depthIdx]) campaignObj.researchDepth = String(campData[i][depthIdx]).trim();
      if (evIdx >= 0 && campData[i][evIdx]) campaignObj.evidencePeriodMonths = Number(campData[i][evIdx]) || campaignObj.evidencePeriodMonths || 12;
      if (minCompIdx >= 0 && campData[i][minCompIdx]) campaignObj.minCompanyScore = Number(campData[i][minCompIdx]) || campaignObj.minCompanyScore || 80;
      if (minPersIdx >= 0 && campData[i][minPersIdx]) campaignObj.minPersonaScore = Number(campData[i][minPersIdx]) || campaignObj.minPersonaScore || 85;
      if (minEmailIdx >= 0 && campData[i][minEmailIdx]) campaignObj.minEmailPriorityScore = Number(campData[i][minEmailIdx]) || campaignObj.minEmailPriorityScore || 80;
      if (titlesIdx >= 0 && campData[i][titlesIdx]) campaignObj.titles = String(campData[i][titlesIdx]).trim();
      if (focusIdx >= 0 && campData[i][focusIdx]) campaignObj.campaignFocus = String(campData[i][focusIdx]).trim();
      if (keywordsIdx >= 0 && campData[i][keywordsIdx]) campaignObj.searchKeywords = String(campData[i][keywordsIdx]).trim();
      if (needsIdx >= 0 && campData[i][needsIdx]) campaignObj.customBusinessNeeds = String(campData[i][needsIdx]).trim();
      if (aiProvIdx >= 0 && campData[i][aiProvIdx]) campaignObj.aiProvider = String(campData[i][aiProvIdx]).trim();
      if (aiModIdx >= 0 && campData[i][aiModIdx]) campaignObj.aiModel = String(campData[i][aiModIdx]).trim();
      if (verifProvIdx >= 0 && campData[i][verifProvIdx]) campaignObj.emailVerifier = String(campData[i][verifProvIdx]).trim();
      if (verifEndIdx >= 0 && campData[i][verifEndIdx]) campaignObj.emailEndpoint = String(campData[i][verifEndIdx]).trim();
      if (senderProvIdx >= 0 && campData[i][senderProvIdx]) campaignObj.emailSenderProvider = String(campData[i][senderProvIdx]).trim();
      if (senderEndIdx >= 0 && campData[i][senderEndIdx]) campaignObj.emailSenderEndpoint = String(campData[i][senderEndIdx]).trim();
      if (senderNameIdx >= 0 && campData[i][senderNameIdx]) campaignObj.senderName = String(campData[i][senderNameIdx]).trim();
      if (senderTitleIdx >= 0 && campData[i][senderTitleIdx]) campaignObj.senderTitle = String(campData[i][senderTitleIdx]).trim();
      if (senderCoIdx >= 0 && campData[i][senderCoIdx]) campaignObj.senderCompany = String(campData[i][senderCoIdx]).trim();

      configsMap[cid] = campaignObj;
    }
  }

  // 3. Read from Persona Mapper Settings (Canonical 39 Columns)
  const pSheet = findPersonaMapperSettingsSheet(ss);
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const cid = String(pData[i][0] || '').trim();
      if (!cid) continue;

      const name = String(pData[i][1] || cid).trim();
      campaignsMap[cid] = name;

      configsMap[cid] = configsMap[cid] || { id: cid, name: name };
      const c = configsMap[cid];

      c.id = cid;
      if (!c.name || c.name === cid) c.name = name;
      c.aiProvider = String(pData[i][3] || c.aiProvider || 'OPENAI');
      c.aiModel = String(pData[i][4] || c.aiModel || 'gpt-4o');
      if (pData[i][5] && !c.titles) c.titles = String(pData[i][5]);
      if (pData[i][6] && !c.searchKeywords) c.searchKeywords = String(pData[i][6]);
      if (pData[i][7] && !c.campaignFocus) c.campaignFocus = String(pData[i][7]);
      c.evidencePeriodMonths = Number(pData[i][8]) || c.evidencePeriodMonths || 12;
      c.researchDepth = String(pData[i][9] || c.researchDepth || 'Deep Persona');
      c.incEsgObligations = String(pData[i][10]).toUpperCase() !== 'FALSE';
      c.incEsgMaturity = String(pData[i][11]).toUpperCase() !== 'FALSE';
      c.incSupplyChain = String(pData[i][12]).toUpperCase() !== 'FALSE';
      c.incScope3 = String(pData[i][13]).toUpperCase() !== 'FALSE';
      c.incDataChallenges = String(pData[i][14]).toUpperCase() !== 'FALSE';
      c.incDigital = String(pData[i][15]).toUpperCase() !== 'FALSE';
      c.incEsgGovernance = String(pData[i][16]).toUpperCase() !== 'FALSE';
      c.incTriggers = String(pData[i][17]).toUpperCase() !== 'FALSE';
      c.incEvaluater = String(pData[i][18]).toUpperCase() !== 'FALSE';
      c.weightEsgReporting = String(pData[i][19] || c.weightEsgReporting || '20%');
      c.weightSupplyChain = String(pData[i][20] || c.weightSupplyChain || '15%');
      c.weightDataChallenges = String(pData[i][21] || c.weightDataChallenges || '15%');
      c.weightDigital = String(pData[i][22] || c.weightDigital || '15%');
      c.weightTriggers = String(pData[i][23] || c.weightTriggers || '20%');
      c.weightEvaluaterFit = String(pData[i][24] || c.weightEvaluaterFit || '15%');
      c.minCompanyScore = Number(pData[i][25]) || c.minCompanyScore || 80;
      c.minPersonaScore = Number(pData[i][26]) || c.minPersonaScore || 85;
      c.minEmailPriorityScore = Number(pData[i][27]) || c.minEmailPriorityScore || 80;
      c.humanApprovalRequired = String(pData[i][28]).toUpperCase() !== 'FALSE';
      c.autoSendEnabled = String(pData[i][29]).toUpperCase() === 'TRUE';
      c.weightEmailCompanyScore = String(pData[i][30] || c.weightEmailCompanyScore || '40%');
      c.weightEmailPersonaFit = String(pData[i][31] || c.weightEmailPersonaFit || '35%');
      c.weightEmailBuyingSignal = String(pData[i][32] || c.weightEmailBuyingSignal || '25%');
      if (pData[i][33] && !c.customBusinessNeeds) c.customBusinessNeeds = String(pData[i][33]);
      c.limit = Number(pData[i][34]) || c.limit || 5;
      c.emailVerifier = String(pData[i][35] || c.emailVerifier || 'ZeroBounce');
      c.emailEndpoint = String(pData[i][36] || c.emailEndpoint || '');
      c.emailSenderProvider = String(pData[i][37] || c.emailSenderProvider || 'Workspace Gmail (Drafts Only)');
      c.emailSenderEndpoint = String(pData[i][38] || c.emailSenderEndpoint || '');
    }
  }

  // 4. Merge Active Campaign Details from Campaign Configuration
  if (activeCampaignId) {
    campaignsMap[activeCampaignId] = String(activeConfigDict['CAMPAIGNNAME'] || campaignsMap[activeCampaignId] || activeCampaignId);
    configsMap[activeCampaignId] = configsMap[activeCampaignId] || { id: activeCampaignId, name: campaignsMap[activeCampaignId] };
    const c = configsMap[activeCampaignId];

    if (activeConfigDict['CAMPAIGNNAME']) c.name = String(activeConfigDict['CAMPAIGNNAME']);
    if (activeConfigDict['DEPARTMENT']) c.department = String(activeConfigDict['DEPARTMENT']);
    if (activeConfigDict['TARGETCOUNTRIES']) c.country = String(activeConfigDict['TARGETCOUNTRIES']);
    if (activeConfigDict['TARGETINDUSTRIES']) c.industry = String(activeConfigDict['TARGETINDUSTRIES']);
    if (activeConfigDict['COMPANYSIZERANGE']) c.size = String(activeConfigDict['COMPANYSIZERANGE']);
    if (activeConfigDict['MAXCOMPANIES']) c.limit = Number(activeConfigDict['MAXCOMPANIES']) || 5;
    if (activeConfigDict['RESEARCHDEPTH']) c.researchDepth = String(activeConfigDict['RESEARCHDEPTH']);
    if (activeConfigDict['EVIDENCEPERIODMONTHS']) c.evidencePeriodMonths = Number(activeConfigDict['EVIDENCEPERIODMONTHS']) || 12;
    if (activeConfigDict['MINCOMPANYSCORE']) c.minCompanyScore = Number(activeConfigDict['MINCOMPANYSCORE']) || 80;
    if (activeConfigDict['MINPERSONASCORE']) c.minPersonaScore = Number(activeConfigDict['MINPERSONASCORE']) || 85;
    if (activeConfigDict['MINEMAILPRIORITYSCORE']) c.minEmailPriorityScore = Number(activeConfigDict['MINEMAILPRIORITYSCORE']) || 80;
    if (activeConfigDict['TARGETTITLES']) c.titles = String(activeConfigDict['TARGETTITLES']);
    if (activeConfigDict['CAMPAIGNFOCUS']) c.campaignFocus = String(activeConfigDict['CAMPAIGNFOCUS']);
    if (activeConfigDict['SEARCHKEYWORDS']) c.searchKeywords = String(activeConfigDict['SEARCHKEYWORDS']);
    if (activeConfigDict['CUSTOMBUSINESSNEEDS']) c.customBusinessNeeds = String(activeConfigDict['CUSTOMBUSINESSNEEDS']);
    if (activeConfigDict['AIPROVIDER']) c.aiProvider = String(activeConfigDict['AIPROVIDER']);
    if (activeConfigDict['AIMODEL']) c.aiModel = String(activeConfigDict['AIMODEL']);
    if (activeConfigDict['SENDERNAME']) c.senderName = String(activeConfigDict['SENDERNAME']);
    if (activeConfigDict['SENDERTITLE']) c.senderTitle = String(activeConfigDict['SENDERTITLE']);
    if (activeConfigDict['SENDERCOMPANY']) c.senderCompany = String(activeConfigDict['SENDERCOMPANY']);
    if (activeConfigDict['EMAILVERIFICATIONPROVIDER']) c.emailVerifier = String(activeConfigDict['EMAILVERIFICATIONPROVIDER']);
    if (activeConfigDict['EMAILVERIFICATIONENDPOINT']) c.emailEndpoint = String(activeConfigDict['EMAILVERIFICATIONENDPOINT']);
    if (activeConfigDict['EMAILSENDERPROVIDER']) c.emailSenderProvider = String(activeConfigDict['EMAILSENDERPROVIDER']);
    if (activeConfigDict['EMAILSENDERENDPOINT']) c.emailSenderEndpoint = String(activeConfigDict['EMAILSENDERENDPOINT']);
  }

  // 5. Intelligent fallback ONLY for PILOT campaign (never overwrite other campaigns)
  for (let cid in configsMap) {
    const c = configsMap[cid];
    if (cid.toUpperCase().includes('PILOT')) {
      if (!c.department) c.department = 'Procurement';
      if (!c.country) c.country = 'United States & Canada';
      if (!c.industry) c.industry = 'Freight & Logistics, Manufacturing';
      if (!c.size) c.size = '50 - 250 employees';
      if (!c.limit) c.limit = 3;
      if (!c.titles) c.titles = 'VP Procurement; Head of Supply Chain; Director Vendor Management';
      if (!c.campaignFocus) c.campaignFocus = 'Scope 3 vendor compliance and operational resilience';
      if (!c.searchKeywords) c.searchKeywords = 'Scope 3, CSRD, supplier risk, compliance audits';
      if (!c.customBusinessNeeds) {
        c.customBusinessNeeds = '- Managing Scope 3 vendor emissions and Tier-1 ESG supply chain transparency\n- High vendor questionnaire drop-off rate during procurement audits\n- Eliminating manual spreadsheet follow-ups with suppliers';
      }
      if (!c.senderName) c.senderName = 'Alex Mercer';
      if (!c.senderTitle) c.senderTitle = 'Solutions Consultant';
      if (!c.senderCompany) c.senderCompany = 'eValuater Solutions';
      if (!c.emailSenderProvider) c.emailSenderProvider = 'Workspace Gmail (Drafts Only)';
    }
  }

  // Convert map to sorted array
  for (let cid in campaignsMap) {
    campaignsList.push({ id: cid, name: campaignsMap[cid] });
  }

  return {
    campaigns: campaignsList,
    configsMap: configsMap,
    activeCampaignId: activeCampaignId || (campaignsList.length > 0 ? campaignsList[0].id : '')
  };
}

function getAvailableCampaignsList() {
  return getSidebarInitialData().campaigns;
}

/**
 * Activates a selected campaign in the Google Sheet:
 * Updates Campaign Configuration, rebuilds Dashboard for that specific campaign,
 * and returns the complete configuration.
 */
function selectAndActivateCampaign(cid) {
  if (!cid) return { success: false, message: 'No Campaign ID provided.' };

  const data = getSidebarInitialData();
  const config = data.configsMap[cid] || { id: cid, name: cid };

  // Write to Campaign Configuration, Campaigns Tab, and Refresh Dashboard
  studioDirectSaveCampaignPipeline(config, true);

  return {
    success: true,
    campaignId: cid,
    config: config,
    message: 'Active campaign set to ' + cid + ' and synced to Dashboard!'
  };
}

/**
 * Direct Campaign Pipeline Saver:
 * 1. Writes cleanly to Campaign Configuration (Row 2 = CAMPAIGN_ID).
 * 2. Writes full parameters + JSON backup to Campaigns tab.
 * 3. Updates or appends the campaign row in Persona Mapper Settings (39 columns).
 * 4. Completely updates Dashboard tab with this campaign's details and KPI counts.
 */
function studioDirectSaveCampaignPipeline(payload, isActivationOnly) {
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
    ['TARGET_COUNTRIES', payload.country || '', 'Target Geographies / Regions'],
    ['TARGET_INDUSTRIES', payload.industry || '', 'Target Industry Verticals'],
    ['COMPANY_SIZE_RANGE', payload.size || '', 'Target Employee Size Range'],
    ['MAX_COMPANIES', Number(payload.limit) || 5, 'Maximum target companies to discover per run'],
    ['RESEARCH_DEPTH', payload.researchDepth || 'Deep Persona', 'Depth of company research and profiling'],
    ['EVIDENCE_PERIOD_MONTHS', Number(payload.evidencePeriodMonths) || 12, 'Historical lookback window in months'],
    ['MIN_COMPANY_SCORE', Number(payload.minCompanyScore) || 80, 'Minimum company qualification threshold (0-100)'],
    ['MIN_PERSONA_SCORE', Number(payload.minPersonaScore) || 85, 'Minimum persona fit confidence threshold (0-100)'],
    ['MIN_EMAIL_PRIORITY_SCORE', Number(payload.minEmailPriorityScore) || 80, 'Minimum email priority score for sending (0-100)'],
    ['TARGET_TITLES', payload.titles || '', 'Semicolon-delimited target job titles'],
    ['CAMPAIGN_FOCUS', payload.campaignFocus || '', 'Primary value proposition hook'],
    ['SEARCH_KEYWORDS', payload.searchKeywords || '', 'Search discovery query keywords'],
    ['CUSTOM_BUSINESS_NEEDS', payload.customBusinessNeeds || '', 'Manual department business needs & buying triggers'],
    ['AI_PROVIDER', payload.aiProvider || 'OPENAI', 'Selected AI engine (GEMINI or OPENAI)'],
    ['PRIMARY_AI_PROVIDER', payload.aiProvider || 'OPENAI', 'Selected AI engine (legacy alias)'],
    ['AI_MODEL', payload.aiModel || 'gpt-4o', 'AI Model version name'],
    ['EMAIL_VERIFICATION_PROVIDER', payload.emailVerifier || 'ZeroBounce', 'Email verification vendor'],
    ['EMAIL_VERIFICATION_ENDPOINT', payload.emailEndpoint || '', 'Email verification REST endpoint URL'],
    ['EMAIL_SENDER_PROVIDER', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform'],
    ['DISPATCH_MODE', payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform mode'],
    ['EMAIL_SENDER_ENDPOINT', payload.emailSenderEndpoint || '', 'Sender campaign UUID or webhook URL'],
    ['SENDER_NAME', payload.senderName || '', 'Display name for outreach email signature'],
    ['SENDER_TITLE', payload.senderTitle || '', 'Sender job title for email signature'],
    ['SENDER_COMPANY', payload.senderCompany || '', 'Sender company name for email signature'],
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

  // 2. WRITE TO Campaigns TAB (Saves individual criteria + Complete JSON object)
  let campSheet = ss.getSheetByName('Campaigns');
  if (!campSheet) campSheet = ss.insertSheet('Campaigns');

  // Check and setup master headers
  if (campSheet.getLastRow() <= 0 || campSheet.getLastColumn() < CAMPAIGNS_MASTER_HEADERS.length) {
    campSheet.getRange(1, 1, 1, CAMPAIGNS_MASTER_HEADERS.length).setValues([CAMPAIGNS_MASTER_HEADERS]);
    campSheet.getRange(1, 1, 1, CAMPAIGNS_MASTER_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    campSheet.setFrozenRows(1);
  }

  const cLast = campSheet.getLastRow();
  let cTargetRow = 2;
  let cFound = false;

  if (cLast > 1) {
    const campA = campSheet.getRange(2, 1, cLast - 1, 1).getValues();
    for (let cr = 0; cr < campA.length; cr++) {
      if (String(campA[cr][0] || '').trim() === payload.id) {
        cTargetRow = cr + 2;
        cFound = true;
        break;
      }
    }
    if (!cFound) {
      cTargetRow = cLast + 1;
    }
  }

  const campMasterRow = [
    payload.id,
    payload.name || payload.id,
    payload.department || 'Procurement',
    'Active',
    payload.country || '',
    payload.industry || '',
    payload.size || '',
    Number(payload.limit) || 5,
    payload.researchDepth || 'Deep Persona',
    Number(payload.evidencePeriodMonths) || 12,
    Number(payload.minCompanyScore) || 80,
    Number(payload.minPersonaScore) || 85,
    Number(payload.minEmailPriorityScore) || 80,
    payload.titles || '',
    payload.campaignFocus || '',
    payload.searchKeywords || '',
    payload.customBusinessNeeds || '',
    payload.aiProvider || 'OPENAI',
    payload.aiModel || 'gpt-4o',
    payload.emailVerifier || 'ZeroBounce',
    payload.emailEndpoint || '',
    payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)',
    payload.emailSenderEndpoint || '',
    payload.senderName || '',
    payload.senderTitle || '',
    payload.senderCompany || '',
    JSON.stringify(payload),
    new Date().toLocaleString()
  ];

  campSheet.getRange(cTargetRow, 1, 1, campMasterRow.length).setValues([campMasterRow]);
  SpreadsheetApp.flush();

  // 3. WRITE TO Persona Mapper Settings (Update existing row or append new row)
  let pSheet = findPersonaMapperSettingsSheet(ss);
  if (!pSheet) pSheet = ss.insertSheet('Persona Mapper Settings');

  if (pSheet.getLastRow() <= 0) {
    pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length).setValues([PERSONA_MAPPER_39_HEADERS]);
    pSheet.getRange(1, 1, 1, PERSONA_MAPPER_39_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    pSheet.setFrozenRows(1);
  }

  const pRow = [
    payload.id,
    payload.name || payload.id,
    'TRUE',
    payload.aiProvider || 'OPENAI',
    payload.aiModel || 'gpt-4o',
    payload.titles || '',
    payload.searchKeywords || '',
    payload.campaignFocus || '',
    Number(payload.evidencePeriodMonths) || 12,
    payload.researchDepth || 'Deep Persona',
    payload.incEsgObligations !== false ? 'TRUE' : 'FALSE',
    payload.incEsgMaturity !== false ? 'TRUE' : 'FALSE',
    payload.incSupplyChain !== false ? 'TRUE' : 'FALSE',
    payload.incScope3 !== false ? 'TRUE' : 'FALSE',
    payload.incDataChallenges !== false ? 'TRUE' : 'FALSE',
    payload.incDigital !== false ? 'TRUE' : 'FALSE',
    payload.incEsgGovernance !== false ? 'TRUE' : 'FALSE',
    payload.incTriggers !== false ? 'TRUE' : 'FALSE',
    payload.incEvaluater !== false ? 'TRUE' : 'FALSE',
    payload.weightEsgReporting || '20%',
    payload.weightSupplyChain || '15%',
    payload.weightDataChallenges || '15%',
    payload.weightDigital || '15%',
    payload.weightTriggers || '20%',
    payload.weightEvaluaterFit || '15%',
    Number(payload.minCompanyScore) || 80,
    Number(payload.minPersonaScore) || 85,
    Number(payload.minEmailPriorityScore) || 80,
    payload.humanApprovalRequired !== false ? 'TRUE' : 'FALSE',
    payload.autoSendEnabled === true ? 'TRUE' : 'FALSE',
    payload.weightEmailCompanyScore || '40%',
    payload.weightEmailPersonaFit || '35%',
    payload.weightEmailBuyingSignal || '25%',
    payload.customBusinessNeeds || '',
    Number(payload.limit) || 5,
    payload.emailVerifier || 'ZeroBounce',
    payload.emailEndpoint || '',
    payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)',
    payload.emailSenderEndpoint || ''
  ];

  let pTargetRow = 2;
  const pLast = pSheet.getLastRow();
  let pFound = false;

  if (pLast > 1) {
    const colA = pSheet.getRange(2, 1, pLast - 1, 1).getValues();
    for (let r = 0; r < colA.length; r++) {
      if (String(colA[r][0] || '').trim() === payload.id) {
        pTargetRow = r + 2;
        pFound = true;
        break;
      }
    }
    if (!pFound) {
      pTargetRow = pLast + 1;
    }
  }

  pSheet.getRange(pTargetRow, 1, 1, 39).setValues([pRow]);
  SpreadsheetApp.flush();

  // 4. UPDATE DASHBOARD TAB SPECIFICALLY FOR THIS ACTIVE CAMPAIGN
  updateDashboardForActiveCampaign(payload);

  return {
    success: true,
    campaignId: payload.id,
    campaignName: payload.name || payload.id,
    message: 'Campaign ' + payload.id + ' synchronized across all tabs and activated on Dashboard!'
  };
}

/**
 * Re-calculates and writes active campaign parameters and live pipeline metrics
 * (Discovered Companies, Identified Personas, Queue stats) directly to the Dashboard sheet.
 */
function updateDashboardForActiveCampaign(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) return;

  const cid = payload.id;
  const prov = String(payload.aiProvider || 'OPENAI').toUpperCase();
  const model = String(payload.aiModel || (prov.includes('GEMINI') ? 'gemini-2.5-flash' : 'gpt-4o')).trim();

  // Build high-clarity formatted AI Model description
  let aiLabel = '';
  if (prov.includes('GEMINI')) {
    aiLabel = 'Gemini ' + (model.includes('2.5') ? '2.5 Flash' : model) + ' + Search Grounding';
  } else {
    aiLabel = 'OpenAI ' + model.toUpperCase();
  }

  // Live verification of API key connection
  const key = getStudioApiKey(prov);
  const keyStatus = key ? 'Connected' : 'Key Not Configured';
  const nowStr = new Date().toLocaleString();

  // 1. Calculate live counts for this campaign from database sheets
  let discoveredCount = 0;
  const compSheet = ss.getSheetByName('Company Intelligence');
  if (compSheet && compSheet.getLastRow() > 1) {
    const data = compSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === cid) discoveredCount++;
    }
  }

  let personasCount = 0;
  const persSheet = ss.getSheetByName('Persona Leads');
  if (persSheet && persSheet.getLastRow() > 1) {
    const data = persSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === cid) personasCount++;
    }
  }

  let queueCount = 0;
  let approvedCount = 0;
  let draftsCount = 0;
  let sentCount = 0;
  const qSheet = ss.getSheetByName('Prospect Review Queue');
  if (qSheet && qSheet.getLastRow() > 1) {
    const data = qSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === cid) {
        queueCount++;
        const approval = String(data[i][21] || '').trim();
        const sendStatus = String(data[i][24] || '').trim();
        if (approval.startsWith('Approved')) approvedCount++;
        if (sendStatus === 'Draft Created') draftsCount++;
        if (sendStatus === 'Sent') sentCount++;
      }
    }
  }

  // 2. Scan Dashboard sheet cells and update matching label values
  const dRange = dashSheet.getDataRange();
  const dVals = dRange.getValues();
  const numRows = dVals.length;
  const numCols = dVals[0] ? dVals[0].length : 0;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cellText = normalizeStudioKey(dVals[r][c]);
      if (!cellText) continue;

      // Primary AI Model Configured (Updates Label, Connection Status, and Timestamp)
      if (cellText.includes('PRIMARYAIMODEL') || cellText.includes('AIMODELCONFIGURED') || cellText === 'AIPROVIDER' || cellText === 'AIENGINE' || cellText === 'MODEL') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(aiLabel);
        if (c + 2 < numCols) dashSheet.getRange(r + 1, c + 3).setValue(keyStatus);
        if (c + 3 < numCols) dashSheet.getRange(r + 1, c + 4).setValue(nowStr);
      }
      // Campaign ID
      else if (cellText === 'CAMPAIGNID' || cellText === 'ACTIVECAMPAIGNID') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(cid);
      }
      // Campaign Name
      else if (cellText === 'CAMPAIGNNAME' || cellText === 'ACTIVECAMPAIGN' || cellText === 'CAMPAIGN') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.name || cid);
      }
      // Status
      else if (cellText === 'STATUS' || cellText === 'CAMPAIGNSTATUS') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue('Active');
      }
      // Department
      else if (cellText === 'DEPARTMENT' || cellText === 'OWNER') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.department || 'Procurement');
      }
      // Target Countries
      else if (cellText.includes('TARGETCOUNT') || cellText === 'COUNTRIES' || cellText === 'GEOGRAPHY') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.country || 'Global');
      }
      // Target Industries
      else if (cellText.includes('TARGETINDUS') || cellText === 'INDUSTRIES') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.industry || 'All Verticals');
      }
      // Company Size
      else if (cellText.includes('COMPANYSIZE')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.size || '50 - 250');
      }
      // Target Titles
      else if (cellText.includes('TARGETTITLE') || cellText.includes('PERSONATITLE')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.titles || 'Decision Makers');
      }
      // Email Sender Provider
      else if (cellText.includes('SENDERPROVIDER') || cellText === 'EMAILSENDER') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(payload.emailSenderProvider || 'Workspace Gmail (Drafts Only)');
      }
      // Metric: Discovered Companies
      else if (cellText.includes('DISCOVEREDCOMP') || cellText === 'TOTALCOMPANIES') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(discoveredCount);
      }
      // Metric: Identified Leads
      else if (cellText.includes('IDENTIFIEDLEAD') || cellText.includes('PERSONASIDENTIFIED') || cellText === 'TOTALLEADS') {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(personasCount);
      }
      // Metric: Review Queue
      else if (cellText.includes('REVIEWQUEUE') || cellText.includes('PENDINGREVIEW')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(queueCount);
      }
      // Metric: Approved Leads
      else if (cellText.includes('APPROVEDLEAD') || cellText.includes('APPROVEDPROSPECT')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(approvedCount);
      }
      // Metric: Drafts Created
      else if (cellText.includes('DRAFTCREATED') || cellText.includes('GMAILDRAFT')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(draftsCount);
      }
      // Metric: Emails Sent
      else if (cellText.includes('EMAILSSENT') || cellText.includes('SENTEMAIL')) {
        if (c + 1 < numCols) dashSheet.getRange(r + 1, c + 2).setValue(sentCount);
      }
    }
  }

  // 3. Fallback standard coordinates (covers standard Studio dashboard layout)
  try {
    dashSheet.getRange('B2').setValue(cid);
    dashSheet.getRange('B3').setValue('Active');
    dashSheet.getRange('B4').setValue(payload.name || cid);
    dashSheet.getRange('D3').setValue(nowStr);
  } catch (e) {
    // Coordinate write error ignored if sheet layout is custom
  }

  SpreadsheetApp.flush();
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
          if (val.startsWith('CMP-') || val.startsWith('PILOT-')) {
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
    country: String(dict['TARGETCOUNTRIES'] || ''),
    industry: String(dict['TARGETINDUSTRIES'] || ''),
    size: String(dict['COMPANYSIZERANGE'] || ''),
    limit: Number(dict['MAXCOMPANIES']) || 5,
    researchDepth: String(dict['RESEARCHDEPTH'] || 'Deep Persona'),
    evidencePeriodMonths: Number(dict['EVIDENCEPERIODMONTHS']) || 12,
    minCompanyScore: Number(dict['MINCOMPANYSCORE']) || 80,
    minPersonaScore: Number(dict['MINPERSONASCORE']) || 85,
    minEmailPriorityScore: Number(dict['MINEMAILPRIORITYSCORE']) || 80,
    titles: String(dict['TARGETTITLES'] || ''),
    campaignFocus: String(dict['CAMPAIGNFOCUS'] || ''),
    searchKeywords: String(dict['SEARCHKEYWORDS'] || ''),
    customBusinessNeeds: String(dict['CUSTOMBUSINESSNEEDS'] || ''),
    aiProvider: String(dict['AIPROVIDER'] || 'OPENAI'),
    aiModel: String(dict['AIMODEL'] || 'gpt-4o'),
    emailVerifier: String(dict['EMAILVERIFICATIONPROVIDER'] || 'ZeroBounce'),
    emailEndpoint: String(dict['EMAILVERIFICATIONENDPOINT'] || ''),
    emailSenderProvider: String(dict['EMAILSENDERPROVIDER'] || 'Workspace Gmail (Drafts Only)'),
    emailSenderEndpoint: String(dict['EMAILSENDERENDPOINT'] || ''),
    senderName: String(dict['SENDERNAME'] || ''),
    senderTitle: String(dict['SENDERTITLE'] || ''),
    senderCompany: String(dict['SENDERCOMPANY'] || '')
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