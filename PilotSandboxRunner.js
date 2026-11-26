/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AUTHENTIC PIPELINE PILOT RUNNER (CLIENT-ISOLATED SANDBOX SIMULATION)
 * ============================================================================
 * File: PilotSandboxRunner.gs / PilotSandboxRunner.js
 *
 * CAPABILITIES:
 * 1. Generates & Synchronizes ALL Project Tabs for the Pilot Run (Client: "Pilot Acme"):
 *    - Cold-start safe: Automatically seeds canonical headers across all 13 core tabs
 *      plus EmailEvents & Email Deliverability.
 *    - Updates "Campaign Configuration" strictly preserving the locked 38 parameters.
 *    - Generates 39-column row in "Persona Mapper Settings" for PILOT-CMP-001.
 *    - Restores and populates "Dashboard" with the complete 11-row KPI table.
 *    - Populates "Campaigns" (28 master cols), "Company Database" (Col J Campaign ID),
 *      "Company Intelligence" (25 cols), "Persona Leads" (29 cols),
 *      "Research Evidence" (9 canonical cols), "Prospect Review Queue" (45 cols),
 *      "EmailEvents" (11 cols), and logs client-tagged token/credit usage to "API_Usage_Log".
 * 2. Creates Safe Gmail Drafts (via Pilot Acme Relay if configured, or local sandbox Gmail).
 * 3. Enforces real-time suppression rules and pins "Lists & Settings" client rows cleanly.
 * 4. Option A Full Wipe Purge: Removes all pilot records (PILOT-* and Pilot Acme)
 *    across all tabs including EmailEvents, Dashboard Rollup, Agent Settings,
 *    Lists & Settings, Campaign Configuration, API_Usage_Log, and Execution Logs,
 *    then restores active state.
 * ============================================================================
 */

const PILOT_CONFIG = {
  CAMPAIGN_ID: 'PILOT-CMP-001',
  CAMPAIGN_NAME: '[PILOT] Enterprise Compliance & Operations Readiness',
  DEPARTMENT: 'Pilot Acme',
  COMPANY_LIMIT: 5,
  RESEARCH_DEPTH: 'Deep Persona',
  MIN_COMPANY_SCORE: 80,
  MIN_PERSONA_SCORE: 85,
  SENDER_NAME: 'Alex Mercer',
  SENDER_TITLE: 'Solutions Consultant',
  SENDER_COMPANY: 'eValuater Solutions',
  SENDER_EMAIL: 'outreach@pilotacme.com',
  REPLY_TO_EMAIL: 'outreach@pilotacme.com',
  CUSTOM_BUSINESS_NEEDS: 'Automated Scope 3 emissions reporting, supplier audit evidence collection, and CSRD compliance readiness',
  TABS: {
    DASHBOARD: 'Dashboard',
    API_USAGE_LOG: 'API_Usage_Log',
    CAMPAIGNS: 'Campaigns',
    COMPANY_INTELLIGENCE: 'Company Intelligence',
    PERSONAS: 'Persona Leads',
    PERSONA_SETTINGS: 'Persona Mapper Settings',
    REVIEW_QUEUE: 'Prospect Review Queue',
    CAMPAIGN_CONFIG: 'Campaign Configuration',
    COMPANY_DATABASE: 'Company Database',
    RESEARCH_EVIDENCE: 'Research Evidence',
    AGENT_SETTINGS: 'Agent Settings',
    LOGS: 'Execution Logs',
    LISTS_SETTINGS: 'Lists & Settings'
  }
};

/**
 * MAIN ENTRY POINT: Runs the full pilot for client "Pilot Acme", logs API usage,
 * seeds EmailEvents deliverability samples, and updates Dashboard.
 */
function runPilotOnRealSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  Logger.log('===============================================================');
  Logger.log('>>> STARTING FULL WORKSPACE PILOT: ' + PILOT_CONFIG.CAMPAIGN_ID + ' (Client: ' + PILOT_CONFIG.DEPARTMENT + ') <<<');
  Logger.log('===============================================================');

  try {
    // Step 0: Ensure all core tabs + EmailEvents & Email Deliverability tabs exist with canonical headers
    validateRequiredTabs(ss);

    // Step 1: Clean previous pilot test rows from all operational tabs (including EmailEvents)
    purgeExistingPilotRows(ss, PILOT_CONFIG.CAMPAIGN_ID);
    Logger.log('Step 1: Purged previous pilot records.');

    // Step 2: Ensure client entries, sort Lists & Settings, and load suppression rules
    ensureSuppressionRulesInSettings(ss);
    if (typeof refreshClientDropdownValidations_ === 'function') {
      refreshClientDropdownValidations_(ss);
    }
    const suppressionRules = loadSuppressionRulesFromSettings(ss);
    Logger.log('Step 2: Loaded suppression rules (' + suppressionRules.domains.length + ' domains, ' + suppressionRules.emails.length + ' emails).');

    // Step 3: Update "Campaign Configuration" tab with standardized Pilot parameters (locked 38 rows)
    updateCampaignConfigurationTab(ss, PILOT_CONFIG);
    Logger.log('Step 3: "Campaign Configuration" updated for client ' + PILOT_CONFIG.DEPARTMENT + ' with Company Limit: ' + PILOT_CONFIG.COMPANY_LIMIT);

    // Step 4: Generate 39-column settings row in "Persona Mapper Settings"
    generatePersonaMapperSettingsRow(ss, PILOT_CONFIG);
    Logger.log('Step 4: "Persona Mapper Settings" generated for ' + PILOT_CONFIG.CAMPAIGN_ID);

    // Step 5: Write Pilot Campaign Record to "Campaigns" (28 Master Columns with embedded JSON)
    writePilotCampaign(ss, PILOT_CONFIG);
    Logger.log('Step 5: Master campaign record written to "Campaigns" with limit ' + PILOT_CONFIG.COMPANY_LIMIT);

    // Step 6: Write Master Company Directory to "Company Database" (Stamped with Col J Campaign ID)
    const rawCompanies = getPilotCompanyDatabaseRecords(PILOT_CONFIG);
    writePilotCompanyDatabase(ss, rawCompanies);
    Logger.log('Step 6: Master companies written to "Company Database" with Column J Campaign ID.');

    // Step 7: Write 25-Column Enriched Intelligence to "Company Intelligence"
    const enrichedCompanies = getPilotCompaniesData(PILOT_CONFIG);
    writePilotCompanyIntelligence(ss, enrichedCompanies);
    Logger.log('Step 7: Enriched intelligence mapped to "Company Intelligence".');

    // Step 8: Write Grounding Evidence to "Research Evidence" (9 Canonical Columns)
    const researchEvidence = getPilotResearchEvidence(PILOT_CONFIG);
    writePilotResearchEvidence(ss, researchEvidence);
    Logger.log('Step 8: Search grounding evidence logged in "Research Evidence".');

    // Step 9: Write Persona Leads & Outreach Copy to "Persona Leads" (29 Columns)
    const pilotPersonas = getPilotPersonasData(PILOT_CONFIG);
    writePilotPersonaLeads(ss, pilotPersonas);
    Logger.log('Step 9: Persona leads written to "Persona Leads".');

    // Step 10: Security Audit, Auto-Approval & Gmail Draft Dispatch
    const processedQueue = processSecurityApprovalAndDrafts(pilotPersonas, enrichedCompanies, suppressionRules, PILOT_CONFIG);
    writePilotReviewQueue(ss, processedQueue);
    Logger.log('Step 10: Staged and processed leads in "Prospect Review Queue".');

    // Step 11: Log API Usage & Estimated Cost to "API_Usage_Log" tagged with Client "Pilot Acme"
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: PILOT_CONFIG.CAMPAIGN_ID,
        clientName: PILOT_CONFIG.DEPARTMENT,
        serviceName: 'GEMINI',
        operation: 'PilotSandboxDiscovery',
        httpStatus: 200,
        inputUnits: 1850,
        outputUnits: 920,
        status: 'SUCCESS',
        details: 'Pilot simulation grounding search & company enrichment',
        modelName: 'gemini-3.8-flash'
      });
      logApiUsage_({
        campaignId: PILOT_CONFIG.CAMPAIGN_ID,
        clientName: PILOT_CONFIG.DEPARTMENT,
        serviceName: 'ZeroBounce',
        operation: 'EmailVerification',
        httpStatus: 200,
        inputUnits: 3,
        outputUnits: 0,
        status: 'SUCCESS',
        details: 'Pilot verification check for 3 test leads',
        modelName: 'ZeroBounce'
      });
    }
    Logger.log('Step 11: API token and verifier usage logged to "API_Usage_Log".');

    // Step 11b: Seed sample EmailEvents rows for Pilot Acme to verify Email Deliverability Dashboard
    writePilotDeliverabilityEvents_(ss, PILOT_CONFIG);
    if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
      refreshDeliverabilityFilterDropdowns_(ss);
    }
    Logger.log('Step 11b: Deliverability pilot events written to "EmailEvents".');

    // Step 12: Log Execution in "Execution Logs"
    logPilotExecution(ss, processedQueue);

    // Step 13: Generate and Update Master Dashboard
    const draftsCreated = processedQueue.filter(function(p) { return p['Send Status'] === 'Draft Created in Gmail'; }).length;
    const blockedCount = processedQueue.filter(function(p) { return p['Compliance Status'] !== 'Passed'; }).length;
    updatePilotDashboard(ss, PILOT_CONFIG, enrichedCompanies.length, pilotPersonas.length, draftsCreated, blockedCount);
    Logger.log('Step 13: Master Dashboard updated with live KPIs for client ' + PILOT_CONFIG.DEPARTMENT + '.');

    ss.toast(
      'Pilot for "' + PILOT_CONFIG.DEPARTMENT + '" completed! ' + draftsCreated + ' Gmail drafts created & sample EmailEvents seeded. Check Dashboard, Email Deliverability & Gmail.',
      '🚀 Pilot Complete',
      12
    );

    Logger.log('===============================================================');
    Logger.log('>>> PILOT RUN COMPLETED: ALL TABS, DELIVERABILITY & API TRACKER SYNCHRONIZED <<<');
    Logger.log('===============================================================');
  } catch (err) {
    Logger.log('❌ Pilot Execution Error: ' + err.stack);
    ss.toast('Pilot error: ' + err.message, 'Pilot Error', 10);
  }
}

/**
 * Validates that all 13 core project tabs plus EmailEvents & Email Deliverability exist in canonical layout.
 */
function validateRequiredTabs(ss) {
  const definitions = (typeof APP_CONFIG !== 'undefined' && Array.isArray(APP_CONFIG.sheetDefinitions))
    ? APP_CONFIG.sheetDefinitions
    : [];

  Object.values(PILOT_CONFIG.TABS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    if (name === PILOT_CONFIG.TABS.DASHBOARD) {
      return;
    }

    if (name === PILOT_CONFIG.TABS.AGENT_SETTINGS) {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Category', 'Setting Key', 'Value', 'Status / Description']);
      }
    } else {
      const def = definitions.find(function(d) { return d.name === name; });
      if (sheet.getLastRow() === 0 && def && Array.isArray(def.headers) && def.headers.length > 0) {
        sheet.appendRow(def.headers);
      }
    }

    if (sheet.getLastColumn() > 0) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn())
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(1);
    }
  });

  if (typeof setupEmailEventsTab_ === 'function') {
    setupEmailEventsTab_(ss);
  }
  if (typeof setupEmailDeliverabilityDashboardTab_ === 'function') {
    if (!ss.getSheetByName('Email Deliverability')) {
      setupEmailDeliverabilityDashboardTab_(ss);
    } else if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
      refreshDeliverabilityFilterDropdowns_(ss);
    }
  }
}

/**
 * Seeds 11 clean sandbox events in "EmailEvents" for PILOT-CMP-001 ("Pilot Acme")
 * so all Email Deliverability tiles, rates, SMTP error charts & daily volume charts can be verified.
 */
function writePilotDeliverabilityEvents_(ss, config) {
  let eventsSheet = (typeof setupEmailEventsTab_ === 'function')
    ? setupEmailEventsTab_(ss)
    : ss.getSheetByName('EmailEvents');
  if (!eventsSheet) return;

  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const yesterdayPlus = new Date(yesterday.getTime() + 10 * 60 * 1000);
  const todayEarlier = new Date(now.getTime() - 5 * 60 * 60 * 1000);
  const todayDelivered = new Date(todayEarlier.getTime() + 15 * 60 * 1000);
  const todayReply = new Date(now.getTime() - 2 * 60 * 60 * 1000);

  const sampleRows = [
    ['PILOT-EVT-001', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-1@pilotacme.com>', 'Sent', yesterday, 'marcus.vance@growthtarget.com', '250 2.0.0', '', 'Pilot simulated outreach dispatch'],
    ['PILOT-EVT-002', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-1@pilotacme.com>', 'Delivered', yesterdayPlus, 'marcus.vance@growthtarget.com', '250 2.0.0', '', 'No DSN bounce received within 4h window'],
    ['PILOT-EVT-003', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-1@pilotacme.com>', 'Replied', todayReply, 'marcus.vance@growthtarget.com', '', '', 'Reply received: Re: Apex infrastructure sprint'],
    ['PILOT-EVT-004', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-2@pilotacme.com>', 'Sent', yesterday, 'elena.rostova@novussystems.com', '250 2.0.0', '', 'Pilot simulated outreach dispatch'],
    ['PILOT-EVT-005', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-2@pilotacme.com>', 'Delivered', yesterdayPlus, 'elena.rostova@novussystems.com', '250 2.0.0', '', 'No DSN bounce received within 4h window'],
    ['PILOT-EVT-006', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-3@pilotacme.com>', 'Sent', todayEarlier, 'invalid.user@novussystems.com', '250 2.0.0', '', 'Pilot simulated outreach dispatch'],
    ['PILOT-EVT-007', config.DEPARTMENT, config.CAMPAIGN_ID, 'Gmail', '<pilot-msg-3@pilotacme.com>', 'Bounced', todayDelivered, 'invalid.user@novussystems.com', '550 5.1.1', 'Hard', '550 5.1.1 The email account that you tried to reach does not exist'],
    ['PILOT-EVT-008', config.DEPARTMENT, config.CAMPAIGN_ID, 'Microsoft 365', '<pilot-msg-4@pilotacme.com>', 'Sent', todayEarlier, 'throttled.lead@example.com', '250 2.0.0', '', 'Pilot simulated outreach dispatch'],
    ['PILOT-EVT-009', config.DEPARTMENT, config.CAMPAIGN_ID, 'Microsoft 365', '<pilot-msg-4@pilotacme.com>', 'Bounced', todayDelivered, 'throttled.lead@example.com', '451 4.7.1', 'Soft', '451 4.7.1 Greylisting in action, please try again later'],
    ['PILOT-EVT-010', config.DEPARTMENT, config.CAMPAIGN_ID, 'Microsoft 365', '<pilot-msg-5@pilotacme.com>', 'Sent', todayEarlier, 'ops.lead@example.com', '250 2.0.0', '', 'Pilot simulated outreach dispatch'],
    ['PILOT-EVT-011', config.DEPARTMENT, config.CAMPAIGN_ID, 'Microsoft 365', '<pilot-msg-5@pilotacme.com>', 'Delivered', todayDelivered, 'ops.lead@example.com', '250 2.0.0', '', 'Confirmed Delivered via Exchange Message Trace']
  ];

  const startRow = eventsSheet.getLastRow() + 1;
  eventsSheet.getRange(startRow, 7, sampleRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  eventsSheet.getRange(startRow, 9, sampleRows.length, 1).setNumberFormat('@');
  eventsSheet.getRange(startRow, 1, sampleRows.length, 11).setValues(sampleRows);
}

/**
 * Updates "Campaign Configuration" without adding extra rows (preserves locked 38 parameters).
 */
function updateCampaignConfigurationTab(ss, config) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGN_CONFIG);
  if (!sheet) return;

  const updates = {
    'CAMPAIGN_ID': config.CAMPAIGN_ID,
    'CAMPAIGN_NAME': config.CAMPAIGN_NAME,
    'DEPARTMENT': config.DEPARTMENT,
    'TARGET_COUNTRIES': 'United States & Canada',
    'TARGET_INDUSTRIES': 'Freight & Logistics, Industrial Equipment',
    'COMPANY_SIZE_RANGE': '50 - 250 employees',
    'MAX_COMPANIES': config.COMPANY_LIMIT,
    'RESEARCH_DEPTH': config.RESEARCH_DEPTH,
    'EVIDENCE_PERIOD_MONTHS': 12,
    'MIN_COMPANY_SCORE': config.MIN_COMPANY_SCORE,
    'MIN_PERSONA_SCORE': config.MIN_PERSONA_SCORE,
    'MIN_EMAIL_PRIORITY_SCORE': 80,
    'TARGET_TITLES': 'VP of Technology; Head of Engineering; Director of Sustainability; CISO; VP Procurement',
    'CAMPAIGN_FOCUS': 'Scope 3 vendor compliance and operational resilience',
    'SEARCH_KEYWORDS': 'Scope 3, CSRD, AWS Cloud, compliance automation, supplier audits',
    'CUSTOM_BUSINESS_NEEDS': config.CUSTOM_BUSINESS_NEEDS,
    'AI_PROVIDER': 'GEMINI',
    'PRIMARY_AI_PROVIDER': 'GEMINI',
    'AI_MODEL': 'gemini-3.8-flash',
    'EMAIL_VERIFICATION_PROVIDER': 'ZeroBounce',
    'EMAIL_VERIFICATION_ENDPOINT': 'https://api.zerobounce.net/v2/validate',
    'EMAIL_SENDER_PROVIDER': 'Workspace Gmail (Drafts Only)',
    'DISPATCH_MODE': 'Workspace Gmail (Drafts Only)',
    'EMAIL_SENDER_ENDPOINT': '',
    'SENDER_NAME': config.SENDER_NAME,
    'SENDER_TITLE': config.SENDER_TITLE,
    'SENDER_COMPANY': config.SENDER_COMPANY,
    'INCLUDE_ESG_OBLIGATIONS': 'TRUE',
    'INCLUDE_ESG_MATURITY': 'TRUE',
    'INCLUDE_SUPPLY_CHAIN': 'TRUE',
    'INCLUDE_SCOPE3': 'TRUE',
    'INCLUDE_DATA_CHALLENGES': 'TRUE',
    'INCLUDE_DIGITAL': 'TRUE',
    'INCLUDE_ESG_GOVERNANCE': 'TRUE',
    'INCLUDE_TRIGGERS': 'TRUE',
    'INCLUDE_EVALUATER': 'TRUE',
    'HUMAN_APPROVAL_REQUIRED': 'TRUE',
    'AUTO_SEND_ENABLED': 'FALSE'
  };

  if (sheet.getLastRow() <= 1) {
    const rows = [['Parameter', 'Configured Value', 'Description']];
    Object.keys(updates).forEach(function(key) {
      rows.push([key, updates[key], 'Pilot configured parameter']);
    });
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const key = String(data[r][0] || '').trim().toUpperCase();
      if (Object.prototype.hasOwnProperty.call(updates, key)) {
        sheet.getRange(r + 1, 2).setValue(updates[key]);
      }
    }
  }
}

function generatePersonaMapperSettingsRow(ss, config) {
  const sheet = (typeof findPersonaMapperSettingsSheet === 'function')
    ? findPersonaMapperSettingsSheet(ss)
    : ss.getSheetByName(PILOT_CONFIG.TABS.PERSONA_SETTINGS);
  if (!sheet) return;

  const headers = [
    'Campaign ID','Campaign Name','Persona Mapper Enabled','AI Provider','AI Model','Persona Titles',
    'Search Keywords','Campaign Focus','Evidence Period Months','Research Depth','Include ESG Reporting Obligations',
    'Include ESG Reporting Maturity','Include Supply Chain Complexity','Include Scope 3 Exposure','Include Sustainability Data Challenges',
    'Include Digital Transformation','Include ESG Governance','Include Business Triggers','Include eValuater Solution Fit',
    'Weight ESG Reporting Relevance','Weight Supply Chain Complexity','Weight Sustainability Data Needs','Weight Digital Transformation Relevance',
    'Weight Business Triggers','Weight eValuater Solution Fit','Minimum Company Opportunity Score','Minimum Persona Confidence',
    'Minimum Email Priority Score','Human Approval Required','Auto Send Enabled','Email Priority Weight Company Score',
    'Email Priority Weight Persona Fit','Email Priority Weight Buying Signal','Notes','Max Companies Per Run',
    'Email Verification Provider','Email Verification Endpoint','Email Sender Provider','Email Sender Endpoint'
  ];

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < headers.length) {
    sheet.clear();
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][0] || '').trim() === config.CAMPAIGN_ID) {
      sheet.deleteRow(r + 1);
    }
  }

  const row = [
    config.CAMPAIGN_ID,
    config.CAMPAIGN_NAME,
    'TRUE',
    'GEMINI',
    'gemini-3.8-flash',
    'VP of Technology, Head of Engineering, Director of Sustainability, CISO, VP Procurement',
    'Scope 3, CSRD, AWS Cloud, compliance automation, supplier audits',
    config.CUSTOM_BUSINESS_NEEDS,
    12,
    config.RESEARCH_DEPTH,
    'TRUE', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'TRUE',
    '20%', '15%', '15%', '15%', '20%', '15%',
    config.MIN_COMPANY_SCORE,
    config.MIN_PERSONA_SCORE,
    80,
    'TRUE',
    'FALSE',
    '40%', '35%', '25%',
    'Automated pilot configuration',
    config.COMPANY_LIMIT,
    'ZeroBounce',
    'https://api.zerobounce.net/v2/validate',
    'Workspace Gmail (Drafts Only)',
    ''
  ];

  const targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 9).setNumberFormat('0');
  sheet.getRange(targetRow, 26, 1, 3).setNumberFormat('0');
  sheet.getRange(targetRow, 35).setNumberFormat('0');
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
}

function writePilotCampaign(ss, config) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGNS);
  if (!sheet) return;

  const nowStr = new Date().toLocaleString();

  const campaignPayload = {
    id: config.CAMPAIGN_ID,
    name: config.CAMPAIGN_NAME,
    department: config.DEPARTMENT,
    country: 'United States & Canada',
    industry: 'Freight & Logistics, Industrial Equipment',
    size: '50 - 250 employees',
    limit: config.COMPANY_LIMIT,
    researchDepth: config.RESEARCH_DEPTH,
    evidencePeriodMonths: 12,
    minCompanyScore: config.MIN_COMPANY_SCORE,
    minPersonaScore: config.MIN_PERSONA_SCORE,
    minEmailPriorityScore: 80,
    titles: 'VP of Technology; Head of Engineering; Director of Sustainability; CISO; VP Procurement',
    campaignFocus: 'Scope 3 vendor compliance and operational resilience',
    searchKeywords: 'Scope 3, CSRD, AWS Cloud, compliance automation, supplier audits',
    customBusinessNeeds: config.CUSTOM_BUSINESS_NEEDS,
    aiProvider: 'GEMINI',
    aiModel: 'gemini-3.8-flash',
    emailVerifier: 'ZeroBounce',
    emailEndpoint: 'https://api.zerobounce.net/v2/validate',
    emailSenderProvider: 'Workspace Gmail (Drafts Only)',
    emailSenderEndpoint: '',
    senderName: config.SENDER_NAME,
    senderTitle: config.SENDER_TITLE,
    senderCompany: config.SENDER_COMPANY,
    senderEmail: config.SENDER_EMAIL || '',
    replyToEmail: config.REPLY_TO_EMAIL || config.SENDER_EMAIL || ''
  };

  const master28Row = [
    config.CAMPAIGN_ID,
    config.CAMPAIGN_NAME,
    config.DEPARTMENT,
    'Pilot Active',
    'United States & Canada',
    'Freight & Logistics, Industrial Equipment',
    '50 - 250 employees',
    config.COMPANY_LIMIT,
    config.RESEARCH_DEPTH,
    12,
    config.MIN_COMPANY_SCORE,
    config.MIN_PERSONA_SCORE,
    80,
    'VP of Technology; Head of Engineering; Director of Sustainability; CISO; VP Procurement',
    'Scope 3 vendor compliance and operational resilience',
    'Scope 3, CSRD, AWS Cloud, compliance automation, supplier audits',
    config.CUSTOM_BUSINESS_NEEDS,
    'GEMINI',
    'gemini-3.8-flash',
    'ZeroBounce',
    'https://api.zerobounce.net/v2/validate',
    'Workspace Gmail (Drafts Only)',
    '',
    config.SENDER_NAME,
    config.SENDER_TITLE,
    config.SENDER_COMPANY,
    JSON.stringify(campaignPayload),
    nowStr
  ];

  const campDef = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.sheetDefinitions)
    ? APP_CONFIG.sheetDefinitions.find(function(d) { return d.name === 'Campaigns'; })
    : null;

  if (sheet.getLastRow() === 0 && campDef) {
    sheet.appendRow(campDef.headers);
    sheet.getRange(1, 1, 1, 28)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  }

  const targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 8).setNumberFormat('0');
  sheet.getRange(targetRow, 10, 1, 4).setNumberFormat('0');
  sheet.getRange(targetRow, 1, 1, master28Row.length).setValues([master28Row]);
}

function getPilotCompanyDatabaseRecords(config) {
  const dateStr = new Date().toISOString().split('T')[0];
  return [
    {
      'Company ID': 'PILOT-CO-001',
      'Campaign ID': config.CAMPAIGN_ID,
      'Company Name': 'Apex Global Operations',
      'Domain': 'growthtarget.com',
      'Country': 'United States',
      'Industry': 'Freight & Logistics',
      'Employee Size': '180 employees',
      'Discovery Method': 'Search Grounding (' + config.DEPARTMENT + ')',
      'Discovery Date': dateStr,
      'Status': 'Enriched'
    },
    {
      'Company ID': 'PILOT-CO-002',
      'Campaign ID': config.CAMPAIGN_ID,
      'Company Name': 'Novus Systems Corp',
      'Domain': 'novussystems.com',
      'Country': 'Canada',
      'Industry': 'Industrial Equipment',
      'Employee Size': '140 employees',
      'Discovery Method': 'Search Grounding (' + config.DEPARTMENT + ')',
      'Discovery Date': dateStr,
      'Status': 'Enriched'
    },
    {
      'Company ID': 'PILOT-CO-003',
      'Campaign ID': config.CAMPAIGN_ID,
      'Company Name': 'EcoTransit Competitor Systems',
      'Domain': 'competitor-esg.com',
      'Country': 'United States',
      'Industry': 'Direct Competitor',
      'Employee Size': '90 employees',
      'Discovery Method': 'Grounding Scan',
      'Discovery Date': dateStr,
      'Status': 'Disqualified (Competitor)'
    }
  ];
}

function writePilotCompanyDatabase(ss, list) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.COMPANY_DATABASE);
  if (!sheet) return;

  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hasCampaignCol = currentHeaders.some(function(h) {
    return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '') === 'campaignid';
  });

  if (!hasCampaignCol) {
    sheet.getRange(1, 10).setValue('Campaign ID')
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  }

  const updatedHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 10)).getValues()[0];
  const rows = list.map(function(item) { return buildRowFromHeaders(updatedHeaders, item); });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, updatedHeaders.length).setValues(rows);
}

function getPilotCompaniesData(config) {
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const nowStr = now.toLocaleString();

  return [
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-001',
      'Company Name': 'Apex Global Operations',
      'Country': 'United States',
      'Industry': 'Freight & Logistics',
      'Company Size': '180 employees',
      'ESG Frameworks': 'CSRD, CDP, GHG Protocol Scope 1-3',
      'Reporting Obligations': 'Mandatory Tier 1 Scope 3 Reporting',
      'ESG Reporting Maturity': 'Developing / Spreadsheet-reliant',
      'Supply Chain Profile': 'Multi-tier supplier & carrier network',
      'Scope 3 Exposure': 'High (Freight fleet emissions)',
      'Sustainability Data Needs': 'Automated primary carrier data ingestion',
      'Digital Transformation Signals': 'Migrated core services to AWS; active modernization sprint',
      'ESG Governance': 'Director of Sustainability reporting to COO',
      'Business Triggers': 'Active Trigger: ' + config.CUSTOM_BUSINESS_NEEDS.slice(0, 80),
      'eValuater Solution Fit': 'High',
      'Company Opportunity Score': 94,
      'Company Opportunity': 'Urgent bottleneck: manual evidence collection pulling engineering bandwidth (Client: ' + config.DEPARTMENT + ')',
      'Evidence URLs': 'https://example.com/apex-operations-audit',
      'Evidence Checked Date': dateStr,
      'Verification Status': 'Verified',
      'Human Review Status': 'Accepted',
      'Last Enriched At': nowStr,
      'Company Score Rationale': 'High alignment with ' + config.DEPARTMENT + ' campaign criteria.'
    },
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-002',
      'Company Name': 'Novus Systems Corp',
      'Country': 'Canada',
      'Industry': 'Industrial Equipment',
      'Company Size': '140 employees',
      'ESG Frameworks': 'SEC Climate Disclosure, ISSB',
      'Reporting Obligations': 'Customer Tier-1 vendor disclosures',
      'ESG Reporting Maturity': 'Early Stage',
      'Supply Chain Profile': 'Global Tier-1 vendor network',
      'Scope 3 Exposure': 'High (Upstream supplier footprint)',
      'Sustainability Data Needs': 'Automated vendor ESG data collection',
      'Digital Transformation Signals': 'Deploying continuous CI/CD pipelines',
      'ESG Governance': 'Head of Compliance & Procurement',
      'Business Triggers': 'Customer audit deadline causing procurement backlog',
      'eValuater Solution Fit': 'High',
      'Company Opportunity Score': 88,
      'Company Opportunity': 'High supplier friction; manual evidence follow-ups delaying release cycles (Automates compliance workflow)',
      'Evidence URLs': 'https://example.com/novus-systems-report',
      'Evidence Checked Date': dateStr,
      'Verification Status': 'Verified',
      'Human Review Status': 'Accepted',
      'Last Enriched At': nowStr,
      'Company Score Rationale': 'Qualified mid-market company matching buying signals.'
    },
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-003',
      'Company Name': 'EcoTransit Competitor Systems',
      'Country': 'United States',
      'Industry': 'Direct Industry Competitor',
      'Company Size': '90 employees',
      'ESG Frameworks': 'N/A',
      'Reporting Obligations': 'N/A',
      'ESG Reporting Maturity': 'N/A',
      'Supply Chain Profile': 'N/A',
      'Scope 3 Exposure': 'N/A',
      'Sustainability Data Needs': 'None',
      'Digital Transformation Signals': 'Competitor domain detected',
      'ESG Governance': 'Internal Leadership',
      'Business Triggers': 'Competitor alert',
      'eValuater Solution Fit': 'Disqualified - Direct Competitor',
      'Company Opportunity Score': 30,
      'Company Opportunity': 'Competitor Organization - Do Not Contact',
      'Evidence URLs': 'https://competitor-esg.com',
      'Evidence Checked Date': dateStr,
      'Verification Status': 'Disqualified',
      'Human Review Status': 'Rejected',
      'Last Enriched At': nowStr,
      'Company Score Rationale': 'Disqualified: Domain listed on competitor suppression list.'
    }
  ];
}

function writePilotCompanyIntelligence(ss, companies) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.COMPANY_INTELLIGENCE);
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = companies.map(function(c) { return buildRowFromHeaders(headers, c); });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function getPilotResearchEvidence(config) {
  const nowStr = new Date().toLocaleString();
  return [
    {
      'Evidence ID': 'PILOT-EVD-001',
      'Campaign ID': config.CAMPAIGN_ID,
      'Entity Name': 'Apex Global Operations',
      'Entity Type': 'Company',
      'Search Query Used': 'Apex Global Operations sustainability emissions disclosure',
      'Grounding Evidence URL': 'https://example.com/apex-operations-audit',
      'Extracted Snippet / Signal': 'Documented corporate disclosure indicating multi-tier fleet logistics and Scope 3 reporting obligations.',
      'Verification Score': 'High',
      'Date Captured': nowStr
    },
    {
      'Evidence ID': 'PILOT-EVD-002',
      'Campaign ID': config.CAMPAIGN_ID,
      'Entity Name': 'Marcus Vance',
      'Entity Type': 'Persona',
      'Search Query Used': 'Marcus Vance Apex Operations keynote ' + config.DEPARTMENT,
      'Grounding Evidence URL': 'https://linkedin.com/in/marcus-vance-pilot',
      'Extracted Snippet / Signal': 'Discussed challenges balancing sprint cycles with audit requirements and manual verification overhead.',
      'Verification Score': 'High',
      'Date Captured': nowStr
    },
    {
      'Evidence ID': 'PILOT-EVD-003',
      'Campaign ID': config.CAMPAIGN_ID,
      'Entity Name': 'Novus Systems Corp',
      'Entity Type': 'Company',
      'Search Query Used': 'Novus Systems supplier compliance reporting 2026',
      'Grounding Evidence URL': 'https://example.com/novus-systems-report',
      'Extracted Snippet / Signal': 'Tier-1 industrial equipment supplier report detailing supply chain audit deadlines.',
      'Verification Score': 'High',
      'Date Captured': nowStr
    },
    {
      'Evidence ID': 'PILOT-EVD-004',
      'Campaign ID': config.CAMPAIGN_ID,
      'Entity Name': 'Elena Rostova',
      'Entity Type': 'Persona',
      'Search Query Used': 'Elena Rostova Novus Systems audit compliance drag',
      'Grounding Evidence URL': 'https://linkedin.com/in/elena-rostova-pilot',
      'Extracted Snippet / Signal': 'Published article highlighting vendor audit fatigue and delayed procurement reviews.',
      'Verification Score': 'High',
      'Date Captured': nowStr
    }
  ];
}

function writePilotResearchEvidence(ss, list) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.RESEARCH_EVIDENCE);
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = list.map(function(item) { return buildRowFromHeaders(headers, item); });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function getPilotPersonasData(config) {
  const primaryNeed = config.CUSTOM_BUSINESS_NEEDS.split('\n')[0];
  const nowStr = new Date().toLocaleString();
  const dateStr = new Date().toISOString().split('T')[0];

  return [
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-001',
      'Company Name': 'Apex Global Operations',
      'Industry': 'Freight & Logistics',
      'Persona ID': 'PILOT-PRSP-001',
      'Name': 'Marcus Vance',
      'Title': 'VP of Technology & Infrastructure',
      'Persona Type': 'Economic Buyer',
      'LinkedIn URL': 'https://linkedin.com/in/marcus-vance-pilot',
      'Signal': 'Spoke on challenges managing ' + primaryNeed + ' across growing cloud infrastructure.',
      'Signal Date': '2026-02-15',
      'Evidence URL': 'https://linkedin.com/in/marcus-vance-pilot',
      'Buying Signal': 'Frustration with manual evidence collection pulling engineers from core sprints.',
      'Verification Status': 'Verified',
      'Confidence': 'High',
      'Persona Fit Score': 95,
      'Personalized Hook': 'Saw your keynote discussing the bandwidth lost to ' + primaryNeed + '.',
      'Review Status': 'Ready for Outreach',
      'Last Verified At': nowStr,
      'Business Email': 'marcus.vance@growthtarget.com',
      'Email Domain': 'growthtarget.com',
      'Email Source URL': 'https://apexoperations.com/team',
      'Email Discovery Method': 'Pattern Match & SMTP Validated',
      'Email Verification Status': 'valid',
      'Email Verification Date': dateStr,
      'Email Verification Provider': 'ZeroBounce',
      'Email Identity Confidence': 'High',
      'Email Verification Notes': 'Deliverable, direct mailbox confirmed',
      'Draft Subject': "Apex's infrastructure sprint + " + primaryNeed,
      'Draft Email': "Hi Marcus,\n\n" +
        "Saw your recent perspective regarding the engineering bandwidth pulled away by " + primaryNeed + ". Scaling mission-critical infrastructure shouldn't require senior leaders to spend sprint cycles gathering manual verification evidence.\n\n" +
        "At " + config.SENDER_COMPANY + ", we help scaling organizations automate continuous controls and compliance workflows directly in their systems, cutting prep time by roughly 70% without slowing release cycles.\n\n" +
        "Open to a brief 7-minute intro this Thursday to see how we streamline this for teams like Apex?\n\n" +
        "Best regards,\n" +
        config.SENDER_NAME + "\n" +
        config.SENDER_TITLE + " | " + config.SENDER_COMPANY
    },
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-002',
      'Company Name': 'Novus Systems Corp',
      'Industry': 'Industrial Equipment',
      'Persona ID': 'PILOT-PRSP-002',
      'Name': 'Elena Rostova',
      'Title': 'Head of Engineering Operations',
      'Persona Type': 'Key Decision Maker',
      'LinkedIn URL': 'https://linkedin.com/in/elena-rostova-pilot',
      'Signal': 'Post on balancing sprint release velocity with Tier-1 vendor compliance obligations.',
      'Signal Date': '2026-02-10',
      'Evidence URL': 'https://linkedin.com/in/elena-rostova-pilot',
      'Buying Signal': 'Procurement and audit drag delaying release cycles.',
      'Verification Status': 'Verified',
      'Confidence': 'High',
      'Persona Fit Score': 91,
      'Personalized Hook': 'Really appreciated your perspective on eliminating compliance drag across vendor workflows.',
      'Review Status': 'Ready for Outreach',
      'Last Verified At': nowStr,
      'Business Email': 'elena.rostova@novussystems.com',
      'Email Domain': 'novussystems.com',
      'Email Source URL': 'https://novussystems.com/leadership',
      'Email Discovery Method': 'Direct Pattern Match',
      'Email Verification Status': 'catch-all',
      'Email Verification Date': dateStr,
      'Email Verification Provider': 'ZeroBounce',
      'Email Identity Confidence': 'Medium',
      'Email Verification Notes': 'Soft Warning: Server status is catch-all. Queued for review in Gmail Drafts.',
      'Draft Subject': "Sprint velocity vs. compliance drag at Novus",
      'Draft Email': "Hi Elena,\n\n" +
        "Really appreciated your recent article on balancing engineering velocity with compliance requirements—especially your point on how audit bottlenecks slow down core API releases.\n\n" +
        "We work specifically with engineering and operations leaders to automate verification guardrails directly into existing workflows. This lets your team ship updates at high velocity while staying completely audit-ready.\n\n" +
        "Would you be against exploring how we help technical teams eliminate compliance drag next Tuesday?\n\n" +
        "Best regards,\n" +
        config.SENDER_NAME + "\n" +
        config.SENDER_TITLE + " | " + config.SENDER_COMPANY
    },
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'PILOT-CO-003',
      'Company Name': 'EcoTransit Competitor Systems',
      'Industry': 'Direct Competitor',
      'Persona ID': 'PILOT-PRSP-003',
      'Name': 'Sarah Connor',
      'Title': 'VP of Product',
      'Persona Type': 'Disqualified / Competitor',
      'LinkedIn URL': 'https://linkedin.com/in/sarah-connor-pilot',
      'Signal': 'Competitor leadership match.',
      'Signal Date': '2026-02-01',
      'Evidence URL': 'https://competitor-esg.com/team',
      'Buying Signal': 'None - Competitor',
      'Verification Status': 'Disqualified',
      'Confidence': 'Low',
      'Persona Fit Score': 35,
      'Personalized Hook': 'Competitor domain detected.',
      'Review Status': 'Rejected',
      'Last Verified At': nowStr,
      'Business Email': 'sarah.connor@competitor-esg.com',
      'Email Domain': 'competitor-esg.com',
      'Email Source URL': 'https://competitor-esg.com',
      'Email Discovery Method': 'Scraped',
      'Email Verification Status': 'valid',
      'Email Verification Date': dateStr,
      'Email Verification Provider': 'ZeroBounce',
      'Email Identity Confidence': 'Low',
      'Email Verification Notes': 'Blocked: Domain on competitor suppression list',
      'Draft Subject': "Outreach discussion",
      'Draft Email': "Hi Sarah, checking in."
    }
  ];
}

function writePilotPersonaLeads(ss, personas) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.PERSONAS);
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = personas.map(function(p) { return buildRowFromHeaders(headers, p); });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Checks if Pilot Acme has a live Client Gmail Relay configured; if so, routes sandbox drafts through it.
 * Otherwise safely creates sandbox test drafts in local Gmail with sender options so testing works without a separate Workspace.
 */
function processSecurityApprovalAndDrafts(personas, enrichedCompanies, suppressionRules, config) {
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@'];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const nowStr = now.toLocaleString();

  const companyScoreMap = {};
  (enrichedCompanies || []).forEach(function(c) {
    companyScoreMap[c['Company ID']] = c['Company Opportunity Score'];
  });

  // Check optional Pilot Acme Gmail Relay configuration (sandbox only)
  let relayUrl = '';
  let relaySecret = '';
  if (typeof resolveClientScopedSecret_ === 'function') {
    const urlCheck = resolveClientScopedSecret_(
      ['GMAIL_RELAY_URL'],
      { campaignId: config.CAMPAIGN_ID, clientName: config.DEPARTMENT },
      { throwOnMissing: false }
    );
    const secCheck = resolveClientScopedSecret_(
      ['GMAIL_RELAY_SECRET'],
      { campaignId: config.CAMPAIGN_ID, clientName: config.DEPARTMENT },
      { throwOnMissing: false }
    );
    relayUrl = urlCheck.key || '';
    relaySecret = secCheck.key || '';
  }

  return personas.map(function(p) {
    const email = (p['Business Email'] || '').trim().toLowerCase();
    const domain = (p['Email Domain'] || '').trim().toLowerCase();
    const verificationStatus = String(p['Email Verification Status'] || '').trim().toLowerCase();
    const companyOppScore = companyScoreMap[p['Company ID']] || p['Persona Fit Score'];

    let suppressionCheck = 'Clear';
    let complianceStatus = 'Passed';
    let approvalStatus = 'Pending Review';
    let sendStatus = 'Queued';
    let sendError = '';
    let draftId = '';
    let draftNote = '';

    if (!email || !emailRegex.test(email)) {
      complianceStatus = 'Failed - Invalid Syntax';
      suppressionCheck = 'Blocked: Bad Email Syntax';
    } else if (rolePrefixes.some(function(pre) { return email.startsWith(pre); })) {
      complianceStatus = 'Failed - Role Based';
      suppressionCheck = 'Blocked: Role-Based Account';
    } else if (suppressionRules.emails.includes(email)) {
      complianceStatus = 'Failed - Suppressed';
      suppressionCheck = 'Blocked: Email in Lists & Settings';
    } else if (suppressionRules.domains.includes(domain)) {
      complianceStatus = 'Failed - Suppressed';
      suppressionCheck = 'Blocked: Domain in Lists & Settings';
    }

    if (complianceStatus === 'Passed' && (p['Persona Fit Score'] || 0) >= config.MIN_PERSONA_SCORE) {
      approvalStatus = (verificationStatus === 'catch-all' || verificationStatus === 'risky') ? 'Needs Review' : 'Approved (Auto-Pilot)';

      if (approvalStatus === 'Approved (Auto-Pilot)' || verificationStatus === 'catch-all') {
        try {
          if (relayUrl && relaySecret && /^https:\/\//i.test(relayUrl)) {
            const relayRes = UrlFetchApp.fetch(relayUrl, {
              method: 'post',
              contentType: 'application/json',
              payload: JSON.stringify({
                secret: relaySecret,
                action: 'createDraft',
                to: p['Business Email'],
                subject: p['Draft Subject'],
                body: p['Draft Email'],
                senderName: config.SENDER_NAME,
                senderEmail: config.SENDER_EMAIL || '',
                replyTo: config.REPLY_TO_EMAIL || config.SENDER_EMAIL || '',
                campaignId: config.CAMPAIGN_ID,
                clientName: config.DEPARTMENT
              }),
              muteHttpExceptions: true
            });
            const parsed = JSON.parse(relayRes.getContentText() || '{}');
            if (!parsed || parsed.success !== true) {
              throw new Error(parsed && parsed.error ? parsed.error : 'Relay error');
            }
            draftId = parsed.draftId || 'Client-Relay-Draft';
            draftNote = 'Draft created via ' + config.DEPARTMENT + ' Relay ID: ' + draftId;
          } else {
            const mailOptions = {};
            if (config.SENDER_NAME) mailOptions.name = config.SENDER_NAME;
            if (config.REPLY_TO_EMAIL) mailOptions.replyTo = config.REPLY_TO_EMAIL;

            const draft = GmailApp.createDraft(
              p['Business Email'],
              p['Draft Subject'],
              p['Draft Email'],
              mailOptions
            );
            draftId = draft.getId();
            draftNote = 'Sandbox draft created in local Gmail ID: ' + draftId + ' (Client: ' + config.DEPARTMENT + ')';
          }
          sendStatus = 'Draft Created in Gmail';
          Logger.log('SUCCESS: ' + draftNote + ' for ' + p['Name']);
        } catch (err) {
          sendStatus = 'Draft Failed';
          sendError = err.message;
        }
      }
    } else if (complianceStatus !== 'Passed') {
      approvalStatus = 'Rejected (Suppression Guard)';
      sendStatus = 'Suppressed - Not Sent';
      sendError = suppressionCheck;
    } else {
      approvalStatus = 'Needs Review';
      sendStatus = 'Queued';
    }

    return {
      'Campaign ID': p['Campaign ID'],
      'Campaign Name': p['Campaign Name'],
      'Company ID': p['Company ID'],
      'Company Name': p['Company Name'],
      'Persona ID': p['Persona ID'],
      'Name': p['Name'],
      'Persona Name': p['Name'],
      'Title': p['Title'],
      'Business Email': p['Business Email'],
      'Company Opportunity Score': companyOppScore,
      'Persona Fit Score': p['Persona Fit Score'],
      'Buying Signal': p['Buying Signal'],
      'Priority Score': p['Persona Fit Score'],
      'Priority Tier': p['Persona Fit Score'] >= 90 ? 'Tier 1' : 'Tier 2',
      'Priority Score Rationale': 'Evaluated for ' + config.DEPARTMENT + ' based on: ' + (config.CUSTOM_BUSINESS_NEEDS.slice(0, 50)),
      'Draft Subject': p['Draft Subject'],
      'Email Subject': p['Draft Subject'],
      'Draft Email': p['Draft Email'],
      'Email Body': p['Draft Email'],
      'Evidence URLs': p['Evidence URL'],
      'Review Status': approvalStatus.startsWith('Approved') ? 'Approved' : (approvalStatus === 'Needs Review' ? 'Needs Review' : 'Rejected'),
      'Reviewer': 'Pilot Automated Runner',
      'Reviewer Notes': draftNote ? (draftNote + (verificationStatus === 'catch-all' ? ' (Catch-All Warning)' : '')) : suppressionCheck,
      'Reviewed At': nowStr,
      'Human Approval Status': approvalStatus,
      'Approved By': approvalStatus.startsWith('Approved') ? 'Pilot Sandbox' : '',
      'Approved At': approvalStatus.startsWith('Approved') ? nowStr : '',
      'Send Status': sendStatus,
      'Sent At': '',
      'Dispatch Mode': 'Workspace Gmail (Drafts Only)',
      'Thread / Message ID': draftId,
      'Draft Created At': draftId ? nowStr : '',
      'Send Error': sendError,
      'Last Error': sendError,
      'Country': 'United States',
      'Industry': p['Industry'],
      'Company Opportunity': 'Operational / Compliance Automation',
      'Persona Type': p['Persona Type'],
      'LinkedIn URL': p['LinkedIn URL'],
      'Email Domain': p['Email Domain'],
      'Email Source URL': p['Email Source URL'],
      'Email Discovery Method': p['Email Discovery Method'],
      'Email Verification Status': p['Email Verification Status'],
      'Email Verification Date': p['Email Verification Date'],
      'Email Verification Provider': p['Email Verification Provider'],
      'Email Identity Confidence': p['Email Identity Confidence'],
      'Email Verification Notes': p['Email Verification Notes'],
      'Queue Status': 'Processed',
      'Compliance Status': complianceStatus,
      'Suppression Check': suppressionCheck,
      'Date Added': dateStr,
      'Personalized Hook': p['Personalized Hook']
    };
  });
}

function writePilotReviewQueue(ss, queueItems) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.REVIEW_QUEUE);
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = queueItems.map(function(item) { return buildRowFromHeaders(headers, item); });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function logPilotExecution(ss, processedQueue) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.LOGS);
  if (!sheet) return;

  const timestamp = new Date().toLocaleString();
  const logs = processedQueue.map(function(item) {
    return [
      timestamp,
      PILOT_CONFIG.CAMPAIGN_ID,
      'Client-Scoped Pilot Run (' + PILOT_CONFIG.DEPARTMENT + ')',
      item['Name'] + ' (' + item['Business Email'] + ')',
      item['Compliance Status'] === 'Passed' ? 'SUCCESS' : 'SUPPRESSED',
      item['Send Status'] + (item['Send Error'] ? ' - ' + item['Send Error'] : '')
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);
}

function updatePilotDashboard(ss, config, companiesCount, leadsCount, draftsCount, blockedCount) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
  if (!sheet) return;

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

  let verifierStatus = 'Key Not Configured';
  if (typeof resolveClientScopedSecret_ === 'function') {
    const verifKeyInfo = resolveClientScopedSecret_(
      ['ZEROBOUNCE_API_KEY', 'VERIFIER_API_KEY', 'EMAIL_VERIFICATION_API_KEY'],
      { campaignId: config.CAMPAIGN_ID, clientName: config.DEPARTMENT },
      { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: ss }
    );
    verifierStatus = verifKeyInfo.key
      ? 'Connected'
      : ('Missing: ' + verifKeyInfo.expectedProperty);
  }

  sheet.getRange('A1').setValue('🎯 AI PROSPECTING STUDIO – MASTER DASHBOARD');
  sheet.getRange('B1').setValue('Campaign');
  sheet.getRange('C1').setValue(config.CAMPAIGN_NAME);
  sheet.getRange('D1').setValue('Pilot Active');
  sheet.getRange('A1:D1').setBackground('#2d5945').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('E1:G1').clearContent();

  sheet.getRange('B2').setValue(config.CAMPAIGN_ID).setFontWeight('bold');

  const kpiHeaders = [['Metric KPI', 'Current Value', 'Operational Status', 'Last Updated']];
  sheet.getRange('A3:D3').setValues(kpiHeaders).setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');

  const kpiData = [
    ['Active Campaign Name', config.CAMPAIGN_NAME, 'Active', now],
    ['Campaign ID', config.CAMPAIGN_ID, 'Pilot Mode', now],
    ['Operating Department', config.DEPARTMENT, 'Client Owner', now],
    ['Company Discovery Limit', config.COMPANY_LIMIT + ' Companies', 'Enforced Limit', now],
    ['Companies Discovered & Qualified', companiesCount, companiesCount > 0 ? 'Discovery Complete' : 'Pending Discovery', now],
    ['Persona Leads Identified', leadsCount, leadsCount > 0 ? 'Mapping Complete' : 'Pending Mapping', now],
    ['Outreach Drafts Created', draftsCount, draftsCount > 0 ? 'Outreach Ready' : 'No Drafts Yet', now],
    ['Suppressed / Blocked Records', blockedCount, 'Suppression Active', now],
    ['Email Dispatch Channel', 'Workspace Gmail (Drafts Only)', draftsCount, now],
    ['Primary AI Model Configured', 'Gemini gemini-3.8-flash + Search Grounding', 'Connected', now],
    ['Email Verifier Configured', 'ZeroBounce', verifierStatus, now]
  ];
  if (!String(sheet.getRange('A14').getValue() || '').trim()) {
    sheet.getRange('A13:D13').copyTo(
      sheet.getRange('A14:D14'),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }
  sheet.getRange(4, 1, kpiData.length, 4).setValues(kpiData);

  sheet.getRange('A15:D15').clearContent();

  sheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);

  const rollupHeaders = [
    'Campaign ID',
    'Campaign Name',
    'Department',
    'Status',
    'Companies Found',
    'Outreach Drafts Created',
    'Last Run'
  ];
  sheet.getRange(17, 1, 1, 7)
    .setValues([rollupHeaders])
    .setFontWeight('bold')
    .setBackground('#dadce0')
    .setFontColor('#202124');

  const pilotRollupRow = [
    config.CAMPAIGN_ID,
    config.CAMPAIGN_NAME,
    config.DEPARTMENT,
    'Completed (Drafts Ready)',
    companiesCount,
    draftsCount,
    now
  ];

  // Update existing pilot rollup row if present, otherwise append without overwriting real campaigns
  let targetRollupRow = -1;
  const maxRow = sheet.getLastRow();
  if (maxRow >= 18) {
    const existingIds = sheet.getRange(18, 1, maxRow - 17, 1).getValues();
    for (let i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0] || '').trim() === config.CAMPAIGN_ID) {
        targetRollupRow = 18 + i;
        break;
      }
    }
  }
  if (targetRollupRow === -1) {
    targetRollupRow = Math.max(18, maxRow + 1);
  }
  sheet.getRange(targetRollupRow, 1, 1, 7).setValues([pilotRollupRow]);

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);

  SpreadsheetApp.flush();
}

function ensureSuppressionRulesInSettings(ss) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.LISTS_SETTINGS);
  if (!sheet) return;
  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.appendRow(['List name', 'Value', 'Active status']);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  let hasCompetitorRule = false;
  let hasOptOutRule = false;
  let hasInternalRule = false;
  let hasGrowthTarget = false;
  let hasPilotAcme = false;

  for (let i = 1; i < data.length; i++) {
    const listType = String(data[i][0] || '').trim().toLowerCase();
    const val = String(data[i][1] || '').trim().toLowerCase();
    if (val === 'competitor-esg.com') hasCompetitorRule = true;
    if (val === 'optout-lead@novusmfg.com') hasOptOutRule = true;
    if (val === 'internal-test@growthtarget.com') hasInternalRule = true;
    if ((listType === 'client - account' || listType === 'client / owner' || listType === 'client')) {
      if (val === 'growth target') hasGrowthTarget = true;
      if (val === 'pilot acme') hasPilotAcme = true;
    }
  }

  if (!hasGrowthTarget) {
    sheet.appendRow(['Client - Account', 'Growth Target', 'Active']);
  }
  if (!hasPilotAcme) {
    sheet.appendRow(['Client - Account', 'Pilot Acme', 'Active']);
  }
  if (!hasCompetitorRule) {
    sheet.appendRow(['Suppression - Competitor Domain', 'competitor-esg.com', 'Active']);
  }
  if (!hasOptOutRule) {
    sheet.appendRow(['Suppression - Opt-Out Email', 'optout-lead@novusmfg.com', 'Active']);
  }
  if (!hasInternalRule) {
    sheet.appendRow(['Suppression - Internal Exclusion', 'internal-test@growthtarget.com', 'Active']);
  }

  if (typeof sortListsAndSettingsTab_ === 'function') {
    sortListsAndSettingsTab_(ss);
  }
}

function loadSuppressionRulesFromSettings(ss) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.LISTS_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) return { domains: [], emails: [] };
  const data = sheet.getDataRange().getValues();
  const rules = { domains: [], emails: [] };

  for (let i = 1; i < data.length; i++) {
    const listName = String(data[i][0] || '').trim().toLowerCase();
    const value = String(data[i][1] || '').trim().toLowerCase();
    const status = String(data[i][2] || '').trim().toLowerCase();

    const isActive = (status === 'active' || status === 'true' || status === '');
    const isSuppression = listName.includes('suppression') ||
                          listName.includes('blacklist') ||
                          listName.includes('opt-out') ||
                          listName.includes('competitor');

    if (isActive && isSuppression && value) {
      if (value.includes('@')) {
        rules.emails.push(value);
      } else {
        rules.domains.push(value.replace(/^@/, ''));
      }
    }
  }
  return rules;
}

function buildRowFromHeaders(sheetHeaders, dataObject) {
  return sheetHeaders.map(function(header) {
    const cleanHeader = String(header).trim();
    return Object.prototype.hasOwnProperty.call(dataObject, cleanHeader) ? dataObject[cleanHeader] : '';
  });
}

/**
 * Safely collects all Campaign IDs associated with the Pilot sandbox
 * (PILOT-CMP-001, any PILOT-* ID, and any campaign owned by "Pilot Acme" such as PA-001).
 */
function collectPilotCampaignIds_(ss, primaryCampaignId) {
  const ids = {};
  const pilotClientLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();
  if (primaryCampaignId) {
    ids[String(primaryCampaignId).trim()] = true;
  }

  // 1. Scan Campaigns tab
  const campSheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGNS);
  if (campSheet && campSheet.getLastRow() > 1) {
    const cData = campSheet.getDataRange().getValues();
    for (let r = 1; r < cData.length; r++) {
      const cid = String(cData[r][0] || '').trim();
      const dept = String(cData[r][2] || '').trim().toLowerCase();
      if (cid && (cid.toUpperCase().indexOf('PILOT-') === 0 || dept === pilotClientLower)) {
        ids[cid] = true;
      }
    }
  }

  // 2. Scan Persona Mapper Settings tab
  const pSheet = (typeof findPersonaMapperSettingsSheet === 'function')
    ? findPersonaMapperSettingsSheet(ss)
    : ss.getSheetByName(PILOT_CONFIG.TABS.PERSONA_SETTINGS);
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let r = 1; r < pData.length; r++) {
      const cid = String(pData[r][0] || '').trim();
      const cname = String(pData[r][1] || '').trim().toLowerCase();
      if (cid && (cid.toUpperCase().indexOf('PILOT-') === 0 || cname.indexOf('pilot acme') >= 0)) {
        ids[cid] = true;
      }
    }
  }

  // 3. Scan Dashboard B2/B6 and Campaign Performance Rollup (rows 18+)
  const dashSheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
  if (dashSheet) {
    const b2 = String(dashSheet.getRange('B2').getValue() || '').trim();
    const b6 = String(dashSheet.getRange('B6').getValue() || '').trim().toLowerCase();
    if (b2 && (b2.toUpperCase().indexOf('PILOT-') === 0 || b6 === pilotClientLower)) {
      ids[b2] = true;
    }
    const maxRow = dashSheet.getLastRow();
    if (maxRow >= 18) {
      const rollup = dashSheet.getRange(18, 1, maxRow - 17, 3).getValues();
      for (let r = 0; r < rollup.length; r++) {
        const cid = String(rollup[r][0] || '').trim();
        const dept = String(rollup[r][2] || '').trim().toLowerCase();
        if (cid && (cid.toUpperCase().indexOf('PILOT-') === 0 || dept === pilotClientLower)) {
          ids[cid] = true;
        }
      }
    }
  }

  // 4. Scan Campaign Configuration
  const cfgSheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGN_CONFIG);
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    let cfgCid = '';
    let cfgDept = '';
    for (let r = 1; r < cfgData.length; r++) {
      const k = String(cfgData[r][0] || '').trim().toUpperCase();
      if (k === 'CAMPAIGN_ID') cfgCid = String(cfgData[r][1] || '').trim();
      if (k === 'DEPARTMENT') cfgDept = String(cfgData[r][1] || '').trim().toLowerCase();
    }
    if (cfgCid && (cfgCid.toUpperCase().indexOf('PILOT-') === 0 || cfgDept === pilotClientLower)) {
      ids[cfgCid] = true;
    }
  }

  // 5. Scan EmailEvents
  const eventsSheet = ss.getSheetByName('EmailEvents');
  if (eventsSheet && eventsSheet.getLastRow() > 1) {
    const eData = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 3).getValues();
    for (let r = 0; r < eData.length; r++) {
      const client = String(eData[r][1] || '').trim().toLowerCase();
      const cid = String(eData[r][2] || '').trim();
      if (cid && (cid.toUpperCase().indexOf('PILOT-') === 0 || client === pilotClientLower)) {
        ids[cid] = true;
      }
    }
  }

  return ids;
}

/**
 * Universal Pilot Purge: Scans every row across operational & log tabs (including API_Usage_Log,
 * Company Database Col J, Persona Mapper Settings, Execution Logs, and EmailEvents) and deletes pilot rows.
 * Removes stale sheet filters before deleting rows so filters never break or stay pinned.
 */
function purgeExistingPilotRows(ss, campaignIdOrMap) {
  const isFullWipe = (typeof campaignIdOrMap === 'object' && campaignIdOrMap !== null);
  const pilotIdsMap = isFullWipe ? campaignIdOrMap : {};
  const singleId = !isFullWipe ? String(campaignIdOrMap || PILOT_CONFIG.CAMPAIGN_ID).trim() : '';
  const pilotClientLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();

  const tabsToClean = [
    PILOT_CONFIG.TABS.CAMPAIGNS,
    PILOT_CONFIG.TABS.COMPANY_INTELLIGENCE,
    PILOT_CONFIG.TABS.PERSONAS,
    PILOT_CONFIG.TABS.REVIEW_QUEUE,
    PILOT_CONFIG.TABS.COMPANY_DATABASE,
    PILOT_CONFIG.TABS.RESEARCH_EVIDENCE,
    PILOT_CONFIG.TABS.PERSONA_SETTINGS,
    PILOT_CONFIG.TABS.LOGS,
    PILOT_CONFIG.TABS.API_USAGE_LOG,
    'EmailEvents'
  ];

  tabsToClean.forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet && tabName === PILOT_CONFIG.TABS.PERSONA_SETTINGS && typeof findPersonaMapperSettingsSheet === 'function') {
      sheet = findPersonaMapperSettingsSheet(ss);
    }
    if (!sheet) return;

    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    for (let r = data.length - 1; r >= 1; r--) {
      const isPilot = data[r].some(function(cell) {
        const val = String(cell === null || cell === undefined ? '' : cell).trim();
        if (!val) return false;
        const upperVal = val.toUpperCase();
        const lowerVal = val.toLowerCase();

        if (!isFullWipe) {
          return val === singleId || upperVal.indexOf('PILOT-') === 0;
        }

        if (pilotIdsMap[val] || upperVal.indexOf('PILOT-') === 0) return true;
        if (lowerVal === pilotClientLower) return true;
        if (lowerVal.indexOf('[client: ' + pilotClientLower + ']') >= 0) return true;
        if (lowerVal.indexOf('pilot run (' + pilotClientLower + ')') >= 0) return true;
        return false;
      });

      if (isPilot) {
        sheet.deleteRow(r + 1);
      }
    }
  });
}

/**
 * Option A Cleanup Helper: Removes "Pilot Acme" client entry and the 3 pilot suppression rules
 * from "Lists & Settings", re-sorts the tab, and refreshes sheet client dropdown validations.
 */
function cleanupPilotEntriesFromListsAndSettings_(ss) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.LISTS_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  const pilotClientLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();
  const data = sheet.getDataRange().getValues();

  for (let r = data.length - 1; r >= 1; r--) {
    const listName = String(data[r][0] || '').trim().toLowerCase();
    const val = String(data[r][1] || '').trim().toLowerCase();

    const isPilotClient = (listName === 'client - account' || listName === 'client / owner' || listName === 'client') &&
                          val === pilotClientLower;
    const isPilotSuppression =
      (listName === 'suppression - competitor domain' && val === 'competitor-esg.com') ||
      (listName === 'suppression - opt-out email' && val === 'optout-lead@novusmfg.com') ||
      (listName === 'suppression - internal exclusion' && val === 'internal-test@growthtarget.com');

    if (isPilotClient || isPilotSuppression) {
      sheet.deleteRow(r + 1);
    }
  }

  if (typeof sortListsAndSettingsTab_ === 'function') {
    sortListsAndSettingsTab_(ss);
  }
  if (typeof refreshClientDropdownValidations_ === 'function') {
    refreshClientDropdownValidations_(ss);
  }
}

/**
 * Option A Cleanup Helper: Resets Email Deliverability filters if pointing to Pilot Acme / PILOT-*
 * and refreshes the filter dropdown lists cleanly.
 */
function cleanupPilotEntriesFromDeliverability_(ss, pilotIdsMap) {
  const dash = ss.getSheetByName('Email Deliverability');
  if (!dash) return;

  const pilotLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();
  const b3 = String(dash.getRange('B3').getValue() || '').trim();
  const d3 = String(dash.getRange('D3').getValue() || '').trim();

  if (b3.toLowerCase() === pilotLower) {
    dash.getRange('B3').setValue('All Clients');
  }
  if (d3 && (pilotIdsMap[d3] || d3.toUpperCase().indexOf('PILOT-') === 0)) {
    dash.getRange('D3').setValue('All Campaigns');
  }

  if (typeof refreshDeliverabilityFilterDropdowns_ === 'function') {
    refreshDeliverabilityFilterDropdowns_(ss);
  }
}

/**
 * Option A Cleanup Helper: Removes all "Client: Pilot Acme" / CLIENT_PILOT_ACME_* rows
 * from Tab 11 ("Agent Settings") and clears any stale filter.
 */
function cleanupPilotEntriesFromAgentSettings_(ss) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.AGENT_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  const pilotCatLower = ('Client: ' + PILOT_CONFIG.DEPARTMENT).toLowerCase();
  const data = sheet.getDataRange().getValues();

  for (let r = data.length - 1; r >= 1; r--) {
    const category = String(data[r][0] || '').trim().toLowerCase();
    const settingKey = String(data[r][1] || '').trim().toUpperCase();

    if (category === pilotCatLower || settingKey.indexOf('CLIENT_PILOT_ACME_') === 0) {
      sheet.deleteRow(r + 1);
    }
  }
}

/**
 * Option A Cleanup Helper: Deletes only pilot rows from "Dashboard" Campaign Performance Rollup (Row 18+)
 * without wiping real production campaigns.
 */
function cleanupPilotEntriesFromDashboardRollup_(ss, pilotIdsMap) {
  const dashSheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
  if (!dashSheet) return;

  const lastRow = dashSheet.getLastRow();
  if (lastRow < 18) return;

  const pilotClientLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();
  const rollupData = dashSheet.getRange(18, 1, lastRow - 17, 4).getValues();

  for (let i = rollupData.length - 1; i >= 0; i--) {
    const cid = String(rollupData[i][0] || '').trim();
    const dept = String(rollupData[i][2] || '').trim().toLowerCase();
    const isPilot = Boolean(pilotIdsMap[cid]) ||
                    cid.toUpperCase().indexOf('PILOT-') === 0 ||
                    dept === pilotClientLower;

    if (isPilot) {
      dashSheet.deleteRow(18 + i);
    }
  }
}

/**
 * Resets "Campaign Configuration" Column B cleanly while strictly preserving all 38 locked parameter rows.
 * Used only when zero real campaigns remain after pilot cleanup.
 */
function resetCampaignConfigurationToStandby_(cfgSheet) {
  if (!cfgSheet || cfgSheet.getLastRow() <= 1) return;

  const standbyValues = {
    'CAMPAIGN_ID': '',
    'CAMPAIGN_NAME': '',
    'DEPARTMENT': 'Growth Target',
    'TARGET_COUNTRIES': '',
    'TARGET_INDUSTRIES': '',
    'COMPANY_SIZE_RANGE': '',
    'MAX_COMPANIES': 5,
    'RESEARCH_DEPTH': 'Deep Persona',
    'EVIDENCE_PERIOD_MONTHS': 12,
    'MIN_COMPANY_SCORE': 80,
    'MIN_PERSONA_SCORE': 85,
    'MIN_EMAIL_PRIORITY_SCORE': 80,
    'TARGET_TITLES': '',
    'CAMPAIGN_FOCUS': '',
    'SEARCH_KEYWORDS': '',
    'CUSTOM_BUSINESS_NEEDS': '',
    'AI_PROVIDER': 'GEMINI',
    'PRIMARY_AI_PROVIDER': 'GEMINI',
    'AI_MODEL': 'gemini-3.8-flash',
    'EMAIL_VERIFICATION_PROVIDER': 'ZeroBounce',
    'EMAIL_VERIFICATION_ENDPOINT': '',
    'EMAIL_SENDER_PROVIDER': 'Workspace Gmail (Drafts Only)',
    'DISPATCH_MODE': 'Workspace Gmail (Drafts Only)',
    'EMAIL_SENDER_ENDPOINT': '',
    'SENDER_NAME': '',
    'SENDER_TITLE': '',
    'SENDER_COMPANY': '',
    'INCLUDE_ESG_OBLIGATIONS': 'TRUE',
    'INCLUDE_ESG_MATURITY': 'TRUE',
    'INCLUDE_SUPPLY_CHAIN': 'TRUE',
    'INCLUDE_SCOPE3': 'TRUE',
    'INCLUDE_DATA_CHALLENGES': 'TRUE',
    'INCLUDE_DIGITAL': 'TRUE',
    'INCLUDE_ESG_GOVERNANCE': 'TRUE',
    'INCLUDE_TRIGGERS': 'TRUE',
    'INCLUDE_EVALUATER': 'TRUE',
    'HUMAN_APPROVAL_REQUIRED': 'TRUE',
    'AUTO_SEND_ENABLED': 'FALSE'
  };

  const integerKeys = ['MAX_COMPANIES', 'EVIDENCE_PERIOD_MONTHS', 'MIN_COMPANY_SCORE', 'MIN_PERSONA_SCORE', 'MIN_EMAIL_PRIORITY_SCORE'];
  const data = cfgSheet.getDataRange().getValues();

  for (let r = 1; r < data.length; r++) {
    const key = String(data[r][0] || '').trim().toUpperCase();
    if (Object.prototype.hasOwnProperty.call(standbyValues, key)) {
      const cell = cfgSheet.getRange(r + 1, 2);
      if (integerKeys.indexOf(key) >= 0) {
        cell.setNumberFormat('0');
      }
      cell.setValue(standbyValues[key]);
    }
  }
}

/**
 * OPTION A FULL WIPE CLEANUP:
 * Removes zero-trace of all Pilot test data (PILOT-CMP-001, PILOT-*, and Pilot Acme)
 * across all tabs including EmailEvents, preserves/restores active real campaigns,
 * and resets to Standby only if none exist.
 */
function cleanupPilotDataFromRealSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  // Step 1: Discover all pilot campaign IDs (PILOT-CMP-001, PILOT-*, and any Pilot Acme campaigns like PA-001)
  const pilotIdsMap = collectPilotCampaignIds_(ss, PILOT_CONFIG.CAMPAIGN_ID);

  // Step 2: Purge all pilot rows from operational, EmailEvents & log tabs and remove stale filters
  purgeExistingPilotRows(ss, pilotIdsMap);

  // Step 3: Clean stored DocumentProperties cursors for all purged pilot campaigns
  try {
    const docProps = PropertiesService.getDocumentProperties();
    Object.keys(pilotIdsMap).forEach(function(pid) {
      docProps.deleteProperty('PERSONA_MAPPER_CURSOR_' + pid);
    });
  } catch (e) {}

  // Step 4: Remove Pilot Acme & pilot suppression rules from Lists & Settings and refresh client dropdowns
  cleanupPilotEntriesFromListsAndSettings_(ss);

  // Step 4b: Reset Email Deliverability filter if pointing to Pilot Acme / PILOT-* and refresh dropdowns
  cleanupPilotEntriesFromDeliverability_(ss, pilotIdsMap);

  // Step 5: Remove Client: Pilot Acme block from Agent Settings
  cleanupPilotEntriesFromAgentSettings_(ss);

  // Step 6: Remove only Pilot rows from Dashboard -> Campaign Performance Rollup (preserve real campaigns)
  cleanupPilotEntriesFromDashboardRollup_(ss, pilotIdsMap);

  // Step 7: Remove any legacy stray GEMINI_MODEL/STATUS rows if Campaign Configuration exceeded 38 params
  const cfgSheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGN_CONFIG);
  let activeCfgId = '';
  let activeCfgDept = '';
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    for (let r = cfgData.length - 1; r >= 1; r--) {
      const param = String(cfgData[r][0] || '').trim().toUpperCase();
      if (cfgData.length > 39 && (param === 'GEMINI_MODEL' || param === 'STATUS')) {
        cfgSheet.deleteRow(r + 1);
      }
    }
    const refreshedCfg = cfgSheet.getDataRange().getValues();
    for (let r = 1; r < refreshedCfg.length; r++) {
      const k = String(refreshedCfg[r][0] || '').trim().toUpperCase();
      if (k === 'CAMPAIGN_ID') activeCfgId = String(refreshedCfg[r][1] || '').trim();
      if (k === 'DEPARTMENT') activeCfgDept = String(refreshedCfg[r][1] || '').trim();
    }
  }

  // Step 8: Collect remaining real Campaign IDs from Campaigns tab
  const realCampaignIds = [];
  const campSheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGNS);
  if (campSheet && campSheet.getLastRow() > 1) {
    const campRows = campSheet.getDataRange().getValues();
    for (let r = 1; r < campRows.length; r++) {
      const cid = String(campRows[r][0] || '').trim();
      const dept = String(campRows[r][2] || '').trim().toLowerCase();
      if (cid && !pilotIdsMap[cid] && cid.toUpperCase().indexOf('PILOT-') !== 0 && dept !== PILOT_CONFIG.DEPARTMENT.toLowerCase()) {
        realCampaignIds.push(cid);
      }
    }
  }

  // Step 9: Restore active real campaign cleanly (or reset to Standby if 0 real campaigns exist)
  const pilotLower = PILOT_CONFIG.DEPARTMENT.toLowerCase();
  const cfgPointsToValidReal = Boolean(
    activeCfgId &&
    !pilotIdsMap[activeCfgId] &&
    activeCfgId.toUpperCase().indexOf('PILOT-') !== 0 &&
    activeCfgDept.toLowerCase() !== pilotLower &&
    realCampaignIds.indexOf(activeCfgId) >= 0
  );

  const targetRealId = cfgPointsToValidReal
    ? activeCfgId
    : (realCampaignIds.length > 0 ? realCampaignIds[0] : '');

  if (targetRealId) {
    if (typeof selectAndActivateCampaign === 'function') {
      selectAndActivateCampaign(targetRealId);
    }
  } else {
    resetCampaignConfigurationToStandby_(cfgSheet);

    const dashSheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
    if (dashSheet) {
      const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

      dashSheet.getRange('A1').setValue('🎯 AI PROSPECTING STUDIO – MASTER DASHBOARD');
      dashSheet.getRange('B1').setValue('Campaign');
      dashSheet.getRange('C1').setValue('No Active Campaign');
      dashSheet.getRange('D1').setValue('Standby');
      dashSheet.getRange('A1:D1').setBackground('#2d5945').setFontColor('#ffffff').setFontWeight('bold');
      dashSheet.getRange('E1:G1').clearContent();

      dashSheet.getRange('B2').setValue('');

      dashSheet.getRange('A3:D3')
        .setValues([['Metric KPI', 'Current Value', 'Operational Status', 'Last Updated']])
        .setFontWeight('bold')
        .setBackground('#202124')
        .setFontColor('#ffffff');

      const standbyData = [
        ['Active Campaign Name', 'None (Standby)', 'Standby', now],
        ['Campaign ID', 'None', 'Standby', now],
        ['Operating Department', 'None', 'Standby', now],
        ['Company Discovery Limit', '0 Companies', 'Standby', now],
        ['Companies Discovered & Qualified', 0, 'Standby', now],
        ['Persona Leads Identified', 0, 'Standby', now],
        ['Outreach Drafts Created', 0, 'Standby', now],
        ['Suppressed / Blocked Records', 0, 'Standby', now],
        ['Email Dispatch Channel', 'Workspace Gmail (Drafts Only)', 0, now],
        ['Primary AI Model Configured', 'Gemini gemini-3.8-flash + Search Grounding', 'Connected', now],
        ['Email Verifier Configured', 'ZeroBounce', 'Standby', now]
      ];
      if (!String(dashSheet.getRange('A14').getValue() || '').trim()) {
        dashSheet.getRange('A13:D13').copyTo(
          dashSheet.getRange('A14:D14'),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
      }
      dashSheet.getRange(4, 1, standbyData.length, 4).setValues(standbyData);
      dashSheet.getRange('A15:D15').clearContent();

      dashSheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);
      const rollupHeaders = [
        'Campaign ID',
        'Campaign Name',
        'Department',
        'Status',
        'Companies Found',
        'Outreach Drafts Created',
        'Last Run'
      ];
      dashSheet.getRange(17, 1, 1, 7)
        .setValues([rollupHeaders])
        .setFontWeight('bold')
        .setBackground('#dadce0')
        .setFontColor('#202124');

      if (dashSheet.getLastRow() >= 18) {
        dashSheet.getRange(18, 1, dashSheet.getLastRow() - 17, 7).clearContent();
      }
    }

    if (typeof setupAgentSettingsTabDropdowns === 'function') {
      setupAgentSettingsTabDropdowns(ss, 'Growth Target');
    }
  }

  SpreadsheetApp.flush();
  ss.toast(
    'Pilot data wiped cleanly (including EmailEvents). If the Campaign Control Panel sidebar is open, close and reopen it to refresh dropdowns.',
    '🧹 Pilot Wipe Complete',
    8
  );
  Logger.log('Pilot test data wiped cleanly (Option A Full Wipe).');
}