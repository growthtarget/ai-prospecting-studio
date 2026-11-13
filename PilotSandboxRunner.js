/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AUTHENTIC PIPELINE PILOT RUNNER (FULL WORKSPACE GENERATION & SYNC)
 * ============================================================================
 * File: PilotSandboxRunner.gs / PilotSandboxRunner.js
 * 
 * CAPABILITIES:
 * 1. Generates & Synchronizes ALL 12 Tabs for the Pilot Run:
 *    - Cold-start safe: Automatically seeds canonical headers across all 12 tabs.
 *    - Updates "Campaign Configuration" with canonical parameters & company limit.
 *    - Generates 39-column row in "Persona Mapper Settings" for PILOT-CMP-001.
 *    - Restores and populates "Dashboard" with the complete 10-row KPI table.
 *    - Populates "Campaigns" (28 master cols), "Company Database" (Col J Campaign ID), 
 *      "Company Intelligence" (25 cols), "Persona Leads" (29 cols),
 *      "Research Evidence" (9 canonical cols for Company & Persona), and "Prospect Review Queue" (45 cols).
 * 2. Creates Safe Gmail Drafts for qualified leads in your personal Gmail.
 * 3. Enforces real-time suppression rules from "Lists & Settings".
 * 4. Safe Purge: Removes pilot records from ALL tabs including Execution Logs without erasing Dashboard structure.
 * ============================================================================
 */

const PILOT_CONFIG = {
  CAMPAIGN_ID: 'PILOT-CMP-001',
  CAMPAIGN_NAME: '[PILOT] Enterprise Compliance & Operations Readiness',
  DEPARTMENT: 'ESG Advisory',
  COMPANY_LIMIT: 5,
  RESEARCH_DEPTH: 'Deep Persona',
  MIN_COMPANY_SCORE: 80,
  MIN_PERSONA_SCORE: 85,
  SENDER_NAME: 'Alex Mercer',
  SENDER_TITLE: 'Solutions Consultant',
  SENDER_COMPANY: 'eValuater Solutions',
  CUSTOM_BUSINESS_NEEDS: 'Automated Scope 3 emissions reporting, supplier audit evidence collection, and CSRD compliance readiness',
  TABS: {
    DASHBOARD: 'Dashboard',
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
 * MAIN ENTRY POINT: Runs the full pilot, generates settings, and updates Dashboard.
 */
function runPilotOnRealSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  Logger.log('===============================================================');
  Logger.log('>>> STARTING FULL WORKSPACE PILOT: ' + PILOT_CONFIG.CAMPAIGN_ID + ' <<<');
  Logger.log('===============================================================');

  try {
    // Step 0: Ensure all 12 project tabs exist with canonical headers and high-contrast styling
    validateRequiredTabs(ss);

    // Step 1: Clean previous pilot test rows from all tabs
    purgeExistingPilotRows(ss, PILOT_CONFIG.CAMPAIGN_ID);
    Logger.log('Step 1: Purged previous pilot records.');

    // Step 2: Ensure and load suppression rules from "Lists & Settings"
    ensureSuppressionRulesInSettings(ss);
    const suppressionRules = loadSuppressionRulesFromSettings(ss);
    Logger.log('Step 2: Loaded suppression rules (' + suppressionRules.domains.length + ' domains, ' + suppressionRules.emails.length + ' emails).');

    // Step 3: Update "Campaign Configuration" tab with standardized Pilot parameters
    updateCampaignConfigurationTab(ss, PILOT_CONFIG);
    Logger.log('Step 3: "Campaign Configuration" updated with Company Limit: ' + PILOT_CONFIG.COMPANY_LIMIT);

    // Step 4: Generate 39-column settings row in "Persona Mapper Settings"
    generatePersonaMapperSettingsRow(ss, PILOT_CONFIG);
    Logger.log('Step 4: "Persona Mapper Settings" generated for ' + PILOT_CONFIG.CAMPAIGN_ID);

    // Step 5: Write Pilot Campaign Record to "Campaigns" (28 Master Columns)
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

    // Step 8: Write Grounding Evidence (both Company and Persona) to "Research Evidence" (9 Canonical Columns)
    const researchEvidence = getPilotResearchEvidence(PILOT_CONFIG);
    writePilotResearchEvidence(ss, researchEvidence);
    Logger.log('Step 8: Search grounding evidence logged in "Research Evidence" (Company & Persona citations).');

    // Step 9: Write Persona Leads & Outreach Copy to "Persona Leads" (29 Columns)
    const pilotPersonas = getPilotPersonasData(PILOT_CONFIG);
    writePilotPersonaLeads(ss, pilotPersonas);
    Logger.log('Step 9: Persona leads written to "Persona Leads".');

    // Step 10: Security Audit, Auto-Approval & Gmail Draft Dispatch
    const processedQueue = processSecurityApprovalAndDrafts(pilotPersonas, suppressionRules, PILOT_CONFIG);
    writePilotReviewQueue(ss, processedQueue);
    Logger.log('Step 10: Staged and processed leads in "Prospect Review Queue".');

    // Step 11: Log Execution in "Execution Logs"
    logPilotExecution(ss, processedQueue);

    // Step 12: Generate and Update Master Dashboard
    const draftsCreated = processedQueue.filter(p => p['Send Status'] === 'Draft Created in Gmail').length;
    const blockedCount = processedQueue.filter(p => p['Compliance Status'] !== 'Passed').length;
    updatePilotDashboard(ss, PILOT_CONFIG, enrichedCompanies.length, pilotPersonas.length, draftsCreated, blockedCount);
    Logger.log('Step 12: Master Dashboard updated with live KPIs.');

    // Completion notice
    ss.toast(
      'Pilot generated successfully! ' + draftsCreated + ' Gmail drafts created. Check Dashboard, Persona Mapper Settings, and Gmail.',
      '🚀 Pilot Complete',
      12
    );

    Logger.log('===============================================================');
    Logger.log('>>> PILOT RUN COMPLETED: ALL 12 TABS & SETTINGS SYNCHRONIZED <<<');
    Logger.log('===============================================================');

  } catch (err) {
    Logger.log('❌ Pilot Execution Error: ' + err.stack);
    ss.toast('Pilot error: ' + err.message, 'Pilot Error', 10);
  }
}

/**
 * Validates that all 12 project tabs exist, seeds canonical headers if empty,
 * and applies uniform high-contrast styling (#174ea6 on #e8f0fe).
 */
function validateRequiredTabs(ss) {
  const definitions = (typeof APP_CONFIG !== 'undefined' && Array.isArray(APP_CONFIG.sheetDefinitions))
    ? APP_CONFIG.sheetDefinitions
    : [];

  Object.values(PILOT_CONFIG.TABS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    // Seed canonical headers from APP_CONFIG if sheet is brand new or empty
    const def = definitions.find(d => d.name === name);
    if (sheet.getLastRow() === 0 && def && Array.isArray(def.headers) && def.headers.length > 0) {
      sheet.appendRow(def.headers);
    }

    // Apply high-contrast styling and freeze Row 1
    if (sheet.getLastColumn() > 0) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn())
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(1);
    }
  });
}

/**
 * Step 3: Updates "Campaign Configuration" with standardized canonical parameters
 */
function updateCampaignConfigurationTab(ss, config) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGN_CONFIG);
  if (!sheet) return;

  const updates = {
    'CAMPAIGN_ID': config.CAMPAIGN_ID,
    'CAMPAIGN_NAME': config.CAMPAIGN_NAME,
    'DEPARTMENT': config.DEPARTMENT,
    'PRIMARY_AI_PROVIDER': 'GEMINI',
    'AI_PROVIDER': 'GEMINI',
    'AI_MODEL': 'gemini-3.8-flash',
    'GEMINI_MODEL': 'gemini-3.8-flash',
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
    'SENDER_NAME': config.SENDER_NAME,
    'SENDER_TITLE': config.SENDER_TITLE,
    'SENDER_COMPANY': config.SENDER_COMPANY,
    'EMAIL_SENDER_PROVIDER': 'Workspace Gmail (Drafts Only)',
    'DISPATCH_MODE': 'Workspace Gmail (Drafts Only)',
    'EMAIL_VERIFICATION_PROVIDER': 'ZeroBounce',
    'EMAIL_VERIFICATION_ENDPOINT': 'https://api.zerobounce.net/v2/validate',
    'EMAIL_SENDER_ENDPOINT': '',
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
    'AUTO_SEND_ENABLED': 'FALSE',
    'STATUS': 'Pilot Active'
  };

  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.appendRow(['Parameter', 'Configured Value', 'Description']);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
    Object.keys(updates).forEach(key => {
      sheet.appendRow([key, updates[key], 'Pilot configured parameter']);
    });
  } else {
    const data = sheet.getDataRange().getValues();
    const updatedKeys = [];
    for (let r = 1; r < data.length; r++) {
      const key = String(data[r][0] || '').trim().toUpperCase();
      if (updates.hasOwnProperty(key)) {
        sheet.getRange(r + 1, 2).setValue(updates[key]);
        updatedKeys.push(key);
      }
    }
    Object.keys(updates).forEach(key => {
      if (!updatedKeys.includes(key)) {
        sheet.appendRow([key, updates[key], 'Pilot configured parameter']);
      }
    });
  }
}

/**
 * Step 4: Generates the 39-Column Row in "Persona Mapper Settings"
 */
function generatePersonaMapperSettingsRow(ss, config) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.PERSONA_SETTINGS);
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

  // Remove existing pilot settings row if present
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
    'Medium',
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

  sheet.appendRow(row);
}

/**
 * Step 5: Writes Pilot Campaign Record to "Campaigns" (28 Master Columns)
 */
function writePilotCampaign(ss, config) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.CAMPAIGNS);
  if (!sheet) return;

  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const nowStr = now.toLocaleString();

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
    senderCompany: config.SENDER_COMPANY
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

  // Guarantee 28 column headers if sheet is empty
  const campDef = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.sheetDefinitions)
    ? APP_CONFIG.sheetDefinitions.find(d => d.name === 'Campaigns')
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

  sheet.appendRow(master28Row);
}

/**
 * Step 6: Master Company Database records (Includes Column J: Campaign ID)
 */
function getPilotCompanyDatabaseRecords(config) {
  const dateStr = new Date().toISOString().split('T')[0];
  return [
    {
      'Company ID': 'CMP-001',
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
      'Company ID': 'CMP-002',
      'Campaign ID': config.CAMPAIGN_ID,
      'Company Name': 'Novus Systems Corp',
      'Domain': 'growthtarget.com',
      'Country': 'Canada',
      'Industry': 'Industrial Equipment',
      'Employee Size': '140 employees',
      'Discovery Method': 'Search Grounding (' + config.DEPARTMENT + ')',
      'Discovery Date': dateStr,
      'Status': 'Enriched'
    },
    {
      'Company ID': 'CMP-003',
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

/**
 * Writes pilot companies to "Company Database" and guarantees Column J is Campaign ID.
 */
function writePilotCompanyDatabase(ss, list) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.COMPANY_DATABASE);
  if (!sheet) return;

  // Guarantee Column J Header exists as "Campaign ID"
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let hasCampaignCol = currentHeaders.some(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '') === 'campaignid');

  if (!hasCampaignCol) {
    sheet.getRange(1, 10).setValue('Campaign ID')
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  }

  const updatedHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 10)).getValues()[0];
  const rows = list.map(item => buildRowFromHeaders(updatedHeaders, item));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, updatedHeaders.length).setValues(rows);
}

/**
 * Step 7: Enriched intelligence for "Company Intelligence" (All 25 Columns)
 */
function getPilotCompaniesData(config) {
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const nowStr = now.toLocaleString();

  return [
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'CMP-001',
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
      'Company Opportunity': 'Urgent bottleneck: manual evidence collection pulling engineering bandwidth (Matched for ' + config.DEPARTMENT + ')',
      'Evidence URLs': 'https://example.com/apex-operations-audit',
      'Evidence Checked Date': dateStr,
      'Verification Status': 'Verified',
      'Human Review Status': 'Accepted',
      'Last Enriched At': nowStr,
      'Company Score Rationale': 'High alignment with ' + config.DEPARTMENT + ' criteria.'
    },
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'CMP-002',
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
      'Company ID': 'CMP-003',
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
  const rows = companies.map(c => buildRowFromHeaders(headers, c));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Step 8: Evidence citations (Canonical 9 columns for both Company and Persona entities)
 */
function getPilotResearchEvidence(config) {
  const nowStr = new Date().toLocaleString();
  return [
    {
      'Evidence ID': 'EVD-001',
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
      'Evidence ID': 'EVD-002',
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
      'Evidence ID': 'EVD-003',
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
      'Evidence ID': 'EVD-004',
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
  const rows = list.map(item => buildRowFromHeaders(headers, item));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Step 9: Persona Leads
 */
function getPilotPersonasData(config) {
  const primaryNeed = config.CUSTOM_BUSINESS_NEEDS.split('\n')[0];
  const nowStr = new Date().toLocaleString();
  const dateStr = new Date().toISOString().split('T')[0];

  return [
    {
      'Campaign ID': config.CAMPAIGN_ID,
      'Campaign Name': config.CAMPAIGN_NAME,
      'Company ID': 'CMP-001',
      'Company Name': 'Apex Global Operations',
      'Industry': 'Freight & Logistics',
      'Persona ID': 'PRSP-001',
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
      'Company ID': 'CMP-002',
      'Company Name': 'Novus Systems Corp',
      'Industry': 'Industrial Equipment',
      'Persona ID': 'PRSP-002',
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
      'Business Email': 'elena.rostova@growthtarget.com',
      'Email Domain': 'growthtarget.com',
      'Email Source URL': 'https://novussystems.com/leadership',
      'Email Discovery Method': 'Direct Pattern Match',
      'Email Verification Status': 'valid',
      'Email Verification Date': dateStr,
      'Email Verification Provider': 'MillionVerifier',
      'Email Identity Confidence': 'High',
      'Email Verification Notes': 'Verified Active Mailbox',
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
      'Company ID': 'CMP-003',
      'Company Name': 'EcoTransit Competitor Systems',
      'Industry': 'Direct Competitor',
      'Persona ID': 'PRSP-003',
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
  const rows = personas.map(p => buildRowFromHeaders(headers, p));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Step 10: Security audit, approval & Gmail Drafts
 */
function processSecurityApprovalAndDrafts(personas, suppressionRules, config) {
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@'];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const nowStr = now.toLocaleString();

  return personas.map(p => {
    const email = (p['Business Email'] || '').trim().toLowerCase();
    const domain = (p['Email Domain'] || '').trim().toLowerCase();
    
    let suppressionCheck = 'Clear';
    let complianceStatus = 'Passed';
    let approvalStatus = 'Pending Review';
    let sendStatus = 'Queued';
    let sendError = '';
    let draftId = '';

    if (!email || !emailRegex.test(email)) {
      complianceStatus = 'Failed - Invalid Syntax';
      suppressionCheck = 'Blocked: Bad Email Syntax';
    } else if (rolePrefixes.some(pre => email.startsWith(pre))) {
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
      approvalStatus = 'Approved (Auto-Pilot)';
      try {
        const draft = GmailApp.createDraft(
          p['Business Email'],
          p['Draft Subject'],
          p['Draft Email'],
          { name: config.SENDER_NAME }
        );
        draftId = draft.getId();
        sendStatus = 'Draft Created in Gmail';
        Logger.log('SUCCESS: Gmail Draft created for ' + p['Name'] + ' | ID: ' + draftId);
      } catch (err) {
        sendStatus = 'Draft Failed';
        sendError = err.message;
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
      'Title': p['Title'],
      'Business Email': p['Business Email'],
      'Company Opportunity Score': p['Persona Fit Score'],
      'Persona Fit Score': p['Persona Fit Score'],
      'Buying Signal': p['Buying Signal'],
      'Priority Score': p['Persona Fit Score'],
      'Priority Tier': p['Persona Fit Score'] >= 90 ? 'Tier 1' : 'Tier 2',
      'Priority Score Rationale': 'Evaluated for ' + config.DEPARTMENT + ' based on: ' + (config.CUSTOM_BUSINESS_NEEDS.slice(0, 50)),
      'Draft Subject': p['Draft Subject'],
      'Draft Email': p['Draft Email'],
      'Evidence URLs': p['Evidence URL'],
      'Review Status': approvalStatus.startsWith('Approved') ? 'Approved' : (approvalStatus === 'Needs Review' ? 'Needs Review' : 'Rejected'),
      'Reviewer': 'Pilot Automated Runner',
      'Reviewer Notes': draftId ? 'Draft created in Gmail ID: ' + draftId : suppressionCheck,
      'Reviewed At': nowStr,
      'Human Approval Status': approvalStatus,
      'Approved By': approvalStatus.startsWith('Approved') ? 'Pilot Sandbox' : '',
      'Approved At': approvalStatus.startsWith('Approved') ? nowStr : '',
      'Send Status': sendStatus,
      'Sent At': draftId ? nowStr : '',
      'Send Error': sendError,
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
  const rows = queueItems.map(item => buildRowFromHeaders(headers, item));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function logPilotExecution(ss, processedQueue) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.LOGS);
  if (!sheet) return;

  const timestamp = new Date().toLocaleString();
  const logs = processedQueue.map(item => [
    timestamp,
    PILOT_CONFIG.CAMPAIGN_ID,
    'Multi-Department Pilot Run',
    item['Name'] + ' (' + item['Business Email'] + ')',
    item['Compliance Status'] === 'Passed' ? 'SUCCESS' : 'SUPPRESSED',
    item['Send Status'] + (item['Send Error'] ? ' - ' + item['Send Error'] : '')
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);
}

/**
 * Step 12: Restores and writes the authentic 10-row KPI structure and Rollup table.
 */
function updatePilotDashboard(ss, config, companiesCount, leadsCount, draftsCount, blockedCount) {
  const sheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
  if (!sheet) return;

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

  // Top Green Banner (Row 1) & Campaign ID (Row 2)
  sheet.getRange('A1').setValue('🎯 AI PROSPECTING STUDIO – MASTER DASHBOARD');
  sheet.getRange('B1').setValue('Campaign');
  sheet.getRange('C1').setValue(config.CAMPAIGN_NAME);
  sheet.getRange('D1').setValue('Pilot Active');
  sheet.getRange('A1:D1').setBackground('#2d5945').setFontColor('#ffffff').setFontWeight('bold');

  sheet.getRange('B2').setValue(config.CAMPAIGN_ID).setFontWeight('bold');

  // Header Row (Row 3)
  const kpiHeaders = [['Metric KPI', 'Current Value', 'Operational Status', 'Last Updated']];
  sheet.getRange('A3:D3').setValues(kpiHeaders).setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');

  // Complete 10 Metric Rows (Rows 4 to 13)
  const kpiData = [
    ['Active Campaign Name', config.CAMPAIGN_NAME, 'Active', now],
    ['Campaign ID', config.CAMPAIGN_ID, 'Pilot Mode', now],
    ['Operating Department', config.DEPARTMENT, 'Active Owner', now],
    ['Company Discovery Limit', config.COMPANY_LIMIT + ' Companies', 'Enforced Limit', now],
    ['Companies Discovered & Qualified', companiesCount, companiesCount > 0 ? 'Discovery Complete' : 'Pending Discovery', now],
    ['Persona Leads Identified', leadsCount, leadsCount > 0 ? 'Mapping Complete' : 'Pending Mapping', now],
    ['Gmail Outreach Drafts Created', draftsCount, draftsCount > 0 ? 'Ready in Gmail Drafts' : 'No Drafts Yet', now],
    ['Suppressed / Blocked Records', blockedCount, 'Suppression Active', now],
    ['Email Dispatch Channel', 'Workspace Gmail (Drafts Only)', draftsCount, now],
    ['Primary AI Model Configured', 'Gemini gemini-3.8-flash + Search Grounding', 'Connected', now]
  ];
  sheet.getRange(4, 1, kpiData.length, 4).setValues(kpiData);

  // Spacing (Rows 14 - 15)
  sheet.getRange('A14:D15').clearContent();

  // Campaign Performance Rollup Section Header (Row 16)
  sheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);

  // Rollup Table Headers (Row 17)
  const rollupHeaders = [
    'Campaign ID',
    'Campaign Name',
    'Department',
    'Status',
    'Companies Found',
    'Gmail Drafts',
    'Last Run'
  ];
  sheet.getRange(17, 1, 1, 7)
    .setValues([rollupHeaders])
    .setFontWeight('bold')
    .setBackground('#dadce0')
    .setFontColor('#202124');

  // Rollup Data for Pilot (Row 18)
  const pilotRollupRow = [
    config.CAMPAIGN_ID,
    config.CAMPAIGN_NAME,
    config.DEPARTMENT,
    'Completed (Drafts Ready)',
    companiesCount,
    draftsCount,
    now
  ];
  sheet.getRange(18, 1, 1, 7).setValues([pilotRollupRow]);

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
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');
  }

  const data = sheet.getDataRange().getValues();
  let hasCompetitorRule = false;
  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][1] || '').toLowerCase();
    if (val.includes('competitor-esg.com')) {
      hasCompetitorRule = true;
      break;
    }
  }

  if (!hasCompetitorRule) {
    sheet.appendRow(['Suppression - Competitor Domain', 'competitor-esg.com', 'Active']);
    sheet.appendRow(['Suppression - Opt-Out Email', 'optout-lead@novusmfg.com', 'Active']);
    sheet.appendRow(['Suppression - Internal Exclusion', 'internal-test@growthtarget.com', 'Active']);
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
  return sheetHeaders.map(header => {
    const cleanHeader = String(header).trim();
    return dataObject.hasOwnProperty(cleanHeader) ? dataObject[cleanHeader] : '';
  });
}

/**
 * Universal Pilot Purge: Scans every row across operational tabs (including Company Database Col J
 * and Execution Logs) and reliably deletes rows belonging to PILOT-CMP-001.
 */
function purgeExistingPilotRows(ss, campaignId) {
  const tabsToClean = [
    PILOT_CONFIG.TABS.CAMPAIGNS,
    PILOT_CONFIG.TABS.COMPANY_INTELLIGENCE,
    PILOT_CONFIG.TABS.PERSONAS,
    PILOT_CONFIG.TABS.REVIEW_QUEUE,
    PILOT_CONFIG.TABS.COMPANY_DATABASE,
    PILOT_CONFIG.TABS.RESEARCH_EVIDENCE,
    PILOT_CONFIG.TABS.PERSONA_SETTINGS,
    PILOT_CONFIG.TABS.LOGS
  ];

  tabsToClean.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    for (let r = data.length - 1; r >= 1; r--) {
      const rowValues = data[r].map(v => String(v || '').trim());
      // Checks all columns in the row so Column J in Company Database is cleanly caught
      const isPilot = rowValues.some(val => val === campaignId || val.startsWith('PILOT-'));
      if (isPilot) {
        sheet.deleteRow(r + 1);
      }
    }
  });
}

/**
 * Safe Cleanup: Resets Dashboard to Standby state without destroying Column A structure or labels.
 */
function cleanupPilotDataFromRealSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  purgeExistingPilotRows(ss, PILOT_CONFIG.CAMPAIGN_ID);
  
  const dashSheet = ss.getSheetByName(PILOT_CONFIG.TABS.DASHBOARD);
  if (dashSheet) {
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
    
    // Top Green Banner (Row 1) & Campaign ID (Row 2)
    dashSheet.getRange('A1').setValue('🎯 AI PROSPECTING STUDIO – MASTER DASHBOARD');
    dashSheet.getRange('B1').setValue('Campaign');
    dashSheet.getRange('C1').setValue('No Active Campaign');
    dashSheet.getRange('D1').setValue('Standby');
    dashSheet.getRange('A1:D1').setBackground('#2d5945').setFontColor('#ffffff').setFontWeight('bold');

    dashSheet.getRange('B2').setValue('');

    // Row 3 Header
    dashSheet.getRange('A3:D3')
      .setValues([['Metric KPI', 'Current Value', 'Operational Status', 'Last Updated']])
      .setFontWeight('bold')
      .setBackground('#202124')
      .setFontColor('#ffffff');

    // Reset Rows 4 to 13 values while permanently preserving Column A labels!
    const standbyData = [
      ['Active Campaign Name', 'None (Standby)', 'Standby', now],
      ['Campaign ID', 'None', 'Standby', now],
      ['Operating Department', 'None', 'Standby', now],
      ['Company Discovery Limit', '0 Companies', 'Standby', now],
      ['Companies Discovered & Qualified', 0, 'Standby', now],
      ['Persona Leads Identified', 0, 'Standby', now],
      ['Gmail Outreach Drafts Created', 0, 'Standby', now],
      ['Suppressed / Blocked Records', 0, 'Standby', now],
      ['Email Dispatch Channel', 'Workspace Gmail (Drafts Only)', 0, now],
      ['Primary AI Model Configured', 'Gemini gemini-3.8-flash + Search Grounding', 'Connected', now]
    ];
    dashSheet.getRange(4, 1, standbyData.length, 4).setValues(standbyData);

    // Rollup Headers & clear old data rows
    dashSheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);
    const rollupHeaders = [
      'Campaign ID',
      'Campaign Name',
      'Department',
      'Status',
      'Companies Found',
      'Gmail Drafts',
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
  
  ss.toast('Pilot test data purged. Dashboard reset to Standby with all labels intact.', 'Cleaned', 6);
  Logger.log('Pilot test data purged cleanly.');
}