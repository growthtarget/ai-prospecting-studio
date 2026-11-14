/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AUTOMATED SHEET & DROPDOWN SETUP ENGINE
 * ============================================================================
 * File: SheetSetup.gs / SheetSetup.js
 * 
 * PURPOSE:
 * 1. Pre-populates "Campaign Configuration" with all 38 standardized parameters.
 * 2. Injects native Google Sheets dropdown menus dynamically by parameter key.
 * 3. Enforces plain integer formatting (0) on limit/score columns to prevent date conversion.
 * 4. Enforces high-contrast readable headers (#174ea6 on #e8f0fe) across all 13 tabs.
 * 5. Pre-populates "Agent Settings" and dynamically refreshes Script Properties status.
 * 6. Initializes and formats "API_Usage_Log" (Tab 2).
 * ============================================================================
 */

function setupAllSheetsAndDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  // 1. Setup Dashboard if brand new (non-destructive if already populated)
  setupDashboardTabIfEmpty_(ss);

  // 2. Setup API Usage Log Tab (Tab 2)
  setupApiUsageLogTab(ss);

  // 3. Setup Campaigns Tab with 28 Canonical Headers, Number Formats & Dropdowns
  setupCampaignsTabDropdowns(ss);

  // 4. Setup Company Intelligence Tab with 25 Canonical Headers & Dropdowns
  setupCompanyIntelligenceTabDropdowns(ss);

  // 5. Setup Persona Leads Tab with 29 Canonical Headers & Dropdowns
  setupPersonaLeadsTabDropdowns(ss);

  // 6. Ensure Persona Mapper Settings (39 Canonical Headers) & Plain Integer Formats
  setupHeaderOnlyTab_(ss, 'Persona Mapper Settings', [
    'Campaign ID','Campaign Name','Persona Mapper Enabled','AI Provider','AI Model','Persona Titles',
    'Search Keywords','Campaign Focus','Evidence Period Months','Research Depth','Include ESG Reporting Obligations',
    'Include ESG Reporting Maturity','Include Supply Chain Complexity','Include Scope 3 Exposure','Include Sustainability Data Challenges',
    'Include Digital Transformation','Include ESG Governance','Include Business Triggers','Include eValuater Solution Fit',
    'Weight ESG Reporting Relevance','Weight Supply Chain Complexity','Weight Sustainability Data Needs','Weight Digital Transformation Relevance',
    'Weight Business Triggers','Weight eValuater Solution Fit','Minimum Company Opportunity Score','Minimum Persona Confidence',
    'Minimum Email Priority Score','Human Approval Required','Auto Send Enabled','Email Priority Weight Company Score',
    'Email Priority Weight Persona Fit','Email Priority Weight Buying Signal','Notes','Max Companies Per Run',
    'Email Verification Provider','Email Verification Endpoint','Email Sender Provider','Email Sender Endpoint'
  ]);
  const pSettingsSheet = ss.getSheetByName('Persona Mapper Settings');
  if (pSettingsSheet) {
    pSettingsSheet.getRange('I2:I500').setNumberFormat('0');
    pSettingsSheet.getRange('Z2:AB500').setNumberFormat('0');
    pSettingsSheet.getRange('AI2:AI500').setNumberFormat('0');
  }

  // 7. Setup Prospect Review Queue Tab with 45 Canonical Headers & Dropdowns
  setupReviewQueueTabDropdowns(ss);

  // 8. Setup Campaign Configuration with Locked 38 Parameters & Dynamic Dropdowns
  setupCampaignConfigurationTab(ss);

  // 9. Ensure Company Database (10 Canonical Headers, Col J = Campaign ID)
  setupHeaderOnlyTab_(ss, 'Company Database', [
    'Company ID', 'Company Name', 'Domain', 'Country', 'Industry',
    'Employee Size', 'Discovery Method', 'Discovery Date', 'Status', 'Campaign ID'
  ]);

  // 10. Ensure Research Evidence (9 Canonical Headers)
  setupHeaderOnlyTab_(ss, 'Research Evidence', [
    'Evidence ID', 'Campaign ID', 'Entity Name', 'Entity Type',
    'Search Query Used', 'Grounding Evidence URL', 'Extracted Snippet / Signal',
    'Verification Score', 'Date Captured'
  ]);

  // 11. Setup Agent Settings Tab with Structured Rows & Dynamic Key Status
  setupAgentSettingsTabDropdowns(ss);

  // 12. Ensure Execution Logs (6 Canonical Headers)
  setupHeaderOnlyTab_(ss, 'Execution Logs', [
    'Timestamp', 'Campaign ID', 'Action', 'Target', 'Status', 'Details'
  ]);

  // 13. Setup Lists & Settings Tab with Dropdowns
  setupListsAndSettingsTabDropdowns(ss);

  SpreadsheetApp.flush();
  ss.toast('All 13 tabs formatted and dynamic dropdown menus applied successfully!', '✅ Setup Complete', 10);
  Logger.log('>>> ALL 13 TABS & DROPDOWNS CONFIGURED SUCCESSFULLY <<<');
}

function setupDashboardTabIfEmpty_(ss) {
  const sheet = getOrCreateSheet(ss, 'Dashboard');
  if (sheet.getLastRow() > 0) return;

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
  sheet.getRange('A1').setValue('🎯 AI PROSPECTING STUDIO – MASTER DASHBOARD');
  sheet.getRange('B1').setValue('Campaign');
  sheet.getRange('C1').setValue('No Active Campaign');
  sheet.getRange('D1').setValue('Standby');
  sheet.getRange('A1:D1').setBackground('#2d5945').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('B2').setValue('');
  sheet.getRange('A3:D3').setValues([['Metric KPI', 'Current Value', 'Operational Status', 'Last Updated']])
    .setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');

  const kpis = [
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
  sheet.getRange(4, 1, kpis.length, 4).setValues(kpis);
  sheet.getRange('A16').setValue('Campaign Performance Rollup').setFontWeight('bold').setFontSize(12);
  sheet.getRange(17, 1, 1, 7)
    .setValues([['Campaign ID', 'Campaign Name', 'Department', 'Status', 'Companies Found', 'Gmail Drafts', 'Last Run']])
    .setFontWeight('bold').setBackground('#dadce0').setFontColor('#202124');
}

function setupHeaderOnlyTab_(ss, name, headers) {
  const sheet = getOrCreateSheet(ss, name);
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow(sheet);
}

/**
 * Formats "Campaign Configuration" with 38 Locked Parameters and Dynamic Dropdowns
 * STRICT: Zero Schema Mutations (Locked 38 Parameters)
 */
function setupCampaignConfigurationTab(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaign Configuration');
  sheet.clear();

  const headers = ['Parameter', 'Configured Value', 'Description'];
  sheet.appendRow(headers);
  styleHeaderRow(sheet);

  const configRows = [
    ['CAMPAIGN_ID', 'CMP-2026-001', 'Active Campaign Identifier'],
    ['CAMPAIGN_NAME', 'Enterprise Scope 3 & ESG Supply Chain Readiness', 'Active Campaign Name'],
    ['DEPARTMENT', 'ESG Advisory', 'Operating Department / Business Unit'],
    ['TARGET_COUNTRIES', 'United States & Canada', 'Target Geographies / Regions'],
    ['TARGET_INDUSTRIES', 'Freight & Logistics, Industrial Equipment', 'Target Industry Verticals'],
    ['COMPANY_SIZE_RANGE', '50 - 250 employees', 'Target Employee Size Range'],
    ['MAX_COMPANIES', '5', 'Maximum target companies to discover per run'],
    ['RESEARCH_DEPTH', 'Deep Persona', 'Depth of company research and profiling'],
    ['EVIDENCE_PERIOD_MONTHS', '12', 'Historical lookback window in months'],
    ['MIN_COMPANY_SCORE', '80', 'Minimum company qualification threshold (0-100)'],
    ['MIN_PERSONA_SCORE', '85', 'Minimum persona fit confidence threshold (0-100)'],
    ['MIN_EMAIL_PRIORITY_SCORE', '80', 'Minimum email priority score for sending (0-100)'],
    ['TARGET_TITLES', 'VP Procurement; Head of Supply Chain; Director of Sustainability', 'Semicolon-delimited target job titles'],
    ['CAMPAIGN_FOCUS', 'Scope 3 vendor emissions transparency and CSRD readiness', 'Primary value proposition hook'],
    ['SEARCH_KEYWORDS', 'Scope 3, CSRD, GHG Protocol, supplier risk, compliance audits', 'Search discovery query keywords'],
    ['CUSTOM_BUSINESS_NEEDS', '- Managing Scope 3 vendor emissions and Tier-1 ESG supply chain transparency\n- High vendor questionnaire drop-off rate during procurement audits\n- Eliminating manual spreadsheet follow-ups with suppliers', 'Manual department business needs & buying triggers'],
    ['AI_PROVIDER', 'GEMINI', 'Selected AI engine (GEMINI or OPENAI)'],
    ['PRIMARY_AI_PROVIDER', 'GEMINI', 'Selected AI engine (legacy alias)'],
    ['AI_MODEL', 'gemini-3.8-flash', 'AI Model version name'],
    ['EMAIL_VERIFICATION_PROVIDER', 'ZeroBounce', 'Email verification vendor'],
    ['EMAIL_VERIFICATION_ENDPOINT', 'https://api.zerobounce.net/v2/validate', 'Email verification REST endpoint URL'],
    ['EMAIL_SENDER_PROVIDER', 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform'],
    ['DISPATCH_MODE', 'Workspace Gmail (Drafts Only)', 'Outreach dispatch platform mode'],
    ['EMAIL_SENDER_ENDPOINT', '', 'Sender campaign UUID or webhook URL'],
    ['SENDER_NAME', 'Alex Mercer', 'Display name for outreach email signature'],
    ['SENDER_TITLE', 'Solutions Consultant', 'Sender job title for email signature'],
    ['SENDER_COMPANY', 'eValuater Solutions', 'Sender company name for email signature'],
    ['INCLUDE_ESG_OBLIGATIONS', 'TRUE', 'Evaluate CSRD/GRI obligations'],
    ['INCLUDE_ESG_MATURITY', 'TRUE', 'Evaluate ESG reporting maturity'],
    ['INCLUDE_SUPPLY_CHAIN', 'TRUE', 'Evaluate supply chain complexity'],
    ['INCLUDE_SCOPE3', 'TRUE', 'Evaluate Scope 3 emissions exposure'],
    ['INCLUDE_DATA_CHALLENGES', 'TRUE', 'Evaluate sustainability data challenges'],
    ['INCLUDE_DIGITAL', 'TRUE', 'Evaluate digital transformation & AI adoption'],
    ['INCLUDE_ESG_GOVERNANCE', 'TRUE', 'Evaluate ESG board oversight'],
    ['INCLUDE_TRIGGERS', 'TRUE', 'Evaluate commercial buying triggers'],
    ['INCLUDE_EVALUATER', 'TRUE', 'Evaluate eValuater solution fit'],
    ['HUMAN_APPROVAL_REQUIRED', 'TRUE', 'Requires human review in queue before sending'],
    ['AUTO_SEND_ENABLED', 'FALSE', 'Auto-send emails immediately upon approval']
  ];

  sheet.getRange(2, 1, configRows.length, 3).setValues(configRows);

  const strictRules = {
    'DEPARTMENT': ['Procurement', 'ESG Advisory', 'Enterprise Sales', 'Cybersecurity & IT', 'Operations & Compliance', 'Enterprise SaaS / IT', 'FinTech & Banking', 'HR & Talent Ops', 'Marketing', 'Executive Leadership'],
    'RESEARCH_DEPTH': ['Standard', 'Deep Persona', 'Comprehensive / Forensic'],
    'EVIDENCE_PERIOD_MONTHS': ['6', '12', '24'],
    'MIN_COMPANY_SCORE': ['70', '75', '80', '85', '90'],
    'MIN_PERSONA_SCORE': ['75', '80', '85', '90', '95'],
    'AI_PROVIDER': ['GEMINI', 'OPENAI'],
    'PRIMARY_AI_PROVIDER': ['GEMINI', 'OPENAI'],
    'EMAIL_VERIFICATION_PROVIDER': ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None'],
    'EMAIL_SENDER_PROVIDER': ['Workspace Gmail (Drafts Only)', 'Workspace Gmail (Direct Send)', 'Google Workspace (Drafts Only)', 'Google Workspace (Direct Send)', 'Smartlead', 'Instantly', 'Custom Webhook'],
    'DISPATCH_MODE': ['Workspace Gmail (Drafts Only)', 'Workspace Gmail (Direct Send)', 'Google Workspace (Drafts Only)', 'Google Workspace (Direct Send)', 'Smartlead', 'Instantly', 'Custom Webhook'],
    'INCLUDE_ESG_OBLIGATIONS': ['TRUE', 'FALSE'],
    'INCLUDE_ESG_MATURITY': ['TRUE', 'FALSE'],
    'INCLUDE_SUPPLY_CHAIN': ['TRUE', 'FALSE'],
    'INCLUDE_SCOPE3': ['TRUE', 'FALSE'],
    'INCLUDE_DATA_CHALLENGES': ['TRUE', 'FALSE'],
    'INCLUDE_DIGITAL': ['TRUE', 'FALSE'],
    'INCLUDE_ESG_GOVERNANCE': ['TRUE', 'FALSE'],
    'INCLUDE_TRIGGERS': ['TRUE', 'FALSE'],
    'INCLUDE_EVALUATER': ['TRUE', 'FALSE'],
    'HUMAN_APPROVAL_REQUIRED': ['TRUE', 'FALSE'],
    'AUTO_SEND_ENABLED': ['TRUE', 'FALSE']
  };

  // Flexible preset dropdowns matching ControlPanel.html free-text / custom number inputs
  const flexibleRules = {
    'TARGET_COUNTRIES': ['United States & Canada', 'United States Only', 'Canada Only', 'Sweden, Norway, Denmark', 'United Kingdom & Europe', 'Global'],
    'TARGET_INDUSTRIES': ['Freight & Logistics, Industrial Equipment', 'Energy, CleanTech & Utilities', 'Manufacturing, Industrial, BioTech', 'Financial Services / FinTech SaaS', 'Retail & Consumer Goods'],
    'COMPANY_SIZE_RANGE': ['10 - 50 employees', '50 - 250 employees', '250 - 1000 employees', '1000+ Enterprise'],
    'MAX_COMPANIES': ['3', '5', '10', '25', '50'],
    'MIN_EMAIL_PRIORITY_SCORE': ['70', '75', '80', '85', '90'],
    'AI_MODEL': ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro', 'gpt-4o', 'gpt-4o-mini', 'o1-mini']
  };

  const integerKeys = ['MAX_COMPANIES', 'EVIDENCE_PERIOD_MONTHS', 'MIN_COMPANY_SCORE', 'MIN_PERSONA_SCORE', 'MIN_EMAIL_PRIORITY_SCORE'];

  const lastRow = sheet.getLastRow();
  const paramKeys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let r = 0; r < paramKeys.length; r++) {
    const key = String(paramKeys[r][0] || '').trim().toUpperCase();
    const cellRange = sheet.getRange(r + 2, 2);
    if (integerKeys.indexOf(key) >= 0) {
      cellRange.setNumberFormat('0');
    }
    if (strictRules[key]) {
      applyDropdownToRange(cellRange, strictRules[key], false);
    } else if (flexibleRules[key]) {
      applyDropdownToRange(cellRange, flexibleRules[key], true);
    }
  }

  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(3, 400);
}

/**
 * Injects Dropdowns & Plain Integer Number Formats into "Campaigns" Tab (Locked 28 Columns)
 */
function setupCampaignsTabDropdowns(ss) {
  const headers = [
    'Campaign ID', 'Campaign Name', 'Department', 'Status', 'Target Countries',
    'Target Industries', 'Company Size', 'Discovery Limit', 'Research Depth',
    'Evidence Period Months', 'Min Company Score', 'Min Persona Score',
    'Min Email Priority Score', 'Target Titles', 'Campaign Focus', 'Search Keywords',
    'Custom Business Needs', 'AI Provider', 'AI Model', 'Email Verifier Provider',
    'Email Verifier Endpoint', 'Email Sender Provider', 'Email Sender Endpoint',
    'Sender Name', 'Sender Title', 'Sender Company', 'Campaign Config JSON', 'Last Modified'
  ];
  const sheet = getOrCreateSheet(ss, 'Campaigns');
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow(sheet);

  // Enforce plain integer format so scores (80, 85) and limits never render as 1900 dates
  sheet.getRange('H2:H500').setNumberFormat('0');
  sheet.getRange('J2:M500').setNumberFormat('0');

  applyDropdownToRange(sheet.getRange('D2:D500'), ['Draft', 'Active', 'Paused', 'Completed', 'Pilot Active'], false);
  applyDropdownToRange(sheet.getRange('H2:H500'), ['3', '5', '10', '25', '50'], true);
  applyDropdownToRange(sheet.getRange('I2:I500'), ['Standard', 'Deep Persona', 'Comprehensive / Forensic'], false);
  applyDropdownToRange(sheet.getRange('R2:R500'), ['GEMINI', 'OPENAI'], false);
  applyDropdownToRange(sheet.getRange('T2:T500'), ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None'], false);
  applyDropdownToRange(sheet.getRange('V2:V500'), ['Workspace Gmail (Drafts Only)', 'Workspace Gmail (Direct Send)', 'Google Workspace (Drafts Only)', 'Google Workspace (Direct Send)', 'Smartlead', 'Instantly', 'Custom Webhook'], false);
}

/**
 * Injects Dropdowns into "Prospect Review Queue" Tab (45 Canonical Columns)
 */
function setupReviewQueueTabDropdowns(ss) {
  const headers = [
    'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Persona ID',
    'Name', 'Title', 'Business Email', 'Company Opportunity Score', 'Persona Fit Score',
    'Buying Signal', 'Priority Score', 'Priority Tier', 'Priority Score Rationale',
    'Draft Subject', 'Draft Email', 'Evidence URLs', 'Review Status', 'Reviewer',
    'Reviewer Notes', 'Reviewed At', 'Human Approval Status', 'Approved By',
    'Approved At', 'Send Status', 'Sent At', 'Send Error', 'Country', 'Industry',
    'Company Opportunity', 'Persona Type', 'LinkedIn URL', 'Email Domain',
    'Email Source URL', 'Email Discovery Method', 'Email Verification Status',
    'Email Verification Date', 'Email Verification Provider', 'Email Identity Confidence',
    'Email Verification Notes', 'Queue Status', 'Compliance Status', 'Suppression Check',
    'Date Added', 'Personalized Hook'
  ];
  const sheet = getOrCreateSheet(ss, 'Prospect Review Queue');
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('M2:M1000'), ['Tier 1', 'Tier 2', 'Tier 3'], false);
  applyDropdownToRange(sheet.getRange('R2:R1000'), ['Needs Review', 'Approved', 'Rejected'], false);
  applyDropdownToRange(sheet.getRange('V2:V1000'), ['Approved', 'Approved (Auto-Pilot)', 'Needs Review', 'Pending Review', 'Rejected (Suppression Guard)', 'Rejected'], false);
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), ['Queued', 'Draft Created in Gmail', 'Draft Created', 'Sent', 'Pushed to Sequencer', 'Suppressed - Not Sent', 'Draft Failed'], false);
  applyDropdownToRange(sheet.getRange('AP2:AP1000'), ['Passed', 'Failed - Suppressed', 'Failed - Role Based', 'Failed - Invalid Syntax'], false);
  applyDropdownToRange(sheet.getRange('AQ2:AQ1000'), ['Clear', 'Blocked: Email in Lists & Settings', 'Blocked: Domain in Lists & Settings', 'Blocked: Role-Based Account', 'Blocked: Bad Email Syntax'], false);
}

/**
 * Injects Dropdowns into "Company Intelligence" Tab (25 Canonical Columns)
 */
function setupCompanyIntelligenceTabDropdowns(ss) {
  const headers = [
    'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Country',
    'Industry', 'Company Size', 'ESG Frameworks', 'Reporting Obligations',
    'ESG Reporting Maturity', 'Supply Chain Profile', 'Scope 3 Exposure',
    'Sustainability Data Needs', 'Digital Transformation Signals', 'ESG Governance',
    'Business Triggers', 'eValuater Solution Fit', 'Company Opportunity Score',
    'Company Opportunity', 'Evidence URLs', 'Evidence Checked Date',
    'Verification Status', 'Human Review Status', 'Last Enriched At', 'Company Score Rationale'
  ];
  const sheet = getOrCreateSheet(ss, 'Company Intelligence');
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('J2:J1000'), ['Early Stage', 'Developing / Spreadsheet-reliant', 'Developing', 'Mature', 'Advanced', 'N/A'], false);
  applyDropdownToRange(sheet.getRange('Q2:Q1000'), ['High', 'Medium', 'Low', 'Disqualified - Direct Competitor'], false);
  applyDropdownToRange(sheet.getRange('V2:V1000'), ['Verified', 'Unverified', 'Disqualified', 'Needs review'], false);
  applyDropdownToRange(sheet.getRange('W2:W1000'), ['Accepted', 'Pending', 'Rejected'], false);
}

/**
 * Injects Dropdowns into "Persona Leads" Tab (29 Canonical Columns)
 */
function setupPersonaLeadsTabDropdowns(ss) {
  const headers = [
    'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Industry',
    'Persona ID', 'Name', 'Title', 'Persona Type', 'LinkedIn URL', 'Signal',
    'Signal Date', 'Evidence URL', 'Buying Signal', 'Verification Status',
    'Confidence', 'Persona Fit Score', 'Personalized Hook', 'Review Status',
    'Last Verified At', 'Business Email', 'Email Domain', 'Email Source URL',
    'Email Discovery Method', 'Email Verification Status', 'Email Verification Date',
    'Email Verification Provider', 'Email Identity Confidence', 'Email Verification Notes'
  ];
  const sheet = getOrCreateSheet(ss, 'Persona Leads');
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('I2:I1000'), ['Economic Buyer', 'Technical Evaluator', 'Key Decision Maker', 'Influencer', 'Disqualified / Competitor'], false);
  applyDropdownToRange(sheet.getRange('O2:O1000'), ['Verified', 'Unverified', 'Pending', 'Disqualified'], false);
  applyDropdownToRange(sheet.getRange('S2:S1000'), ['Ready for Outreach', 'Needs Review', 'Approved', 'Rejected'], false);
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), ['valid', 'risky', 'invalid', 'catch-all', 'unverified'], false);
}

/**
 * Injects Dropdowns into "Lists & Settings" Tab
 */
function setupListsAndSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Lists & Settings');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['List name', 'Value', 'Active status']);
  }
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('A2:A500'), [
    'Suppression - Domain',
    'Suppression - Email',
    'Suppression - Competitor Domain',
    'Suppression - Opt-Out Email',
    'Suppression - Internal Exclusion',
    'Whitelist - Verified Domain'
  ], false);
  applyDropdownToRange(sheet.getRange('C2:C500'), ['Active', 'Inactive'], false);
}

/**
 * Formats, Populates & Dynamically Refreshes Tab 11: "Agent Settings"
 * Layout: Category | Setting Key | Value | Status / Description
 */
function setupAgentSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Agent Settings');

  const headers = ['Category', 'Setting Key', 'Value', 'Status / Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet);

  const scriptProps = PropertiesService.getScriptProperties().getProperties();

  function checkPropertyConfigured(primaryKey, aliases) {
    const searchKeys = [primaryKey].concat(aliases || []).map(function(k) {
      return k.toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    for (let k in scriptProps) {
      const cleanK = k.toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (searchKeys.indexOf(cleanK) >= 0) {
        const val = String(scriptProps[k] || '').trim();
        if (val && val.toUpperCase().indexOf('PASTE_') < 0) return true;
      }
    }
    return false;
  }

  const hasGemini = checkPropertyConfigured('GEMINI_API_KEY');
  const hasOpenAi = checkPropertyConfigured('OPENAI_API_KEY');
  const hasSmartlead = checkPropertyConfigured('SMARTLEAD_API_KEY');
  const hasInstantly = checkPropertyConfigured('INSTANTLY_API_KEY');
  const hasVerifier = checkPropertyConfigured('VERIFIER_API_KEY', ['EMAIL_VERIFICATION_API_KEY']);

  const protectedMsg = 'Configured in Script Properties (Protected)';

  const agentRows = [
    ['AI Routing', 'PRIMARY_AI_PROVIDER', 'GEMINI', 'Primary AI provider for discovery & persona mapping'],
    ['AI Models', 'GEMINI_MODEL', 'gemini-3.8-flash', 'Default Google Gemini model version'],
    ['AI Models', 'OPENAI_MODEL', 'gpt-4o', 'Default OpenAI model version'],
    ['API Credentials', 'GEMINI_API_KEY', hasGemini ? protectedMsg : '', 'Google Gemini API key (kept encrypted in Project Settings)'],
    ['API Credentials', 'OPENAI_API_KEY', hasOpenAi ? protectedMsg : '', 'OpenAI API key (kept encrypted in Project Settings)'],
    ['Search Grounding', 'GOOGLE_SEARCH_GROUNDING', 'TRUE', 'Enable live web search grounding citations for Gemini'],
    ['Email Verification', 'VERIFIER_PROVIDER', 'ZeroBounce', 'Default email verification vendor'],
    ['API Credentials', 'VERIFIER_API_KEY', hasVerifier ? protectedMsg : '', 'Email verifier API key (kept encrypted in Project Settings)'],
    ['Outreach Sequencers', 'SMARTLEAD_API_KEY', hasSmartlead ? protectedMsg : '', 'Smartlead sequencer API key (kept encrypted in Project Settings)'],
    ['Outreach Sequencers', 'INSTANTLY_API_KEY', hasInstantly ? protectedMsg : '', 'Instantly sequencer API key (kept encrypted in Project Settings)'],
    ['Budget Guardrails', 'DAILY_SPEND_CAP_USD', '10.00', 'Maximum daily API spend limit in USD before hard exit'],
    ['Budget Guardrails', 'MIN_VERIFIER_CREDITS', '500', 'Minimum email verification credits before email alert']
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, agentRows.length, 4).setValues(agentRows);
  } else {
    const lastR = sheet.getLastRow();
    const existingData = sheet.getRange(2, 1, lastR - 1, 4).getValues();

    for (let i = 0; i < existingData.length; i++) {
      const rowNum = i + 2;
      const key = String(existingData[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');

      if (key === 'GEMINIAPIKEY') {
        sheet.getRange(rowNum, 3).setValue(hasGemini ? protectedMsg : '');
      } else if (key === 'OPENAIAPIKEY') {
        sheet.getRange(rowNum, 3).setValue(hasOpenAi ? protectedMsg : '');
      } else if (key === 'VERIFIERAPIKEY' || key === 'EMAILVERIFICATIONAPIKEY') {
        sheet.getRange(rowNum, 3).setValue(hasVerifier ? protectedMsg : '');
      } else if (key === 'SMARTLEADAPIKEY') {
        sheet.getRange(rowNum, 3).setValue(hasSmartlead ? protectedMsg : '');
      } else if (key === 'INSTANTLYAPIKEY') {
        sheet.getRange(rowNum, 3).setValue(hasInstantly ? protectedMsg : '');
      }
    }

    const existingKeyList = existingData.map(function(r) {
      return String(r[1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    agentRows.forEach(function(row) {
      const cleanKey = row[1].toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (existingKeyList.indexOf(cleanKey) < 0) {
        sheet.appendRow(row);
      }
    });
  }

  const agentKeyDropdowns = {
    'PRIMARYAIPROVIDER': ['GEMINI', 'OPENAI'],
    'GEMINIMODEL': ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro'],
    'OPENAIMODEL': ['gpt-4o', 'gpt-4o-mini', 'o1-mini'],
    'GOOGLESEARCHGROUNDING': ['TRUE', 'FALSE'],
    'VERIFIERPROVIDER': ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None']
  };

  const finalLastR = sheet.getLastRow();
  if (finalLastR > 1) {
    const finalKeys = sheet.getRange(2, 2, finalLastR - 1, 1).getValues();
    for (let r = 0; r < finalKeys.length; r++) {
      const cleanK = String(finalKeys[r][0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (agentKeyDropdowns[cleanK]) {
        applyDropdownToRange(sheet.getRange(r + 2, 3), agentKeyDropdowns[cleanK], false);
      }
    }
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 420);
}

/**
 * Initializes and Formats "API_Usage_Log" Tab (Tab 2)
 */
function setupApiUsageLogTab(ss) {
  const sheet = getOrCreateSheet(ss, 'API_Usage_Log');
  const headers = [
    'Timestamp', 'Campaign ID', 'Service Name', 'Operation', 'HTTP Status',
    'Input Tokens / Units', 'Output Tokens', 'Est. Cost ($)', 'Status', 'Details / Error'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  styleHeaderRow(sheet);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 140);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 260);
}

/**
 * Helper: Applies high-contrast navy on light blue header styling
 */
function styleHeaderRow(sheet) {
  if (sheet.getLastColumn() > 0) {
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  }
}

/**
 * Helper: Builds and applies a Google Sheets Data Validation Dropdown to a Range
 */
function applyDropdownToRange(range, items, allowInvalid) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  range.setDataValidation(rule);
}

/**
 * Helper: Gets existing sheet or creates it safely
 */
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}