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
 * 3. Enforces high-contrast readable headers (#174ea6 on #e8f0fe) across tabs.
 * 4. Ensures 100% column schema consistency across the entire workbook.
 * 5. Pre-populates "Agent Settings" and dynamically refreshes Script Properties status.
 * ============================================================================
 */

function setupAllSheetsAndDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Campaign Configuration with Full 38 Parameters & Dynamic Dropdowns
  setupCampaignConfigurationTab(ss);
  
  // 2. Setup Campaigns Tab with Dropdowns
  setupCampaignsTabDropdowns(ss);
  
  // 3. Setup Prospect Review Queue Tab with Dropdowns
  setupReviewQueueTabDropdowns(ss);
  
  // 4. Setup Company Intelligence Tab with Dropdowns
  setupCompanyIntelligenceTabDropdowns(ss);
  
  // 5. Setup Persona Leads Tab with Dropdowns
  setupPersonaLeadsTabDropdowns(ss);
  
  // 6. Setup Lists & Settings Tab with Dropdowns
  setupListsAndSettingsTabDropdowns(ss);

  // 7. Setup Agent Settings Tab with Structured Rows & Dynamic Key Status
  setupAgentSettingsTabDropdowns(ss);

  SpreadsheetApp.flush();
  ss.toast('All tabs formatted and dynamic dropdown menus applied successfully!', '✅ Setup Complete', 10);
  Logger.log('>>> ALL TABS & DROPDOWNS CONFIGURED SUCCESSFULLY <<<');
}

/**
 * 1. Formats "Campaign Configuration" with 38 Parameters, Canonical Values, and Dynamic Dropdowns
 * STRICT: Zero Schema Mutations (Locked 38 Parameters)
 */
function setupCampaignConfigurationTab(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaign Configuration');
  sheet.clear();

  const headers = ['Parameter', 'Configured Value', 'Description'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);

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

  // Dynamic Parameter Dropdown Rules Definition (Key-Based)
  const dynamicRules = {
    'DEPARTMENT': ['Procurement', 'ESG Advisory', 'Enterprise Sales', 'Cybersecurity & IT', 'Operations & Compliance', 'Enterprise SaaS / IT', 'FinTech & Banking', 'HR & Talent Ops', 'Marketing', 'Executive Leadership'],
    'TARGET_COUNTRIES': ['United States & Canada', 'United States Only', 'Canada Only', 'Sweden, Norway, Denmark', 'United Kingdom & Europe', 'Global'],
    'TARGET_INDUSTRIES': ['Freight & Logistics, Industrial Equipment', 'Energy, CleanTech & Utilities', 'Manufacturing, Industrial, BioTech', 'Financial Services / FinTech SaaS', 'Retail & Consumer Goods'],
    'COMPANY_SIZE_RANGE': ['10 - 50 employees', '50 - 250 employees', '250 - 1000 employees', '1000+ Enterprise'],
    'MAX_COMPANIES': ['3', '5', '10', '25', '50'],
    'RESEARCH_DEPTH': ['Standard', 'Deep Persona', 'Comprehensive / Forensic'],
    'EVIDENCE_PERIOD_MONTHS': ['6', '12', '24'],
    'MIN_COMPANY_SCORE': ['70', '75', '80', '85', '90'],
    'MIN_PERSONA_SCORE': ['75', '80', '85', '90', '95'],
    'MIN_EMAIL_PRIORITY_SCORE': ['70', '75', '80', '85', '90'],
    'AI_PROVIDER': ['GEMINI', 'OPENAI'],
    'PRIMARY_AI_PROVIDER': ['GEMINI', 'OPENAI'],
    'AI_MODEL': ['gemini-3.8-flash', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro', 'gpt-4o', 'gpt-4o-mini', 'o1-mini'],
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

  const lastRow = sheet.getLastRow();
  const paramKeys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let r = 0; r < paramKeys.length; r++) {
    const key = String(paramKeys[r][0] || '').trim().toUpperCase();
    if (dynamicRules[key]) {
      applyDropdownToRange(sheet.getRange(r + 2, 2), dynamicRules[key]);
    }
  }

  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(3, 400);
}

/**
 * 2. Injects Dropdowns into "Campaigns" Tab (Locked 28 Columns)
 */
function setupCampaignsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaigns');
  styleHeaderRow(sheet);
  
  applyDropdownToRange(sheet.getRange('D2:D500'), ['Draft', 'Active', 'Paused', 'Completed', 'Pilot Active']);
  applyDropdownToRange(sheet.getRange('H2:H500'), ['3', '5', '10', '25', '50']);
  applyDropdownToRange(sheet.getRange('I2:I500'), ['Standard', 'Deep Persona', 'Comprehensive / Forensic']);
  applyDropdownToRange(sheet.getRange('R2:R500'), ['GEMINI', 'OPENAI']);
  applyDropdownToRange(sheet.getRange('T2:T500'), ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None']);
  applyDropdownToRange(sheet.getRange('V2:V500'), ['Workspace Gmail (Drafts Only)', 'Workspace Gmail (Direct Send)', 'Smartlead', 'Instantly', 'Custom Webhook']);
}

/**
 * 3. Injects Dropdowns into "Prospect Review Queue" Tab
 */
function setupReviewQueueTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Prospect Review Queue');
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('M2:M1000'), ['Tier 1', 'Tier 2', 'Tier 3']);
  applyDropdownToRange(sheet.getRange('R2:R1000'), ['Needs Review', 'Approved', 'Rejected']);
  applyDropdownToRange(sheet.getRange('V2:V1000'), ['Approved', 'Approved (Auto-Pilot)', 'Needs Review', 'Pending Review', 'Rejected (Suppression Guard)', 'Rejected']);
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), ['Queued', 'Draft Created in Gmail', 'Draft Created', 'Sent', 'Suppressed - Not Sent', 'Draft Failed']);
  applyDropdownToRange(sheet.getRange('AP2:AP1000'), ['Passed', 'Failed - Suppressed', 'Failed - Role Based', 'Failed - Invalid Syntax']);
  applyDropdownToRange(sheet.getRange('AQ2:AQ1000'), ['Clear', 'Blocked: Email in Lists & Settings', 'Blocked: Domain in Lists & Settings', 'Blocked: Role-Based Account', 'Blocked: Bad Email Syntax']);
}

/**
 * 4. Injects Dropdowns into "Company Intelligence" Tab
 */
function setupCompanyIntelligenceTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Company Intelligence');
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('J2:J1000'), ['Early Stage', 'Developing / Spreadsheet-reliant', 'Developing', 'Mature', 'Advanced', 'N/A']);
  applyDropdownToRange(sheet.getRange('Q2:Q1000'), ['High', 'Medium', 'Low', 'Disqualified - Direct Competitor']);
  applyDropdownToRange(sheet.getRange('V2:V1000'), ['Verified', 'Unverified', 'Disqualified', 'Needs review']);
  applyDropdownToRange(sheet.getRange('W2:W1000'), ['Accepted', 'Pending', 'Rejected']);
}

/**
 * 5. Injects Dropdowns into "Persona Leads" Tab
 */
function setupPersonaLeadsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Persona Leads');
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('I2:I1000'), ['Economic Buyer', 'Technical Evaluator', 'Key Decision Maker', 'Influencer', 'Disqualified / Competitor']);
  applyDropdownToRange(sheet.getRange('O2:O1000'), ['Verified', 'Unverified', 'Pending', 'Disqualified']);
  applyDropdownToRange(sheet.getRange('S2:S1000'), ['Ready for Outreach', 'Needs Review', 'Approved', 'Rejected']);
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), ['valid', 'risky', 'invalid', 'catch-all', 'unverified']);
}

/**
 * 6. Injects Dropdowns into "Lists & Settings" Tab
 */
function setupListsAndSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Lists & Settings');
  styleHeaderRow(sheet);

  applyDropdownToRange(sheet.getRange('A2:A500'), [
    'Suppression - Domain',
    'Suppression - Email',
    'Suppression - Competitor Domain',
    'Suppression - Opt-Out Email',
    'Suppression - Internal Exclusion',
    'Whitelist - Verified Domain'
  ]);
  applyDropdownToRange(sheet.getRange('C2:C500'), ['Active', 'Inactive']);
}

/**
 * 7. Formats, Populates & Dynamically Refreshes Tab 10: "Agent Settings"
 * Layout: Category | Setting Key | Value | Status / Description
 * Dynamically checks Script Properties and sets "Configured in Script Properties (Protected)"
 */
function setupAgentSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Agent Settings');
  
  const headers = ['Category', 'Setting Key', 'Value', 'Status / Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet);

  // Check existing Script Properties securely with alias tolerance
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  
  function checkPropertyConfigured(primaryKey, aliases) {
    const searchKeys = [primaryKey].concat(aliases || []).map(function(k) {
      return k.toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    for (let k in scriptProps) {
      const cleanK = k.toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (searchKeys.indexOf(cleanK) >= 0) {
        const val = String(scriptProps[k] || '').trim();
        if (val && !val.includes('PASTE_')) return true;
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

  // 10 Canonical Operational Settings Rows
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
    ['Outreach Sequencers', 'INSTANTLY_API_KEY', hasInstantly ? protectedMsg : '', 'Instantly sequencer API key (kept encrypted in Project Settings)']
  ];

  // If sheet is empty below header, write entire table
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, agentRows.length, 4).setValues(agentRows);
  } else {
    // If rows already exist, refresh Column C for API Key status while preserving model selections
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

    // Append any missing rows if table was incomplete
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

  // Apply Dropdowns to Column C (Value)
  applyDropdownToRange(sheet.getRange('C2'), ['GEMINI', 'OPENAI']);
  applyDropdownToRange(sheet.getRange('C3'), ['gemini-3.8-flash', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro']);
  applyDropdownToRange(sheet.getRange('C4'), ['gpt-4o', 'gpt-4o-mini', 'o1-mini']);
  applyDropdownToRange(sheet.getRange('C7'), ['TRUE', 'FALSE']);
  applyDropdownToRange(sheet.getRange('C8'), ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None']);

  // Column Widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 420);
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
function applyDropdownToRange(range, items) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(false)
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