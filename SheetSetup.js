/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AUTOMATED SHEET & DROPDOWN SETUP ENGINE
 * ============================================================================
 * File: SheetSetup.gs
 * 
 * PURPOSE:
 * 1. Pre-populates "Campaign Configuration" with all required parameters.
 * 2. Injects native Google Sheets dropdown menus across ALL tabs to eliminate
 *    manual typing and human error.
 * 3. Ensures 100% column schema consistency across the entire workbook.
 * ============================================================================
 */

function setupAllSheetsAndDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Campaign Configuration with Parameter Table & In-Cell Dropdowns
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

  // 7. Setup Agent Settings Tab with Dropdowns
  setupAgentSettingsTabDropdowns(ss);

  ss.toast('All tabs formatted and native dropdown menus applied successfully!', '✅ Setup Complete', 10);
  Logger.log('>>> ALL TABS & DROPDOWNS CONFIGURED SUCCESSFULLY <<<');
}

/**
 * 1. Formats "Campaign Configuration" with Parameters, Default Values, and Dropdowns
 */
function setupCampaignConfigurationTab(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaign Configuration');
  sheet.clear();

  const headers = ['Parameter', 'Configured Value', 'Description'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  const configRows = [
    ['CAMPAIGN_ID', 'CMP-2026-001', 'Unique campaign identifier tag'],
    ['CAMPAIGN_NAME', 'Enterprise Scope 3 & ESG Supply Chain Readiness', 'Human-readable campaign title'],
    ['DEPARTMENT_OWNER', 'ESG Advisory', 'Department managing this campaign'],
    ['TARGET_INDUSTRIES', 'Freight & Logistics, Industrial Equipment', 'Target industry verticals'],
    ['TARGET_GEOGRAPHY', 'United States & Canada', 'Target geographic focus'],
    ['COMPANY_SIZE', '50 - 250 employees', 'Ideal employee size range'],
    ['RESEARCH_DEPTH', 'Deep Persona', 'Depth of web & executive footprint scanning'],
    ['MIN_COMPANY_FIT_SCORE', '80', 'Threshold to qualify companies for persona research'],
    ['MIN_PERSONA_FIT_SCORE', '85', 'Threshold for automated email draft approval'],
    ['MAX_COMPANIES_LIMIT', '25', 'Max target companies to evaluate in this run'],
    ['MAX_PERSONAS_PER_COMPANY', '2', 'Maximum decision-makers mapped per company'],
    ['SENDER_NAME', 'Alex Mercer', 'Sender name appearing in outreach copy & email signature'],
    ['SENDER_COMPANY', 'eValuater Solutions', 'Sender company name appearing in email body'],
    ['SENDER_TITLE', 'Solutions Consultant', 'Sender job title in signature'],
    ['DISPATCH_MODE', 'Google Workspace (Drafts Only)', 'Outreach dispatch channel'],
    ['ENABLE_SEARCH_GROUNDING', 'TRUE', 'Enables real-time Google Search grounding with Gemini'],
    ['EMAIL_VERIFICATION_PROVIDER', 'ZeroBounce', 'Selected email verification service'],
    ['STATUS', 'Active', 'Operational status of this campaign configuration']
  ];

  sheet.getRange(2, 1, configRows.length, 3).setValues(configRows);

  // Apply In-Cell Dropdown Validations to Column B for specific parameters
  const dropdownRules = {
    3: ['ESG Advisory', 'Enterprise Sales', 'Procurement', 'Marketing', 'Executive Leadership'], // Row 4 (idx 3): DEPARTMENT_OWNER
    4: ['Freight & Logistics, Industrial Equipment', 'FinTech / SaaS', 'Healthcare & Pharma', 'Energy & CleanTech', 'Retail & Consumer Goods'], // Row 5: TARGET_INDUSTRIES
    5: ['United States & Canada', 'United States Only', 'Canada Only', 'Europe / UK', 'Global'], // Row 6: TARGET_GEOGRAPHY
    6: ['10 - 50 employees', '50 - 250 employees', '250 - 1000 employees', '1000+ Enterprise'], // Row 7: COMPANY_SIZE
    7: ['Standard', 'Deep Persona', 'Comprehensive / Forensic'], // Row 8: RESEARCH_DEPTH
    8: ['70', '75', '80', '85', '90'], // Row 9: MIN_COMPANY_FIT_SCORE
    9: ['75', '80', '85', '90', '95'], // Row 10: MIN_PERSONA_FIT_SCORE
    10: ['10', '25', '50', '100', '250'], // Row 11: MAX_COMPANIES_LIMIT
    11: ['1', '2', '3', '5'], // Row 12: MAX_PERSONAS_PER_COMPANY
    15: ['Google Workspace (Drafts Only)', 'Google Workspace (Direct Send)', 'Smartlead', 'Instantly', 'Custom Webhook'], // Row 16: DISPATCH_MODE
    16: ['TRUE', 'FALSE'], // Row 17: ENABLE_SEARCH_GROUNDING
    17: ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None'], // Row 18: EMAIL_VERIFICATION_PROVIDER
    18: ['Draft', 'Active', 'Paused', 'Completed'] // Row 19: STATUS
  };

  Object.keys(dropdownRules).forEach(rowOffset => {
    const rowNum = parseInt(rowOffset, 10) + 1; // header offset
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownRules[rowOffset], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(rowNum, 2).setDataValidation(rule);
  });

  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 400);
}

/**
 * 2. Injects Dropdowns into "Campaigns" Tab
 */
function setupCampaignsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaigns');
  
  // Column D: Status
  applyDropdownToRange(sheet.getRange('D2:D500'), [
    'Draft', 'Active', 'Paused', 'Completed', 'Pilot Active'
  ]);

  // Column E: Company Limit
  applyDropdownToRange(sheet.getRange('E2:E500'), [
    '10', '25', '50', '100', '250'
  ]);

  // Column F: Research Depth
  applyDropdownToRange(sheet.getRange('F2:F500'), [
    'Standard', 'Deep Persona', 'Comprehensive / Forensic'
  ]);
}

/**
 * 3. Injects Dropdowns into "Prospect Review Queue" Tab
 */
function setupReviewQueueTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Prospect Review Queue');

  // Column M (Col 13): Priority Tier
  applyDropdownToRange(sheet.getRange('M2:M1000'), ['Tier 1', 'Tier 2', 'Tier 3']);

  // Column R (Col 18): Review Status
  applyDropdownToRange(sheet.getRange('R2:R1000'), ['Needs Review', 'Approved', 'Rejected']);

  // Column V (Col 22): Human Approval Status
  applyDropdownToRange(sheet.getRange('V2:V1000'), [
    'Approved', 'Approved (Auto-Pilot)', 'Pending Review', 'Rejected (Suppression Guard)', 'Rejected'
  ]);

  // Column Y (Col 25): Send Status
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), [
    'Queued', 'Draft Created in Gmail', 'Sent', 'Suppressed - Not Sent', 'Draft Failed'
  ]);

  // Column AP (Col 42): Compliance Status
  applyDropdownToRange(sheet.getRange('AP2:AP1000'), [
    'Passed', 'Failed - Suppressed', 'Failed - Role Based', 'Failed - Invalid Syntax'
  ]);

  // Column AQ (Col 43): Suppression Check
  applyDropdownToRange(sheet.getRange('AQ2:AQ1000'), [
    'Clear', 'Blocked: Email in Lists & Settings', 'Blocked: Domain in Lists & Settings', 'Blocked: Role-Based Account', 'Blocked: Bad Email Syntax'
  ]);
}

/**
 * 4. Injects Dropdowns into "Company Intelligence" Tab
 */
function setupCompanyIntelligenceTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Company Intelligence');

  // Column J (Col 10): ESG Reporting Maturity
  applyDropdownToRange(sheet.getRange('J2:J1000'), [
    'Early Stage', 'Developing', 'Mature', 'Advanced'
  ]);

  // Column Q (Col 17): eValuater Solution Fit
  applyDropdownToRange(sheet.getRange('Q2:Q1000'), [
    'High', 'Medium', 'Low', 'Disqualified - Direct Competitor'
  ]);

  // Column V (Col 22): Verification Status
  applyDropdownToRange(sheet.getRange('V2:V1000'), [
    'Verified', 'Unverified', 'Disqualified'
  ]);

  // Column W (Col 23): Human Review Status
  applyDropdownToRange(sheet.getRange('W2:W1000'), [
    'Pending', 'Accepted', 'Rejected'
  ]);
}

/**
 * 5. Injects Dropdowns into "Persona Leads" Tab
 */
function setupPersonaLeadsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Persona Leads');

  // Column I (Col 9): Persona Type
  applyDropdownToRange(sheet.getRange('I2:I1000'), [
    'Economic Buyer', 'Technical Evaluator', 'Key Decision Maker', 'Influencer', 'Disqualified / Competitor'
  ]);

  // Column O (Col 15): Verification Status
  applyDropdownToRange(sheet.getRange('O2:O1000'), [
    'Verified', 'Unverified', 'Pending'
  ]);

  // Column S (Col 19): Review Status
  applyDropdownToRange(sheet.getRange('S2:S1000'), [
    'Ready for Outreach', 'Needs Review', 'Approved', 'Rejected'
  ]);

  // Column Y (Col 25): Email Verification Status
  applyDropdownToRange(sheet.getRange('Y2:Y1000'), [
    'valid', 'risky', 'invalid', 'catch-all', 'unverified'
  ]);
}

/**
 * 6. Injects Dropdowns into "Lists & Settings" Tab
 */
function setupListsAndSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Lists & Settings');

  // Column A: List Name
  applyDropdownToRange(sheet.getRange('A2:A500'), [
    'Suppression - Domain',
    'Suppression - Email',
    'Suppression - Competitor Domain',
    'Suppression - Opt-Out Email',
    'Suppression - Internal Exclusion',
    'Whitelist - Verified Domain'
  ]);

  // Column C: Active Status
  applyDropdownToRange(sheet.getRange('C2:C500'), [
    'Active', 'Inactive'
  ]);
}

/**
 * 7. Injects Dropdowns into "Agent Settings" Tab
 */
function setupAgentSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Agent Settings');
  
  // Row 2 Col C (Primary AI Provider)
  applyDropdownToRange(sheet.getRange('C2'), ['gemini', 'openai']);

  // Row 3 Col C (Gemini Model)
  applyDropdownToRange(sheet.getRange('C3'), ['gemini-2.5-flash', 'gemini-1.5-pro', 'gemini-2.0-flash']);

  // Row 4 Col C (OpenAI Model)
  applyDropdownToRange(sheet.getRange('C4'), ['gpt-4o', 'gpt-4o-mini', 'o1-mini']);

  // Row 7 Col C (Verifier Provider)
  applyDropdownToRange(sheet.getRange('C7'), ['ZeroBounce', 'MillionVerifier', 'NeverBounce', 'Custom HTTPS', 'None']);

  // Row 9 Col C (Google Search Grounding)
  applyDropdownToRange(sheet.getRange('C9'), ['TRUE', 'FALSE']);
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