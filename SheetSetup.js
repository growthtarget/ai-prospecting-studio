/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AUTOMATED SHEET & DROPDOWN SETUP ENGINE
 * ============================================================================
 * File: SheetSetup.gs / SheetSetup.js
 * 
 * PURPOSE:
 * 1. Manages "Lists & Settings" client registry ("Client - Account") & suppression dropdowns,
 *    keeping Client - Account rows cleanly pinned at the top.
 * 2. Preserves locked 38 parameters in "Campaign Configuration" & 28 columns in "Campaigns".
 * 3. Dynamically binds Client / Owner dropdowns from active "Client - Account" rows.
 * 4. Enforces plain integer formatting (0) on limit/score columns to prevent date conversion.
 * 5. Enforces high-contrast readable headers (#174ea6 on #e8f0fe) across all 13 tabs.
 * 6. Pre-populates "Agent Settings" with Growth Target defaults + per-client status & guardrails.
 * 7. Filters "Agent Settings" cleanly to the active Client / Owner.
 * 8. Initializes and formats "API_Usage_Log" (Tab 2).
 * ============================================================================
 */

function setupAllSheetsAndDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  // 1. Setup Dashboard if brand new (non-destructive if already populated)
  setupDashboardTabIfEmpty_(ss);

  // 2. Setup Lists & Settings first so active "Client - Account" rows are seeded, sorted & ready
  setupListsAndSettingsTabDropdowns(ss);

  // 3. Migrate legacy department strings (e.g. "ESG Advisory") to real Client names before strict validations
  migrateLegacyClientNames_(ss);

  // 4. Setup API Usage Log Tab (Tab 2)
  setupApiUsageLogTab(ss);

  // 5. Setup Campaigns Tab with 28 Canonical Headers, Number Formats & Dynamic Client Dropdown
  setupCampaignsTabDropdowns(ss);

  // 6. Setup Company Intelligence Tab with 25 Canonical Headers & Dropdowns
  setupCompanyIntelligenceTabDropdowns(ss);

  // 7. Setup Persona Leads Tab with 29 Canonical Headers & Dropdowns
  setupPersonaLeadsTabDropdowns(ss);

  // 8. Ensure Persona Mapper Settings (39 Canonical Headers) & Plain Integer Formats
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

  // 9. Setup Prospect Review Queue Tab with 45 Canonical Headers & Dropdowns
  setupReviewQueueTabDropdowns(ss);

  // 10. Setup Campaign Configuration with Locked 38 Parameters & Dynamic Client Dropdowns
  setupCampaignConfigurationTab(ss);

  // 11. Ensure Company Database (10 Canonical Headers, Col J = Campaign ID)
  setupHeaderOnlyTab_(ss, 'Company Database', [
    'Company ID', 'Company Name', 'Domain', 'Country', 'Industry',
    'Employee Size', 'Discovery Method', 'Discovery Date', 'Status', 'Campaign ID'
  ]);

  // 12. Ensure Research Evidence (9 Canonical Headers)
  setupHeaderOnlyTab_(ss, 'Research Evidence', [
    'Evidence ID', 'Campaign ID', 'Entity Name', 'Entity Type',
    'Search Query Used', 'Grounding Evidence URL', 'Extracted Snippet / Signal',
    'Verification Score', 'Date Captured'
  ]);

  // 13. Setup Agent Settings Tab with Local Defaults + Per-Client Blocks & Active Filter
  setupAgentSettingsTabDropdowns(ss);

  // 14. Ensure Execution Logs (6 Canonical Headers)
  setupHeaderOnlyTab_(ss, 'Execution Logs', [
    'Timestamp', 'Campaign ID', 'Action', 'Target', 'Status', 'Details'
  ]);

  SpreadsheetApp.flush();
  ss.toast('All 13 tabs formatted, Lists & Settings organized, and dropdowns applied!', '✅ Setup Complete', 10);
  Logger.log('>>> ALL 13 TABS & DROPDOWNS CONFIGURED SUCCESSFULLY <<<');
}

/**
 * Helper: Identifies local/agency brand ("Growth Target") safely inside SheetSetup.
 */
function isSetupLocalBrand_(clientName) {
  const clean = String(clientName || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  return !clean || clean === 'GROWTHTARGET' || clean === 'LOCAL' || clean === 'AGENCY' || clean === 'DEFAULT';
}

/**
 * Helper: Builds canonical Script Property prefix e.g. "eValuater" -> "CLIENT_EVALUATER_".
 */
function normalizeSetupClientPrefix_(clientName) {
  const clean = String(clientName || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return clean ? ('CLIENT_' + clean + '_') : '';
}

/**
 * Reads active Client / Account names from Tab 13 ("Lists & Settings").
 * Never resurrects a client explicitly marked Inactive in Lists & Settings.
 */
function getRegisteredClientsFromSettings_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const clients = [];
  const seen = {};
  let hasAnyClientRow = false;

  if (spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Lists & Settings');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const listName = String(data[i][0] || '').trim().toLowerCase();
        const val = String(data[i][1] || '').trim();
        const status = String(data[i][2] || '').trim().toLowerCase();
        const isActive = (status === 'active' || status === 'true' || status === '');

        if (val && (listName === 'client - account' || listName === 'client / owner' || listName === 'client')) {
          hasAnyClientRow = true;
          if (isActive) {
            const norm = val.toLowerCase();
            if (!seen[norm]) {
              clients.push(val);
              seen[norm] = true;
            }
          }
        }
      }
    }
  }

  if (clients.length === 0) {
    return hasAnyClientRow ? ['Growth Target'] : ['Growth Target', 'Pilot Acme'];
  }
  return clients;
}

/**
 * Returns Client - Account names explicitly marked Inactive in Tab 13 ("Lists & Settings").
 * Used by the Control Panel to archive Inactive clients out of the Client / Owner dropdown
 * even if historical campaigns still exist in the "Campaigns" tab.
 */
function getInactiveClientsFromSettings_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const inactive = [];
  const activeSet = {};
  const seenInactive = {};

  if (spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Lists & Settings');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const listName = String(data[i][0] || '').trim().toLowerCase();
        const val = String(data[i][1] || '').trim();
        const status = String(data[i][2] || '').trim().toLowerCase();

        if (!val || !(listName === 'client - account' || listName === 'client / owner' || listName === 'client')) {
          continue;
        }
        const isActive = (status === 'active' || status === 'true' || status === '');
        if (isActive) {
          activeSet[val.toLowerCase()] = true;
        } else if (!seenInactive[val.toLowerCase()]) {
          inactive.push(val);
          seenInactive[val.toLowerCase()] = true;
        }
      }
    }
  }

  return inactive.filter(function(name) {
    return !activeSet[name.toLowerCase()];
  });
}

/**
 * Organizes Tab 13 ("Lists & Settings") cleanly so all "Client - Account" rows stay pinned
 * at the top (Growth Target first, then Active clients A-Z, then Inactive clients A-Z),
 * followed by Whitelist rules and Suppression rules grouped by List name.
 */
function sortListsAndSettingsTab_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const sheet = spreadsheet.getSheetByName('Lists & Settings');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const lastRow = sheet.getLastRow();
  const rawRows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  const cleanRows = [];
  const seenPairs = {};

  for (let i = 0; i < rawRows.length; i++) {
    const listName = String(rawRows[i][0] || '').trim();
    const val = String(rawRows[i][1] || '').trim();
    const status = String(rawRows[i][2] || '').trim();

    if (!listName && !val && !status) continue;

    // Deduplicate identical (List name + Value) rows if both exist
    if (listName && val) {
      const pairKey = listName.toLowerCase() + '||' + val.toLowerCase();
      if (seenPairs[pairKey]) continue;
      seenPairs[pairKey] = true;
    }

    cleanRows.push([listName, val, status]);
  }

  function getCategoryRank(listName) {
    const l = String(listName || '').trim().toLowerCase();
    if (l === 'client - account' || l === 'client / owner' || l === 'client') return 0;
    if (l.indexOf('whitelist') >= 0) return 1;
    if (l.indexOf('suppression') >= 0 || l.indexOf('blacklist') >= 0 || l.indexOf('opt-out') >= 0 || l.indexOf('competitor') >= 0) return 2;
    return 3;
  }

  function isRowActive(status) {
    const s = String(status || '').trim().toLowerCase();
    return (s === 'active' || s === 'true' || s === '');
  }

  cleanRows.sort(function(a, b) {
    const rankA = getCategoryRank(a[0]);
    const rankB = getCategoryRank(b[0]);
    if (rankA !== rankB) return rankA - rankB;

    if (rankA === 0) {
      // Client - Account block: Active first, then Growth Target pinned #1, then alphabetical
      const actA = isRowActive(a[2]) ? 0 : 1;
      const actB = isRowActive(b[2]) ? 0 : 1;
      if (actA !== actB) return actA - actB;

      const localA = isSetupLocalBrand_(a[1]) ? 0 : 1;
      const localB = isSetupLocalBrand_(b[1]) ? 0 : 1;
      if (localA !== localB) return localA - localB;

      return String(a[1] || '').toLowerCase().localeCompare(String(b[1] || '').toLowerCase());
    }

    // Other blocks: group by List name, then Active first, then alphabetical by Value
    const typeCmp = String(a[0] || '').toLowerCase().localeCompare(String(b[0] || '').toLowerCase());
    if (typeCmp !== 0) return typeCmp;

    const actA = isRowActive(a[2]) ? 0 : 1;
    const actB = isRowActive(b[2]) ? 0 : 1;
    if (actA !== actB) return actA - actB;

    return String(a[1] || '').toLowerCase().localeCompare(String(b[1] || '').toLowerCase());
  });

  const clearCount = Math.max(lastRow - 1, cleanRows.length);
  if (clearCount > 0) {
    sheet.getRange(2, 1, clearCount, 3).clearContent();
  }
  if (cleanRows.length > 0) {
    sheet.getRange(2, 1, cleanRows.length, 3).setValues(cleanRows);
  }
}

/**
 * Refreshes the Client / Owner dropdown validation rules on "Campaign Configuration" (Row 4)
 * and "Campaigns" (Col C) using the latest active "Client - Account" entries in "Lists & Settings".
 */
function refreshClientDropdownValidations_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return [];

  const clientList = getRegisteredClientsFromSettings_(spreadsheet);

  // 1. Update Campaign Configuration Row for DEPARTMENT
  const cfgSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const keys = cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 1).getValues();
    for (let r = 0; r < keys.length; r++) {
      const key = String(keys[r][0] || '').trim().toUpperCase();
      if (key === 'DEPARTMENT') {
        applyDropdownToRange(cfgSheet.getRange(r + 2, 2), clientList, false);
        break;
      }
    }
  }

  // 2. Update Campaigns Tab Column C (Department / Client Owner)
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet) {
    applyDropdownToRange(campSheet.getRange('C2:C500'), clientList, false);
  }

  return clientList;
}

/**
 * Safely migrates legacy internal department strings (e.g. "ESG Advisory", "Procurement")
 * to "Growth Target" (for production/CMP-005) or "Pilot Acme" (for PILOT- campaigns).
 * Clears old cell validations before writing so legacy strict lists never block migration.
 */
function migrateLegacyClientNames_(ss) {
  if (!ss) return;
  const legacySet = {
    'esg advisory': true,
    'procurement': true,
    'enterprise sales': true,
    'cybersecurity & it': true,
    'operations & compliance': true,
    'enterprise saas / it': true,
    'fintech & banking': true,
    'hr & talent ops': true,
    'marketing': true,
    'executive leadership': true
  };

  function mapLegacyClient(campaignId, currentVal) {
    const clean = String(currentVal || '').trim();
    if (!clean || legacySet[clean.toLowerCase()]) {
      return String(campaignId || '').trim().toUpperCase().indexOf('PILOT') === 0
        ? 'Pilot Acme'
        : 'Growth Target';
    }
    return clean;
  }

  // 1. Migrate Campaign Configuration (Row 4)
  const cfgSheet = ss.getSheetByName('Campaign Configuration');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const cfgData = cfgSheet.getDataRange().getValues();
    let cfgCampId = '';
    let deptRow = -1;
    let deptVal = '';
    for (let r = 1; r < cfgData.length; r++) {
      const k = String(cfgData[r][0] || '').trim().toUpperCase();
      if (k === 'CAMPAIGN_ID') cfgCampId = String(cfgData[r][1] || '').trim();
      if (k === 'DEPARTMENT') {
        deptRow = r + 1;
        deptVal = String(cfgData[r][1] || '').trim();
      }
    }
    if (deptRow > 1 && legacySet[deptVal.toLowerCase()]) {
      const targetCell = cfgSheet.getRange(deptRow, 2);
      targetCell.clearDataValidations();
      targetCell.setValue(mapLegacyClient(cfgCampId, deptVal));
    }
  }

  // 2. Migrate Campaigns Tab (Col C + JSON backup in Col AA)
  const campSheet = ss.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const lastR = campSheet.getLastRow();
    const lastC = campSheet.getLastColumn();
    const cData = campSheet.getRange(2, 1, lastR - 1, lastC).getValues();
    const headers = campSheet.getRange(1, 1, 1, lastC).getValues()[0].map(function(h) {
      return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    const jsonCol = headers.indexOf('CAMPAIGNCONFIGJSON');

    for (let r = 0; r < cData.length; r++) {
      const cid = String(cData[r][0] || '').trim();
      const dept = String(cData[r][2] || '').trim();
      if (cid && legacySet[dept.toLowerCase()]) {
        const mapped = mapLegacyClient(cid, dept);
        const cCell = campSheet.getRange(r + 2, 3);
        cCell.clearDataValidations();
        cCell.setValue(mapped);
        if (jsonCol >= 0 && cData[r][jsonCol]) {
          try {
            const parsed = JSON.parse(cData[r][jsonCol]);
            if (parsed && typeof parsed === 'object') {
              parsed.department = mapped;
              campSheet.getRange(r + 2, jsonCol + 1).setValue(JSON.stringify(parsed));
            }
          } catch (e) {}
        }
      }
    }
  }

  // 3. Migrate Dashboard KPI & Performance Rollup if displaying legacy label
  const dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet && dashSheet.getLastRow() >= 6) {
    const activeId = String(dashSheet.getRange('B2').getValue() || '').trim();
    const b6 = String(dashSheet.getRange('B6').getValue() || '').trim();
    if (legacySet[b6.toLowerCase()]) {
      dashSheet.getRange('B6').clearDataValidations();
      dashSheet.getRange('B6').setValue(mapLegacyClient(activeId, b6));
    }
    const lastDashRow = dashSheet.getLastRow();
    if (lastDashRow >= 18) {
      for (let r = 18; r <= lastDashRow; r++) {
        const rId = String(dashSheet.getRange(r, 1).getValue() || '').trim();
        const rDept = String(dashSheet.getRange(r, 3).getValue() || '').trim();
        if (rId && legacySet[rDept.toLowerCase()]) {
          dashSheet.getRange(r, 3).clearDataValidations();
          dashSheet.getRange(r, 3).setValue(mapLegacyClient(rId, rDept));
        }
      }
    }
  }
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
 * STRICT: Zero Schema Mutations (Locked 38 Parameters), non-destructive if already populated.
 */
function setupCampaignConfigurationTab(ss) {
  const sheet = getOrCreateSheet(ss, 'Campaign Configuration');
  const activeClients = getRegisteredClientsFromSettings_(ss);

  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    const headers = ['Parameter', 'Configured Value', 'Description'];
    sheet.appendRow(headers);

    const configRows = [
      ['CAMPAIGN_ID', 'CMP-2026-001', 'Active Campaign Identifier'],
      ['CAMPAIGN_NAME', 'Enterprise Scope 3 & ESG Supply Chain Readiness', 'Active Campaign Name'],
      ['DEPARTMENT', activeClients[0] || 'Growth Target', 'Client / Account Owner'],
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
      ['CUSTOM_BUSINESS_NEEDS', '- Managing Scope 3 vendor emissions and Tier-1 ESG supply chain transparency\n- High vendor questionnaire drop-off rate during procurement audits\n- Eliminating manual spreadsheet follow-ups with suppliers', 'Manual client business needs & buying triggers'],
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
  }

  styleHeaderRow(sheet);

  const strictRules = {
    'DEPARTMENT': activeClients,
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

  const activeClients = getRegisteredClientsFromSettings_(ss);
  applyDropdownToRange(sheet.getRange('C2:C500'), activeClients, false);
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
 * Injects Dropdowns into "Lists & Settings" Tab, seeds active Client accounts if missing,
 * and pins all "Client - Account" rows cleanly at the top.
 */
function setupListsAndSettingsTabDropdowns(ss) {
  const sheet = getOrCreateSheet(ss, 'Lists & Settings');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['List name', 'Value', 'Active status']);
  }
  styleHeaderRow(sheet);

  // Apply updated dropdown validation FIRST so seeding "Client - Account" never collides with old rules
  applyDropdownToRange(sheet.getRange('A2:A500'), [
    'Client - Account',
    'Suppression - Domain',
    'Suppression - Email',
    'Suppression - Competitor Domain',
    'Suppression - Opt-Out Email',
    'Suppression - Internal Exclusion',
    'Whitelist - Verified Domain'
  ], false);
  applyDropdownToRange(sheet.getRange('C2:C500'), ['Active', 'Inactive'], false);

  // Seed initial "Client - Account" entries if none exist yet
  let hasClientEntries = false;
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      const listType = String(rows[i][0] || '').trim().toLowerCase();
      if (listType === 'client - account' || listType === 'client / owner' || listType === 'client') {
        hasClientEntries = true;
        break;
      }
    }
  }

  if (!hasClientEntries) {
    sheet.appendRow(['Client - Account', 'Growth Target', 'Active']);
    sheet.appendRow(['Client - Account', 'Pilot Acme', 'Active']);
  }

  sortListsAndSettingsTab_(ss);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 140);
}

/**
 * Formats, Populates & Dynamically Refreshes Tab 11: "Agent Settings"
 * Layout: Category | Setting Key | Value | Status / Description
 * Includes Growth Target (Local Default) rows AND per-client rows for every external client.
 */
function setupAgentSettingsTabDropdowns(ss, activeClientOverride) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const sheet = getOrCreateSheet(spreadsheet, 'Agent Settings');

  const headers = ['Category', 'Setting Key', 'Value', 'Status / Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet);

  const scriptProps = PropertiesService.getScriptProperties().getProperties();

  function isValidPropVal(val) {
    const s = String(val || '').trim();
    if (!s) return false;
    const upper = s.toUpperCase();
    if (upper.indexOf('PASTE_') >= 0) return false;
    if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
    if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
    return true;
  }

  // Strictly checks LOCAL agency keys (ignores CLIENT_* keys)
  function checkLocalPropertyConfigured(primaryKey, aliases) {
    const searchKeys = [primaryKey].concat(aliases || []).map(function(k) {
      return k.toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    for (let k in scriptProps) {
      if (String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
      const cleanK = k.toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (searchKeys.indexOf(cleanK) >= 0 && isValidPropVal(scriptProps[k])) {
        return true;
      }
    }
    return false;
  }

  // Strictly checks EXTERNAL CLIENT keys matching CLIENT_{NAME}_{SUFFIX}
  function checkClientPropertyConfigured(prefix, suffixes) {
    const searchKeys = suffixes.map(function(suf) {
      return (prefix + suf).toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
    for (let k in scriptProps) {
      const cleanK = k.toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (searchKeys.indexOf(cleanK) >= 0 && isValidPropVal(scriptProps[k])) {
        return true;
      }
    }
    return false;
  }

  const hasGemini = checkLocalPropertyConfigured('GEMINI_API_KEY');
  const hasOpenAi = checkLocalPropertyConfigured('OPENAI_API_KEY');
  const hasSmartlead = checkLocalPropertyConfigured('SMARTLEAD_API_KEY');
  const hasInstantly = checkLocalPropertyConfigured('INSTANTLY_API_KEY');
  const hasVerifier = checkLocalPropertyConfigured('VERIFIER_API_KEY', [
    'EMAIL_VERIFICATION_API_KEY',
    'ZEROBOUNCE_API_KEY',
    'MILLIONVERIFIER_API_KEY',
    'NEVERBOUNCE_API_KEY'
  ]);

  const protectedMsg = 'Configured in Script Properties (Protected)';

  // 1. Local Agency ("Growth Target") & Global Defaults
  const localRows = [
    ['AI Routing', 'PRIMARY_AI_PROVIDER', 'GEMINI', 'Default AI provider for Growth Target (overridden per campaign)'],
    ['AI Models', 'GEMINI_MODEL', 'gemini-3.8-flash', 'Default Gemini model for Growth Target'],
    ['AI Models', 'OPENAI_MODEL', 'gpt-4o', 'Default OpenAI model for Growth Target'],
    ['API Credentials', 'GEMINI_API_KEY', hasGemini ? protectedMsg : '', 'Growth Target Gemini API key (encrypted in Script Properties)'],
    ['API Credentials', 'OPENAI_API_KEY', hasOpenAi ? protectedMsg : '', 'Growth Target OpenAI API key (encrypted in Script Properties)'],
    ['Search Grounding', 'GOOGLE_SEARCH_GROUNDING', 'TRUE', 'Enable live Google Search Grounding for Growth Target'],
    ['Email Verification', 'VERIFIER_PROVIDER', 'ZeroBounce', 'Default email verification vendor for Growth Target'],
    ['API Credentials', 'VERIFIER_API_KEY', hasVerifier ? protectedMsg : '', 'Growth Target verifier API key (encrypted in Script Properties)'],
    ['Outreach Sequencers', 'SMARTLEAD_API_KEY', hasSmartlead ? protectedMsg : '', 'Growth Target Smartlead key (encrypted in Script Properties)'],
    ['Outreach Sequencers', 'INSTANTLY_API_KEY', hasInstantly ? protectedMsg : '', 'Growth Target Instantly key (encrypted in Script Properties)'],
    ['Budget Guardrails', 'DAILY_SPEND_CAP_USD', '10.00', 'Growth Target daily API spend limit in USD (editable)'],
    ['Budget Guardrails', 'MIN_VERIFIER_CREDITS', '500', 'Growth Target min verifier credits before email alert (editable)']
  ];

  // Track exact live protection values & descriptions by normalized Setting Key
  const secretValueByKey = {
    'GEMINIAPIKEY': hasGemini ? protectedMsg : '',
    'OPENAIAPIKEY': hasOpenAi ? protectedMsg : '',
    'VERIFIERAPIKEY': hasVerifier ? protectedMsg : '',
    'EMAILVERIFICATIONAPIKEY': hasVerifier ? protectedMsg : '',
    'SMARTLEADAPIKEY': hasSmartlead ? protectedMsg : '',
    'INSTANTLYAPIKEY': hasInstantly ? protectedMsg : ''
  };
  const categorySyncByKey = {};
  const descriptionSyncByKey = {};
  const numericFormatByKey = {
    'DAILYSPENDCAPUSD': '0.00',
    'MINVERIFIERCREDITS': '0'
  };

  localRows.forEach(function(r) {
    const cleanKey = r[1].toUpperCase().replace(/[^A-Z0-9]/g, '');
    descriptionSyncByKey[cleanKey] = r[3];
  });

  // 2. Per-Client Rows for every active External Client in "Lists & Settings"
  const clientRows = [];
  const allClients = getRegisteredClientsFromSettings_(spreadsheet);

  allClients.forEach(function(clientName) {
    if (isSetupLocalBrand_(clientName)) return;

    const prefix = normalizeSetupClientPrefix_(clientName);
    const catLabel = 'Client: ' + clientName;

    const cGemini = checkClientPropertyConfigured(prefix, ['GEMINI_API_KEY']);
    const cOpenAi = checkClientPropertyConfigured(prefix, ['OPENAI_API_KEY']);
    const cVerifier = checkClientPropertyConfigured(prefix, [
      'ZEROBOUNCE_API_KEY',
      'MILLIONVERIFIER_API_KEY',
      'NEVERBOUNCE_API_KEY',
      'CUSTOM_VERIFIER_API_KEY',
      'VERIFIER_API_KEY',
      'EMAIL_VERIFICATION_API_KEY'
    ]);
    const cSmartlead = checkClientPropertyConfigured(prefix, ['SMARTLEAD_API_KEY']);
    const cInstantly = checkClientPropertyConfigured(prefix, ['INSTANTLY_API_KEY']);
    const cRelayUrl = checkClientPropertyConfigured(prefix, ['GMAIL_RELAY_URL']);
    const cRelaySec = checkClientPropertyConfigured(prefix, ['GMAIL_RELAY_SECRET']);
    const relayStatus = (cRelayUrl && cRelaySec)
      ? protectedMsg
      : ((cRelayUrl || cRelaySec) ? 'Incomplete (Set both _URL and _SECRET in Script Properties)' : '');

    const cRows = [
      [catLabel, prefix + 'GEMINI_API_KEY', cGemini ? protectedMsg : '', clientName + ' Gemini API key (add in Project Settings > Script Properties)'],
      [catLabel, prefix + 'OPENAI_API_KEY', cOpenAi ? protectedMsg : '', clientName + ' OpenAI API key (add in Project Settings > Script Properties)'],
      [catLabel, prefix + 'GOOGLE_SEARCH_GROUNDING', 'TRUE', 'Enable live Search Grounding for ' + clientName + ' (optional Script Property override)'],
      [catLabel, prefix + 'VERIFIER_API_KEY', cVerifier ? protectedMsg : '', clientName + ' verifier key (ZeroBounce / MillionVerifier / NeverBounce in Script Properties)'],
      [catLabel, prefix + 'SMARTLEAD_API_KEY', cSmartlead ? protectedMsg : '', clientName + ' Smartlead API key (add in Script Properties)'],
      [catLabel, prefix + 'INSTANTLY_API_KEY', cInstantly ? protectedMsg : '', clientName + ' Instantly API key (add in Script Properties)'],
      [catLabel, prefix + 'GMAIL_RELAY_URL', relayStatus, clientName + ' Google Workspace Relay (_URL and _SECRET in Script Properties)'],
      [catLabel, prefix + 'DAILY_SPEND_CAP_USD', '10.00', clientName + ' daily API spend limit in USD (editable)'],
      [catLabel, prefix + 'MIN_VERIFIER_CREDITS', '500', clientName + ' min verifier credits before alert (editable)']
    ];

    cRows.forEach(function(r) {
      clientRows.push(r);
      const cleanK = r[1].toUpperCase().replace(/[^A-Z0-9]/g, '');
      categorySyncByKey[cleanK] = r[0];
      descriptionSyncByKey[cleanK] = r[3];
      if (r[1].indexOf('DAILY_SPEND_CAP_USD') < 0 &&
          r[1].indexOf('MIN_VERIFIER_CREDITS') < 0 &&
          r[1].indexOf('GOOGLE_SEARCH_GROUNDING') < 0) {
        secretValueByKey[cleanK] = r[2];
      }
      if (r[1].indexOf('DAILY_SPEND_CAP_USD') >= 0) {
        numericFormatByKey[cleanK] = '0.00';
      } else if (r[1].indexOf('MIN_VERIFIER_CREDITS') >= 0) {
        numericFormatByKey[cleanK] = '0';
      }
    });
  });

  const allTargetRows = localRows.concat(clientRows);

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, allTargetRows.length, 4).setValues(allTargetRows);
  } else {
    const lastR = sheet.getLastRow();
    const existingData = sheet.getRange(2, 1, lastR - 1, 4).getValues();
    const existingKeyList = [];

    for (let i = 0; i < existingData.length; i++) {
      const rowNum = i + 2;
      const cleanKey = String(existingData[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      existingKeyList.push(cleanKey);

      if (categorySyncByKey[cleanKey]) {
        sheet.getRange(rowNum, 1).setValue(categorySyncByKey[cleanKey]);
      }
      if (Object.prototype.hasOwnProperty.call(secretValueByKey, cleanKey)) {
        sheet.getRange(rowNum, 3).setValue(secretValueByKey[cleanKey]);
      }
      if (descriptionSyncByKey[cleanKey]) {
        sheet.getRange(rowNum, 4).setValue(descriptionSyncByKey[cleanKey]);
      }
    }

    allTargetRows.forEach(function(row) {
      const cleanKey = row[1].toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (existingKeyList.indexOf(cleanKey) < 0) {
        sheet.appendRow(row);
        existingKeyList.push(cleanKey);
      }
    });
  }

  // Apply dropdowns & clean number formats across all rows (preserving user-edited cap/credit values)
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
      const rawKey = String(finalKeys[r][0] || '').trim().toUpperCase();
      const cleanK = rawKey.replace(/[^A-Z0-9]/g, '');
      const valCell = sheet.getRange(r + 2, 3);

      if (agentKeyDropdowns[cleanK]) {
        applyDropdownToRange(valCell, agentKeyDropdowns[cleanK], false);
      } else if ((rawKey.indexOf('CLIENT_') === 0 && rawKey.indexOf('GOOGLESEARCHGROUNDING') >= 0) || cleanK.endsWith('GOOGLESEARCHGROUNDING')) {
        applyDropdownToRange(valCell, ['TRUE', 'FALSE'], false);
      }

      if (numericFormatByKey[cleanK]) {
        valCell.setNumberFormat(numericFormatByKey[cleanK]);
      }
    }
  }

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 310);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 440);

  // Filter Agent Settings to show only the active Client / Owner's rows
  filterAgentSettingsByClient_(spreadsheet, activeClientOverride);
}

/**
 * Filters "Agent Settings" Column A (Category) to show only the active Client / Owner's block
 * (or only Growth Target / Local defaults when Growth Target is active).
 * Users can clear the Column A filter in Sheets at any time to compare all clients side by side.
 */
function filterAgentSettingsByClient_(ss, explicitClientName) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const sheet = spreadsheet.getSheetByName('Agent Settings');
  if (!sheet || sheet.getLastRow() <= 1 || sheet.getLastColumn() < 4) return;

  let clientName = String(explicitClientName || '').trim();
  if (!clientName) {
    const cfgSheet = spreadsheet.getSheetByName('Campaign Configuration');
    if (cfgSheet && cfgSheet.getLastRow() > 1) {
      const cfgData = cfgSheet.getDataRange().getValues();
      for (let r = 1; r < cfgData.length; r++) {
        const k = String(cfgData[r][0] || '').trim().toUpperCase();
        if (k === 'DEPARTMENT' && cfgData[r][1]) {
          clientName = String(cfgData[r][1]).trim();
          break;
        }
      }
    }
  }
  if (!clientName) {
    const dashSheet = spreadsheet.getSheetByName('Dashboard');
    if (dashSheet && dashSheet.getLastRow() >= 6) {
      const b6 = String(dashSheet.getRange('B6').getValue() || '').trim();
      if (b6 && b6 !== 'None' && b6 !== 'Standby') clientName = b6;
    }
  }
  clientName = clientName || 'Growth Target';

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let filter = sheet.getFilter();
  if (filter) {
    const fRange = filter.getRange();
    if (fRange.getRow() !== 1 || fRange.getLastRow() < lastRow || fRange.getLastColumn() < lastCol) {
      filter.remove();
      filter = null;
    }
  }
  if (!filter) {
    filter = sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  }

  const catValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const isLocal = isSetupLocalBrand_(clientName);
  const targetClientCat = ('Client: ' + clientName).toLowerCase();

  let hasMatchingRows = false;
  const hiddenValues = [];
  const seenHidden = {};

  for (let i = 0; i < catValues.length; i++) {
    const cat = String(catValues[i][0] || '').trim();
    if (!cat) continue;
    const lowerCat = cat.toLowerCase();
    const isClientCat = lowerCat.indexOf('client:') === 0;

    const shouldShow = isLocal ? !isClientCat : (lowerCat === targetClientCat);
    if (shouldShow) {
      hasMatchingRows = true;
    } else if (!seenHidden[cat]) {
      hiddenValues.push(cat);
      seenHidden[cat] = true;
    }
  }

  if (hasMatchingRows && hiddenValues.length > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(hiddenValues)
      .build();
    filter.setColumnFilterCriteria(1, criteria);
  } else {
    filter.removeColumnFilterCriteria(1);
  }
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