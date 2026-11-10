/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * ISOLATED PILOT SANDBOX RUNNER
 * ============================================================================
 * File: PilotSandboxRunner.gs
 * 
 * SAFETY GUARANTEES:
 * 1. Contains NO onOpen() trigger: Will NOT clash with or alter your live menu.
 * 2. Dedicated Sandbox Sheets: Operates strictly in sheets prefixed with "[PILOT]".
 * 3. Does not call or overwrite your live campaign data, properties, or keys.
 * 
 * HOW TO RUN:
 * In the Apps Script toolbar above, select "runIsolatedDummyPilot" from the 
 * function dropdown, then click "Run".
 * ============================================================================
 */

// Dedicated sandbox sheets to protect your live sheets
const PILOT_TABS = {
  MASTER: '[PILOT] Dashboard',
  COMPANIES: '[PILOT] Discovered Companies',
  PROSPECTS: '[PILOT] Prospects & Outreach',
  REVIEW_QUEUE: '[PILOT] Review Queue'
};

/**
 * Main Sandbox Function: Runs the 5-step pilot flow in isolation.
 */
function runIsolatedDummyPilot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('>>> STARTING ISOLATED DUMMY PILOT RUN <<<');
  
  try {
    // 1. Create or reset isolated sandbox tabs
    setupPilotSandboxSheets(ss);
    
    // 2. Define the dummy campaign parameters
    const pilotCampaign = {
      id: 'CMP-PILOT-SANDBOX-01',
      name: 'FinTech Cloud & SOC 2 Compliance Accelerator',
      targetCountry: 'United States & Canada',
      industry: 'Financial Services / FinTech SaaS',
      companySize: '50 - 250 employees',
      focus: 'Automated SOC 2 readiness & AWS cloud security hardening',
      businessNeeds: 'Streamline audits, eliminate engineering compliance backlog',
      keywords: 'SOC 2, cloud security, AWS architecture, compliance automation',
      targetJobTitles: 'Chief Technology Officer (CTO), VP of Engineering, Head of Information Security / CISO',
      senderName: 'Alex Mercer',
      senderCompany: 'CloudGuard Systems'
    };
    
    // 3. Step 1 - Company Discovery Simulation (Agent 1)
    const discoveredCompanies = getSimulatedCompanyDiscovery();
    writeSandboxCompanies(ss, discoveredCompanies);
    Logger.log('Step 1 Complete: 3 Companies discovered and qualified.');
    
    // 4. Step 2 - Persona Identification Simulation (Agent 2)
    const personas = getSimulatedPersonaMapping();
    Logger.log('Step 2 Complete: 3 Decision-makers identified.');
    
    // 5. Step 3, 4 & 5 - Persona Research, Email Curation & Fallback Execution
    const curatedOutreach = executeSimulatedResearchAndCuration(personas, pilotCampaign);
    writeSandboxProspects(ss, curatedOutreach);
    Logger.log('Steps 3, 4 & 5 Complete: 2 deep-personalized emails + 1 fallback email created.');
    
    // 6. Review Queue Staging
    writeSandboxReviewQueue(ss, curatedOutreach);
    Logger.log('Step 6 Complete: Review Queue staged.');
    
    // 7. Update Sandbox Master Dashboard
    updateSandboxDashboard(ss, pilotCampaign, discoveredCompanies.length, curatedOutreach.length);
    
    // Non-intrusive completion notice (Spreadsheet toast)
    ss.toast(
      'Pilot complete! Check tabs: [PILOT] Discovered Companies, [PILOT] Prospects & Outreach, and [PILOT] Review Queue.',
      '🚀 Pilot Sandbox Success',
      10
    );

    Logger.log('>>> PILOT RUN COMPLETED SUCCESSFULLY WITHOUT INTERFERING WITH LIVE CODE <<<');

  } catch (error) {
    Logger.log('❌ Pilot Execution Error: ' + error.stack);
    ss.toast('Error during pilot: ' + error.message, 'Pilot Error', 10);
  }
}

/**
 * Helper: Sets up sandbox sheets with clear styling
 */
function setupPilotSandboxSheets(ss) {
  const configs = [
    {
      name: PILOT_TABS.MASTER,
      headers: ['Campaign ID', 'Campaign Name', 'Run Status', 'Companies Found', 'Qualified Leads', 'Need Review', 'Last Run']
    },
    {
      name: PILOT_TABS.COMPANIES,
      headers: ['Company ID', 'Company Name', 'Domain', 'Country', 'Employee Size', 'Industry', 'Fit Score', 'Qualification Reasoning', 'Status']
    },
    {
      name: PILOT_TABS.PROSPECTS,
      headers: ['Prospect ID', 'Company Name', 'Full Name', 'Job Title', 'LinkedIn URL', 'Research Signal / Context', 'Generation Mode', 'Email Subject', 'Email Body', 'Verification Status', 'Send Status']
    },
    {
      name: PILOT_TABS.REVIEW_QUEUE,
      headers: ['Approve (TRUE/FALSE)', 'Prospect Name', 'Company', 'Job Title', 'Generation Mode', 'Email Subject', 'Curated Email Body', 'Status']
    }
  ];

  configs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    } else {
      // Clear previous sandbox contents to provide a fresh run
      sheet.clear();
    }
    
    sheet.appendRow(cfg.headers);
    sheet.getRange(1, 1, 1, cfg.headers.length)
      .setFontWeight('bold')
      .setBackground('#202124')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  });
}

/**
 * Step 1 Data: Simulated Company Discovery
 */
function getSimulatedCompanyDiscovery() {
  return [
    {
      id: 'CMP-001',
      name: 'ApexPay Financial',
      domain: 'growthtarget.com',
      country: 'United States',
      size: '120 employees',
      industry: 'FinTech / Payment Gateways',
      fitScore: 96,
      reason: 'B2B payment platform scaling enterprise tier; high cloud security and SOC 2 requirements.',
      status: 'Qualified'
    },
    {
      id: 'CMP-002',
      name: 'NovaWealth Technologies',
      domain: 'growthtarget.com',
      country: 'Canada',
      size: '85 employees',
      industry: 'WealthTech SaaS',
      fitScore: 92,
      reason: 'Canadian wealth management API platform; manages regulated client capital.',
      status: 'Qualified'
    },
    {
      id: 'CMP-003',
      name: 'ClearVault Systems',
      domain: 'growthtarget.com',
      country: 'United States',
      size: '190 employees',
      industry: 'Embedded Banking',
      fitScore: 89,
      reason: 'Embedded banking infrastructure; heavy regulatory scrutiny and developer compliance load.',
      status: 'Qualified'
    }
  ];
}

/**
 * Writes Discovered Companies into the Sandbox Tab
 */
function writeSandboxCompanies(ss, companies) {
  const sheet = ss.getSheetByName(PILOT_TABS.COMPANIES);
  const rows = companies.map(c => [
    c.id,
    c.name,
    c.domain,
    c.country,
    c.size,
    c.industry,
    c.fitScore,
    c.reason,
    c.status
  ]);
  
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, rows[0].length);
}

/**
 * Step 2 Data: Simulated Persona Mapping
 */
function getSimulatedPersonaMapping() {
  return [
    {
      prospectId: 'PRSP-001',
      companyId: 'CMP-001',
      companyName: 'ApexPay Financial',
      fullName: 'Marcus Vance',
      jobTitle: 'Chief Technology Officer (CTO)',
      linkedinUrl: 'https://linkedin.com/in/marcus-vance-pilot',
      hasDigitalFootprint: true
    },
    {
      prospectId: 'PRSP-002',
      companyId: 'CMP-002',
      companyName: 'NovaWealth Technologies',
      fullName: 'Elena Rostova',
      jobTitle: 'VP of Engineering',
      linkedinUrl: 'https://linkedin.com/in/elena-rostova-pilot',
      hasDigitalFootprint: true
    },
    {
      prospectId: 'PRSP-003',
      companyId: 'CMP-003',
      companyName: 'ClearVault Systems',
      fullName: 'David Chen',
      jobTitle: 'Head of Information Security',
      linkedinUrl: 'https://linkedin.com/in/david-chen-pilot',
      hasDigitalFootprint: false // Triggers Step 5 Fallback
    }
  ];
}

/**
 * Steps 3, 4 & 5: Persona Research, Email Curation & Fallback Execution
 */
function executeSimulatedResearchAndCuration(personas, campaign) {
  const results = [];

  personas.forEach(p => {
    if (p.hasDigitalFootprint && p.fullName === 'Marcus Vance') {
      // Deep Persona Personalization
      const signal = 'LinkedIn Post (2 weeks ago): Discussed challenges migrating microservices to AWS and developer bandwidth lost to Q4 SOC 2 prep.';
      const subject = "ApexPay's microservices migration + SOC 2 prep";
      const body = "Hi Marcus,\n\n" +
        "Saw your note recently regarding migrating ApexPay's microservices to AWS and the bandwidth Q4 SOC 2 prep is pulling from your sprint cycles. Scaling payment infrastructure shouldn't require senior developers to spend dozens of hours gathering manual audit evidence.\n\n" +
        "We help mid-market FinTechs automate continuous SOC 2 and AWS security controls directly inside CI/CD pipelines, cutting audit prep time by roughly 70% without slowing down feature releases.\n\n" +
        "Open to a brief 7-minute intro this Thursday to see how we streamline this for teams like ApexPay?\n\n" +
        "Best regards,\n" +
        campaign.senderName + "\n" +
        campaign.senderCompany;

      results.push({
        ...p,
        signal: signal,
        mode: 'Deep Personalized (Post Match)',
        subject: subject,
        body: body,
        verificationStatus: 'simulated_valid',
        sendStatus: 'Draft Ready'
      });

    } else if (p.hasDigitalFootprint && p.fullName === 'Elena Rostova') {
      // Deep Persona Personalization
      const signal = 'Keynote at Canadian FinTech Summit: Discussed balancing engineering sprint velocity with Tier-1 banking security requirements.';
      const subject = "Sprint velocity vs. banking compliance at NovaWealth";
      const body = "Hi Elena,\n\n" +
        "Really appreciated your talk at the Canadian FinTech Summit on balancing sprint velocity with banking compliance—especially your point on how regulatory bottlenecks slow down core API releases.\n\n" +
        "We work specifically with Canadian wealth tech engineering leaders to bake compliance guardrails and cloud security controls directly into the developer workflow. This lets your team ship updates at high velocity while staying completely audit-ready for banking partners.\n\n" +
        "Would you be against exploring how we help FinTech engineering teams eliminate compliance drag next Tuesday?\n\n" +
        "Best regards,\n" +
        campaign.senderName + "\n" +
        campaign.senderCompany;

      results.push({
        ...p,
        signal: signal,
        mode: 'Deep Personalized (Article/Talk Match)',
        subject: subject,
        body: body,
        verificationStatus: 'simulated_valid',
        sendStatus: 'Draft Ready'
      });

    } else {
      // Step 5: Fallback Mechanism (Triggered when digital footprint is sparse/absent)
      const signal = 'No public posts or publications found. Sparse digital footprint. [Fallback Activated]';
      const subject = "Security controls for ClearVault's embedded banking stack";
      const body = "Hi David,\n\n" +
        "Given ClearVault's expansion across embedded banking, managing continuous cloud security controls across multiple banking integrations is likely a heavy lift for your security team.\n\n" +
        "We help InfoSec leaders at scaling FinTechs automate cloud compliance and SOC 2 evidence collection across their infrastructure, ensuring every integration stays compliant without burdening internal developers.\n\n" +
        "Would you be open to a quick 5-minute sync next week to see how we support embedded banking platforms?\n\n" +
        "Best regards,\n" +
        campaign.senderName + "\n" +
        campaign.senderCompany;

      results.push({
        ...p,
        signal: signal,
        mode: 'Company/Role Fallback (Triggered)',
        subject: subject,
        body: body,
        verificationStatus: 'simulated_valid',
        sendStatus: 'Draft Ready'
      });
    }
  });

  return results;
}

/**
 * Writes Prospects & Generated Outreach into the Sandbox Tab
 */
function writeSandboxProspects(ss, list) {
  const sheet = ss.getSheetByName(PILOT_TABS.PROSPECTS);
  const rows = list.map(item => [
    item.prospectId,
    item.companyName,
    item.fullName,
    item.jobTitle,
    item.linkedinUrl,
    item.signal,
    item.mode,
    item.subject,
    item.body,
    item.verificationStatus,
    item.sendStatus
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, rows[0].length);
  sheet.getRange(2, 9, rows.length, 1).setWrap(true); // Wrap body column
}

/**
 * Writes Curated Leads to the Sandbox Review Queue
 */
function writeSandboxReviewQueue(ss, list) {
  const sheet = ss.getSheetByName(PILOT_TABS.REVIEW_QUEUE);
  const rows = list.map(item => [
    false, // Checkbox for manual approval
    item.fullName,
    item.companyName,
    item.jobTitle,
    item.mode,
    item.subject,
    item.body,
    'Needs Human Review'
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  
  // Insert checkboxes in Column A
  const checkboxRange = sheet.getRange(2, 1, rows.length, 1);
  checkboxRange.insertCheckboxes();

  sheet.autoResizeColumns(1, 6);
  sheet.getRange(2, 7, rows.length, 1).setWrap(true);
}

/**
 * Updates the Sandbox Dashboard
 */
function updateSandboxDashboard(ss, campaign, companiesCount, leadsCount) {
  const sheet = ss.getSheetByName(PILOT_TABS.MASTER);
  const row = [
    campaign.id,
    campaign.name,
    'Completed (Sandbox Pilot)',
    companiesCount,
    leadsCount,
    leadsCount, // Staged for review
    new Date().toLocaleString()
  ];
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  sheet.autoResizeColumns(1, row.length);
}

/**
 * Utility: Completely deletes the sandbox pilot sheets whenever you are finished
 */
function deletePilotSandboxSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = [
    PILOT_TABS.MASTER,
    PILOT_TABS.COMPANIES,
    PILOT_TABS.PROSPECTS,
    PILOT_TABS.REVIEW_QUEUE
  ];

  tabs.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  });

  ss.toast('Sandbox sheets successfully removed.', 'Pilot Cleaned', 5);
}