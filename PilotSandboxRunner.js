/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * ISOLATED PILOT SANDBOX RUNNER WITH GMAIL DRAFTS & SUPPRESSION ENGINE
 * ============================================================================
 * File: PilotSandboxRunner.gs
 * 
 * SAFETY GUARANTEES:
 * 1. Contains NO onOpen() trigger: Will NOT clash with or alter your live menu.
 * 2. Dedicated Sandbox Sheets: Operates strictly in sheets prefixed with "[PILOT]".
 * 3. Safe Gmail Draft Creation: Creates drafts in your Gmail "Drafts" folder;
 *    it NEVER clicks send or contacts actual prospects.
 * 4. Suppression Guard: Blocks internal domains, competitors, role emails,
 *    and any record listed on the [PILOT] Suppression List tab.
 * 
 * HOW TO RUN:
 * 1. Open your Apps Script editor.
 * 2. Select "runIsolatedDummyPilot" from the function dropdown at the top.
 * 3. Click "Run". (Grant authorization for Gmail Compose if prompted).
 * 4. Open your Gmail Drafts folder to inspect the created drafts!
 * ============================================================================
 */

// Dedicated sandbox sheets to protect your live production sheets
const PILOT_TABS = {
  MASTER: '[PILOT] Dashboard',
  COMPANIES: '[PILOT] Discovered Companies',
  PROSPECTS: '[PILOT] Prospects & Outreach',
  REVIEW_QUEUE: '[PILOT] Review Queue',
  SUPPRESSION: '[PILOT] Suppression List'
};

/**
 * Main Sandbox Function: Runs the end-to-end pilot flow in complete isolation.
 */
function runIsolatedDummyPilot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('====================================================');
  Logger.log('>>> STARTING ISOLATED DUMMY PILOT RUN (END-TO-END) <<<');
  Logger.log('====================================================');
  
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
    Logger.log('Step 1 Complete: 4 Companies discovered and qualified.');
    
    // 4. Step 2 - Persona Identification Simulation (Agent 2)
    const personas = getSimulatedPersonaMapping();
    Logger.log('Step 2 Complete: 4 Prospects identified with contact emails.');
    
    // 5. Steps 3, 4 & 5 - Persona Research, Email Curation & Fallback Execution
    const curatedOutreach = executeSimulatedResearchAndCuration(personas, pilotCampaign);
    Logger.log('Steps 3, 4 & 5 Complete: Curated personalized outreach copy generated.');
    
    // 6. Security Check & Suppression List Screening
    const suppressionData = getSuppressionRules(ss);
    const auditedOutreach = runSecurityAndSuppressionAudit(curatedOutreach, suppressionData);
    Logger.log('Step 6 Complete: Security & suppression checks evaluated.');
    
    // 7. Auto-Approve & Create Gmail Drafts for Approved Leads
    const finalProcessedOutreach = processApprovalsAndCreateGmailDrafts(auditedOutreach, pilotCampaign);
    Logger.log('Step 7 Complete: Approved leads converted to Gmail drafts.');
    
    // 8. Write Results to Sheets
    writeSandboxProspects(ss, finalProcessedOutreach);
    writeSandboxReviewQueue(ss, finalProcessedOutreach);
    
    // 9. Update Sandbox Master Dashboard
    const draftsCount = finalProcessedOutreach.filter(p => p.sendStatus === 'Gmail Draft Created').length;
    const blockedCount = finalProcessedOutreach.filter(p => p.securityStatus !== 'Passed').length;
    updateSandboxDashboard(ss, pilotCampaign, discoveredCompanies.length, finalProcessedOutreach.length, draftsCount, blockedCount);
    
    // Non-intrusive completion notice (Spreadsheet toast)
    ss.toast(
      'Pilot complete! ' + draftsCount + ' Gmail drafts created. Check your Gmail Drafts & [PILOT] tabs.',
      '🚀 Pilot Sandbox Success',
      12
    );

    Logger.log('====================================================');
    Logger.log('>>> PILOT RUN COMPLETED: ' + draftsCount + ' DRAFTS CREATED IN GMAIL <<<');
    Logger.log('====================================================');

  } catch (error) {
    Logger.log('❌ Pilot Execution Error: ' + error.stack);
    ss.toast('Error during pilot: ' + error.message, 'Pilot Error', 10);
  }
}

/**
 * Sets up sandbox sheets with clear styling and default headers.
 */
function setupPilotSandboxSheets(ss) {
  const configs = [
    {
      name: PILOT_TABS.MASTER,
      headers: ['Campaign ID', 'Campaign Name', 'Run Status', 'Companies Found', 'Total Leads', 'Drafts Created in Gmail', 'Blocked by Suppression', 'Last Run']
    },
    {
      name: PILOT_TABS.COMPANIES,
      headers: ['Company ID', 'Company Name', 'Domain', 'Country', 'Employee Size', 'Industry', 'Fit Score', 'Qualification Reasoning', 'Status']
    },
    {
      name: PILOT_TABS.PROSPECTS,
      headers: ['Prospect ID', 'Company Name', 'Full Name', 'Email Address', 'Job Title', 'Fit Score', 'Security / Suppression Status', 'Approval Status', 'Send Status', 'Gmail Draft ID', 'Subject', 'Email Body']
    },
    {
      name: PILOT_TABS.REVIEW_QUEUE,
      headers: ['Approve (TRUE/FALSE)', 'Prospect Name', 'Email Address', 'Company', 'Job Title', 'Security Status', 'Approval Status', 'Draft Status', 'Email Subject', 'Curated Email Body']
    },
    {
      name: PILOT_TABS.SUPPRESSION,
      headers: ['Suppression Type', 'Blocked Value / Pattern', 'Reason', 'Date Added']
    }
  ];

  configs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    } else {
      sheet.clear();
    }
    
    sheet.appendRow(cfg.headers);
    sheet.getRange(1, 1, 1, cfg.headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  });

  // Seed default demonstration suppression data
  seedSandboxSuppressionList(ss);
}

/**
 * Seeds the [PILOT] Suppression List with realistic sample rules
 */
function seedSandboxSuppressionList(ss) {
  const sheet = ss.getSheetByName(PILOT_TABS.SUPPRESSION);
  const sampleSuppressionRows = [
    ['Domain', 'competitorcloud.com', 'Competitor organization - Do not contact', '2026-01-15'],
    ['Domain', 'blockeddomain.com', 'High bounce risk domain', '2026-02-01'],
    ['Email', 'optout-user@example.com', 'Explicit Unsubscribe request', '2026-02-10'],
    ['Email', 'ceo@growthtarget.com', 'Internal team member / exclusion', '2026-02-15']
  ];
  sheet.getRange(2, 1, sampleSuppressionRows.length, sampleSuppressionRows[0].length).setValues(sampleSuppressionRows);
  sheet.autoResizeColumns(1, 4);
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
    },
    {
      id: 'CMP-004',
      name: 'Competitor Cloud Solutions',
      domain: 'competitorcloud.com',
      country: 'United States',
      size: '220 employees',
      industry: 'Cloud Security / Direct Competitor',
      fitScore: 35,
      reason: 'Direct industry competitor identified during discovery scan.',
      status: 'Disqualified'
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
 * Step 2 Data: Simulated Persona Mapping with Email Addresses
 */
function getSimulatedPersonaMapping() {
  return [
    {
      prospectId: 'PRSP-001',
      companyId: 'CMP-001',
      companyName: 'ApexPay Financial',
      domain: 'growthtarget.com',
      fullName: 'Marcus Vance',
      email: 'marcus.vance@growthtarget.com',
      jobTitle: 'Chief Technology Officer (CTO)',
      linkedinUrl: 'https://linkedin.com/in/marcus-vance-pilot',
      fitScore: 96,
      hasDigitalFootprint: true
    },
    {
      prospectId: 'PRSP-002',
      companyId: 'CMP-002',
      companyName: 'NovaWealth Technologies',
      domain: 'growthtarget.com',
      fullName: 'Elena Rostova',
      email: 'elena.rostova@growthtarget.com',
      jobTitle: 'VP of Engineering',
      linkedinUrl: 'https://linkedin.com/in/elena-rostova-pilot',
      fitScore: 92,
      hasDigitalFootprint: true
    },
    {
      prospectId: 'PRSP-003',
      companyId: 'CMP-003',
      companyName: 'ClearVault Systems',
      domain: 'growthtarget.com',
      fullName: 'David Chen',
      email: 'david.chen@growthtarget.com',
      jobTitle: 'Head of Information Security',
      linkedinUrl: 'https://linkedin.com/in/david-chen-pilot',
      fitScore: 89,
      hasDigitalFootprint: false // Triggers Step 5 Fallback
    },
    {
      prospectId: 'PRSP-004',
      companyId: 'CMP-004',
      companyName: 'Competitor Cloud Solutions',
      domain: 'competitorcloud.com',
      fullName: 'Alex Stone',
      email: 'alex.stone@competitorcloud.com',
      jobTitle: 'Chief Product Officer',
      linkedinUrl: 'https://linkedin.com/in/alex-stone-pilot',
      fitScore: 40,
      hasDigitalFootprint: true // Demonstrates suppression security rejection
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
        verificationStatus: 'valid'
      });

    } else if (p.hasDigitalFootprint && p.fullName === 'Elena Rostova') {
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
        verificationStatus: 'valid'
      });

    } else if (p.fullName === 'David Chen') {
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
        verificationStatus: 'valid'
      });

    } else {
      // Alex Stone (Competitor Lead)
      const signal = 'Public profile match. Competitor domain flagged.';
      const subject = "Cloud compliance automation discussion";
      const body = "Hi Alex,\n\nReaching out regarding cloud compliance automation.\n\nBest,\n" + campaign.senderName;

      results.push({
        ...p,
        signal: signal,
        mode: 'Standard Outreach',
        subject: subject,
        body: body,
        verificationStatus: 'valid'
      });
    }
  });

  return results;
}

/**
 * Reads suppression rules from the [PILOT] Suppression List tab
 */
function getSuppressionRules(ss) {
  const sheet = ss.getSheetByName(PILOT_TABS.SUPPRESSION);
  const data = sheet.getDataRange().getValues();
  const rules = {
    domains: [],
    emails: []
  };

  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0]).trim().toLowerCase();
    const value = String(data[i][1]).trim().toLowerCase();
    if (type === 'domain') {
      rules.domains.push(value);
    } else if (type === 'email') {
      rules.emails.push(value);
    }
  }

  return rules;
}

/**
 * Security & Compliance Engine:
 * Evaluates syntax, role-based mailboxes, and suppression list records.
 */
function runSecurityAndSuppressionAudit(prospects, suppressionRules) {
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@', 'team@'];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return prospects.map(item => {
    const email = (item.email || '').trim().toLowerCase();
    const domain = (item.domain || '').trim().toLowerCase();

    // 1. Email Syntax Validation
    if (!email || !emailRegex.test(email)) {
      item.securityStatus = 'Blocked: Invalid Email Syntax';
      return item;
    }

    // 2. Role-Based Mailbox Screening
    const isRoleBased = rolePrefixes.some(prefix => email.startsWith(prefix));
    if (isRoleBased) {
      item.securityStatus = 'Blocked: Role-Based Mailbox';
      return item;
    }

    // 3. Suppression List Matching (Email level)
    if (suppressionRules.emails.includes(email)) {
      item.securityStatus = 'Blocked: Email in Suppression List';
      return item;
    }

    // 4. Suppression List Matching (Domain level)
    if (suppressionRules.domains.includes(domain)) {
      item.securityStatus = 'Blocked: Domain in Suppression List';
      return item;
    }

    // 5. Verification Status Check
    if (item.verificationStatus !== 'valid') {
      item.securityStatus = 'Blocked: Verification Status Invalid';
      return item;
    }

    // Passed all security and suppression controls
    item.securityStatus = 'Passed';
    return item;
  });
}

/**
 * Approves qualified, clean leads and creates drafts directly in Gmail
 */
function processApprovalsAndCreateGmailDrafts(prospects, campaign) {
  return prospects.map(item => {
    // Lead Qualification & Approval Rule:
    // Requires Security Passed AND Fit Score >= 85
    if (item.securityStatus === 'Passed' && (item.fitScore || 0) >= 85) {
      item.approvalStatus = 'Approved (Auto-Pilot)';
      
      try {
        // Create draft directly in connected Gmail account
        const draft = GmailApp.createDraft(
          item.email,
          item.subject,
          item.body,
          {
            name: campaign.senderName
          }
        );
        
        item.gmailDraftId = draft.getId();
        item.sendStatus = 'Gmail Draft Created';
        Logger.log('Created Gmail Draft for ' + item.fullName + ' (' + item.email + ') | Draft ID: ' + item.gmailDraftId);

      } catch (err) {
        Logger.log('Draft Creation Error for ' + item.email + ': ' + err.message);
        item.gmailDraftId = 'Error';
        item.sendStatus = 'Draft Failed: ' + err.message;
      }

    } else if (item.securityStatus !== 'Passed') {
      item.approvalStatus = 'Rejected (Security Suppression)';
      item.gmailDraftId = 'N/A';
      item.sendStatus = 'Suppressed';

    } else {
      item.approvalStatus = 'Needs Human Review (Score < 85)';
      item.gmailDraftId = 'Pending';
      item.sendStatus = 'Staged in Review Queue';
    }

    return item;
  });
}

/**
 * Writes Prospects & Outreach Details to the Sandbox Tab
 */
function writeSandboxProspects(ss, list) {
  const sheet = ss.getSheetByName(PILOT_TABS.PROSPECTS);
  const rows = list.map(item => [
    item.prospectId,
    item.companyName,
    item.fullName,
    item.email,
    item.jobTitle,
    item.fitScore,
    item.securityStatus,
    item.approvalStatus,
    item.sendStatus,
    item.gmailDraftId,
    item.subject,
    item.body
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, 10);
  sheet.getRange(2, 12, rows.length, 1).setWrap(true);
}

/**
 * Writes Leads and their Draft status to the Sandbox Review Queue
 */
function writeSandboxReviewQueue(ss, list) {
  const sheet = ss.getSheetByName(PILOT_TABS.REVIEW_QUEUE);
  const rows = list.map(item => [
    item.approvalStatus.startsWith('Approved'),
    item.fullName,
    item.email,
    item.companyName,
    item.jobTitle,
    item.securityStatus,
    item.approvalStatus,
    item.sendStatus,
    item.subject,
    item.body
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  
  // Column A checkbox format
  const checkboxRange = sheet.getRange(2, 1, rows.length, 1);
  checkboxRange.insertCheckboxes();

  sheet.autoResizeColumns(1, 8);
  sheet.getRange(2, 10, rows.length, 1).setWrap(true);
}

/**
 * Updates the Sandbox Master Dashboard
 */
function updateSandboxDashboard(ss, campaign, companiesCount, leadsCount, draftsCount, blockedCount) {
  const sheet = ss.getSheetByName(PILOT_TABS.MASTER);
  const row = [
    campaign.id,
    campaign.name,
    'Pilot Completed (Drafts Generated)',
    companiesCount,
    leadsCount,
    draftsCount,
    blockedCount,
    new Date().toLocaleString()
  ];
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  sheet.autoResizeColumns(1, row.length);
}

/**
 * Optional Helper: Lets you manually approve queued leads via Checkboxes
 * and create Gmail drafts only for rows marked TRUE.
 */
function approveAndCreateDraftsFromSandboxQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(PILOT_TABS.REVIEW_QUEUE);

  if (!queueSheet) {
    ss.toast('Sandbox Review Queue not found. Run "runIsolatedDummyPilot" first.', 'Error', 5);
    return;
  }

  const data = queueSheet.getDataRange().getValues();
  let createdCount = 0;

  for (let i = 1; i < data.length; i++) {
    const isApproved = data[i][0];
    const name = data[i][1];
    const email = data[i][2];
    const draftStatus = data[i][7];
    const subject = data[i][8];
    const body = data[i][9];

    // Only process rows checked TRUE that haven't been created yet
    if (isApproved === true && draftStatus !== 'Gmail Draft Created') {
      try {
        GmailApp.createDraft(email, subject, body);
        queueSheet.getRange(i + 1, 8).setValue('Gmail Draft Created');
        createdCount++;
        Logger.log('Manually Approved & Created Gmail Draft for: ' + name + ' (' + email + ')');
      } catch (err) {
        queueSheet.getRange(i + 1, 8).setValue('Draft Error: ' + err.message);
      }
    }
  }

  ss.toast('Processed manual queue! Created ' + createdCount + ' new drafts.', 'Drafts Updated', 6);
}

/**
 * Utility: Completely deletes all sandbox pilot sheets whenever you want a clean slate.
 */
function deletePilotSandboxSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = [
    PILOT_TABS.MASTER,
    PILOT_TABS.COMPANIES,
    PILOT_TABS.PROSPECTS,
    PILOT_TABS.REVIEW_QUEUE,
    PILOT_TABS.SUPPRESSION
  ];

  tabs.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  });

  ss.toast('All [PILOT] sandbox sheets successfully deleted.', 'Sandbox Reset', 5);
}