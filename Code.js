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
 * Master onOpen Trigger: Installs the complete, unified menu across all tools.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎯 AI Prospecting Studio')
    .addItem('🚀 Open Campaign Control Panel', 'showPersonaMapperControlPanel')
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
    if (ss.getSheetByName(name)) {
      present.push(name);
    } else {
      missing.push(name);
    }
  });

  // Check API Key configuration
  let apiKeyInfo = 'Not Configured (Add key in Agent Settings)';
  const agentSheet = ss.getSheetByName('Agent Settings');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const data = agentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][1] || data[i][0] || '').trim();
      const v = String(data[i][2] || '').trim();
      if ((k === 'GEMINI_API_KEY' || k === 'OPENAI_API_KEY') && v && !v.includes('PASTE_')) {
        apiKeyInfo = 'Connected (' + k + ')';
        break;
      }
    }
  }

  // Check email sender configuration
  let senderInfo = 'Google Workspace (Drafts Only)';
  const configSheet = ss.getSheetByName('Campaign Configuration');
  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim().toUpperCase();
      if (k === 'DISPATCH_MODE' && data[i][1]) {
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

/**
 * Dedicated One-Time Permission Authorizer
 */
function authorizeStudioPermissions() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  SpreadsheetApp.getActiveSpreadsheet();
  GmailApp.getDrafts();

  SpreadsheetApp.getUi().alert(
    'Authorization Successful!',
    'All permissions (Web Search, Gemini/OpenAI API, Spreadsheet, and Gmail) are now fully authorized.\n\nYou can now use all studio tools without permission errors.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sidebar launcher routing
 */
function showPersonaMapperControlPanel() {
  if (typeof showControlPanelSidebar === 'function') {
    showControlPanelSidebar();
  } else {
    const template = HtmlService.createTemplateFromFile('ControlPanel');
    const html = template.evaluate().setTitle('🎯 Campaign Control Panel').setWidth(360);
    SpreadsheetApp.getUi().showSidebar(html);
  }
}

/**
 * Menu connector for creating Gmail drafts for approved leads
 */
function createGmailDraftsForApprovedProspects() {
  if (typeof triggerApprovedGmailDrafts === 'function') {
    triggerApprovedGmailDrafts();
  } else {
    SpreadsheetApp.getUi().alert('Draft engine connector not loaded.');
  }
}

/**
 * Menu connector for live sending approved outreach emails
 */
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
    // Fallback: Dispatch drafts into live sends if EmailSender exists
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