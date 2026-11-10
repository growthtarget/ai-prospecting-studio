/** @OnlyCurrentDoc */

/**
 * Dedicated One-Time Permission Authorizer
 * Run this once inside the Apps Script editor to grant external web and Gmail permissions.
 */
function authorizeStudioPermissions() {
  // 1. Touch External Request scope (UrlFetchApp)
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });

  // 2. Touch Spreadsheet scope
  SpreadsheetApp.getActiveSpreadsheet();

  // 3. Touch Gmail scope
  GmailApp.getDrafts();

  SpreadsheetApp.getUi().alert(
    'Authorization Successful!',
    'All permissions (Web Search, Gemini API, Spreadsheet, and Gmail) are now fully authorized.\n\nYou can now use the sidebar without permission errors.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('AI Prospecting Studio')
    .addItem('Open Control Panel Sidebar', 'showPersonaMapperControlPanel')
    .addSeparator()
    .addItem('🔍 1. Discover Target Companies', 'runCompanyDiscoveryForSelectedCampaign')
    .addItem('👤 2. Run Persona Mapper', 'runPersonaMapperForSelectedCampaign')
    .addSeparator()
    .addItem('📋 Go to Prospect Review Queue', 'navigateToReviewQueueSheet')
    .addItem('✅ Approve Selected Prospect', 'approveSelectedProspect')
    .addItem('✏️ Mark Selected: Needs Editing', 'markSelectedProspectNeedsEditing')
    .addItem('❌ Reject Selected Prospect', 'rejectSelectedProspect')
    .addSeparator()
    .addItem('✉️ Send Approved Outreach Emails', 'sendApprovedOutreachEmails')
    .addItem('📝 Create Gmail Drafts (Safe Review)', 'createGmailDraftsForApprovedProspects')
    .addSeparator()
    .addItem('🔑 Authorize Studio Permissions', 'authorizeStudioPermissions')
    .addItem('⚙️ Prepare / Verify All Tabs', 'preparePersonaMapperSheets')
    .addItem('❓ Check Spreadsheet Setup', 'showSetupStatus')
    .addToUi();
}