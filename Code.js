/** @OnlyCurrentDoc */
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
    .addItem('✉️ Send Approved Outreach Emails (Using Configured Provider)', 'sendApprovedOutreachEmails')
    .addItem('📝 Create Gmail Drafts (Safe Review)', 'createGmailDraftsForApprovedProspects')
    .addSeparator()
    .addItem('⚙️ Prepare / Verify All Tabs', 'preparePersonaMapperSheets')
    .addItem('❓ Check Spreadsheet Setup', 'showSetupStatus')
    .addToUi();
}