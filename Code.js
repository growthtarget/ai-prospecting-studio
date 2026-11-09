/** @OnlyCurrentDoc */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('AI Prospecting Studio')
    .addItem('Open Control Panel', 'showPersonaMapperControlPanel')
    .addItem('Check spreadsheet setup', 'showSetupStatus')
    .addItem('Prepare Persona Mapper tabs', 'preparePersonaMapperSheets')
    .addSeparator()
    .addItem('1. Discover companies for campaign', 'runCompanyDiscoveryForSelectedCampaign')
    .addItem('2. Run Persona Mapper for campaign', 'runPersonaMapperForSelectedCampaign')
    .addSeparator()
    .addItem('Approve selected prospect', 'approveSelectedProspect')
    .addItem('Mark selected prospect: Needs Editing', 'markSelectedProspectNeedsEditing')
    .addItem('Mark selected prospect: Needs Further Research', 'markSelectedProspectNeedsFurtherResearch')
    .addItem('Reject selected prospect', 'rejectSelectedProspect')
    .addToUi();
}