/** @OnlyCurrentDoc */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('AI Prospecting Studio')
    .addItem('Check spreadsheet setup','showSetupStatus')
    .addSeparator()
    .addItem('Persona Mapper Control Panel','showPersonaMapperControlPanel')
    .addItem('Prepare Persona Mapper tabs','preparePersonaMapperSheets')
    .addItem('Run Persona Mapper for selected campaign','runPersonaMapperForSelectedCampaign')
    .addSeparator()
    .addItem('Approve selected prospect','approveSelectedProspect')
    .addItem('Mark selected prospect: Needs Editing','markSelectedProspectNeedsEditing')
    .addItem('Mark selected prospect: Needs Further Research','markSelectedProspectNeedsFurtherResearch')
    .addItem('Reject selected prospect','rejectSelectedProspect')
    .addToUi();
}