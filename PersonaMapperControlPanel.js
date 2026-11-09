/**
 * Opens the modern Prospecting Control Panel in the sidebar.
 */
function showPersonaMapperControlPanel() {
  const html = HtmlService.createHtmlOutputFromFile('ControlPanel')
    .setTitle('AI Prospecting Studio')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns available campaigns for the sidebar dropdown.
 */
function getAvailableCampaignsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Campaigns');
  if (!sheet) return [];
  const records = getSheetRecords_(sheet);
  return records.map(function(r) {
    return {
      id: String(r['Campaign ID'] || '').trim(),
      name: String(r['Campaign name'] || '').trim()
    };
  }).filter(function(c) { return Boolean(c.id); });
}

/**
 * Returns targeting criteria for the selected campaign.
 */
function getCampaignTargetingSummary(campaignId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Campaign Configuration');
  if (!sheet) return null;
  const records = getSheetRecords_(sheet);
  const match = records.find(function(r) {
    return String(r['Campaign ID'] || '').trim() === String(campaignId).trim();
  });
  if (!match) return null;
  return {
    country: match['Countries/regions'] || '',
    industry: match['Target industries'] || '',
    size: match['Company size'] || ''
  };
}

/**
 * Jumps to the Prospect Review Queue tab.
 */
function navigateToReviewQueueSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Prospect Review Queue');
  if (sheet) spreadsheet.setActiveSheet(sheet);
}