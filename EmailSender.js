/**
 * Pluggable Email Sender Engine
 * Supports: Google Workspace (Direct & Drafts), Smartlead, Instantly, and Custom Webhooks.
 */

/**
 * Universal sender: Dispatches outreach based on the campaign's configured provider.
 */
function sendApprovedOutreachEmails() {
  processOutreachQueue_(false);
}

/**
 * Force creates Gmail Drafts regardless of provider setting (Safe testing mode).
 */
function createGmailDraftsForApprovedProspects() {
  processOutreachQueue_(true);
}

/**
 * Core outreach processor.
 */
function processOutreachQueue_(forceDrafts) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensurePersonaMapperSheets_(spreadsheet);

  const sheet = spreadsheet.getSheetByName('Prospect Review Queue');
  if (!sheet) {
    ui.alert('Prospect Review Queue sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    ui.alert('No prospects found in Prospect Review Queue.');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  const campIdCol = headers.indexOf('Campaign ID');
  const emailCol = headers.indexOf('Business Email');
  const subjectCol = headers.indexOf('Draft Subject');
  const bodyCol = headers.indexOf('Draft Email');
  const humanApprCol = headers.indexOf('Human Approval Status');
  const reviewStatusCol = headers.indexOf('Review Status');
  const sendStatusCol = headers.indexOf('Send Status');
  const sentAtCol = headers.indexOf('Sent At');
  const sendErrorCol = headers.indexOf('Send Error');
  const leadNameCol = headers.indexOf('Name');
  const companyCol = headers.indexOf('Company Name');

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const queueToProcess = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const email = String(row[emailCol] || '').trim();
    const humanAppr = String(row[humanApprCol] || '').trim().toLowerCase();
    const reviewStatus = String(row[reviewStatusCol] || '').trim().toLowerCase();
    const sendStatus = String(row[sendStatusCol] || '').trim().toLowerCase();
    const campaignId = String(row[campIdCol] || '').trim();

    const isApproved = (humanAppr === 'approved' || reviewStatus === 'approved for outreach');
    const isAlreadySent = (sendStatus === 'sent' || sendStatus === 'draft created' || sendStatus === 'pushed to sequencer');

    if (email && isApproved && !isAlreadySent) {
      queueToProcess.push({
        rowIndex: i + 2,
        campaignId: campaignId,
        email: email,
        subject: String(row[subjectCol] || 'Introduction'),
        body: String(row[bodyCol] || ''),
        leadName: String(row[leadNameCol] || ''),
        companyName: String(row[companyCol] || '')
      });
    }
  }

  if (queueToProcess.length === 0) {
    ui.alert(
      'No Prospects Ready',
      'No approved prospects are pending outreach.\n\n' +
      'To approve a prospect:\n' +
      '1. Set Human Approval Status to "Approved"\n' +
      '2. Ensure Send Status is empty (not "Sent")\n' +
      '3. Ensure Business Email, Draft Subject, and Draft Email are filled.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Load provider configuration from the first campaign in the batch
  const firstCampaignId = queueToProcess[0].campaignId;
  const settings = getPersonaMapperSettingsSafe_(spreadsheet, firstCampaignId);
  const configuredProvider = forceDrafts ? 'Workspace Gmail (Drafts Only)' : String(settings['Email Sender Provider'] || 'Workspace Gmail (Direct Send)').trim();

  const confirm = ui.alert(
    'Confirm Outreach Execution',
    'Provider: ' + configuredProvider + '\n' +
    'Prospects to process: ' + queueToProcess.length + '\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let successCount = 0;
  let failCount = 0;
  const now = new Date();

  queueToProcess.forEach(function(item) {
    try {
      dispatchOutreach_(item, configuredProvider, settings);
      
      let finalStatus = 'Sent';
      if (configuredProvider.indexOf('Drafts') >= 0) finalStatus = 'Draft Created';
      if (configuredProvider === 'Smartlead' || configuredProvider === 'Instantly' || configuredProvider === 'Custom Webhook') finalStatus = 'Pushed to Sequencer';

      sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue(finalStatus);
      if (sentAtCol >= 0) sheet.getRange(item.rowIndex, sentAtCol + 1).setValue(now);
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue('');
      successCount++;
    } catch (err) {
      failCount++;
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue(err.message);
      if (sendStatusCol >= 0) sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue('Failed');
    }
  });

  ui.alert(
    'Outreach Completed',
    'Processed: ' + successCount + ' successfully.\n' +
    (failCount > 0 ? ('Failed: ' + failCount + '. Check "Send Error" column for details.') : ''),
    ui.ButtonSet.OK
  );
}

/**
 * Universal Dispatcher: Routes to Google Workspace, Smartlead, Instantly, or Webhook.
 */
function dispatchOutreach_(item, provider, settings) {
  const p = provider.toLowerCase();

  // 1. Google Workspace Drafts
  if (p.indexOf('draft') >= 0) {
    GmailApp.createDraft(item.email, item.subject, item.body);
    return;
  }

  // 2. Google Workspace Direct Send
  if (p.indexOf('gmail') >= 0 || p.indexOf('workspace') >= 0) {
    GmailApp.sendEmail(item.email, item.subject, item.body);
    return;
  }

  // 3. Smartlead Integration
  if (p.indexOf('smartlead') >= 0) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('SMARTLEAD_API_KEY') || '';
    const campaignId = String(settings['Email Sender Endpoint'] || settings['Smartlead Campaign ID'] || '').trim();
    if (!apiKey || !campaignId) throw new Error('SMARTLEAD_API_KEY or Smartlead Campaign ID is missing.');

    const url = 'https://server.smartlead.ai/api/v1/campaigns/' + encodeURIComponent(campaignId) + '/leads?api_key=' + apiKey;
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        lead_list: [{
          email: item.email,
          first_name: item.leadName.split(' ')[0] || '',
          last_name: item.leadName.split(' ').slice(1).join(' ') || '',
          company_name: item.companyName,
          custom_fields: { draft_subject: item.subject, draft_body: item.body }
        }]
      }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) throw new Error('Smartlead API error: ' + res.getContentText());
    return;
  }

  // 4. Instantly Integration
  if (p.indexOf('instantly') >= 0) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('INSTANTLY_API_KEY') || '';
    const campaignId = String(settings['Email Sender Endpoint'] || settings['Instantly Campaign ID'] || '').trim();
    if (!apiKey) throw new Error('INSTANTLY_API_KEY is missing.');

    const url = 'https://api.instantly.ai/api/v1/lead/add';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        api_key: apiKey,
        campaign_id: campaignId,
        email: item.email,
        first_name: item.leadName.split(' ')[0] || '',
        last_name: item.leadName.split(' ').slice(1).join(' ') || '',
        company_name: item.companyName,
        custom_variables: { draft_subject: item.subject, draft_body: item.body }
      }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) throw new Error('Instantly API error: ' + res.getContentText());
    return;
  }

  // 5. Custom HTTPS Webhook (Zapier, Make, or CRM)
  if (p.indexOf('webhook') >= 0 || p.indexOf('custom') >= 0) {
    const webhookUrl = String(settings['Email Sender Endpoint'] || '').trim();
    if (!webhookUrl || !/^https:\/\//i.test(webhookUrl)) throw new Error('A valid HTTPS Webhook Endpoint is required.');

    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(item),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) throw new Error('Webhook error: ' + res.getContentText());
    return;
  }

  throw new Error('Unsupported email sender provider: ' + provider);
}

function getPersonaMapperSettingsSafe_(spreadsheet, campaignId) {
  const sheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (!sheet) return {};
  const records = getSheetRecords_(sheet);
  const match = records.find(function(r) {
    return String(r['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();
  });
  return match || {};
}