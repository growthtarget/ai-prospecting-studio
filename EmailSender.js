/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * PLUGGABLE EMAIL SENDER & OUTREACH DISPATCH ENGINE
 * ============================================================================
 * File: EmailSender.gs / EmailSender.js
 * 
 * CAPABILITIES:
 * 1. Multi-Provider Dispatch: Google Workspace (Drafts & Direct), Smartlead, Instantly, Custom Webhooks.
 * 2. Campaign Isolation: Resolves sender profile and provider dynamically per prospect Campaign ID.
 * 3. Sender Identity Stamping: Attaches configured Sender Display Name to Gmail drafts and emails.
 * 4. Multi-Source Credential Resolution: Reads API keys from Script Properties, User Properties,
 *    or Tab 10 (Agent Settings: Category | Setting Key | Value | Status / Description).
 * 5. Headless Safe: Gracefully executes in both UI dialog and silent background/sidebar modes.
 * ============================================================================
 */

/**
 * Universal sender: Dispatches outreach based on each campaign's configured provider.
 */
function sendApprovedOutreachEmails() {
  processOutreachQueue_(false);
}

/**
 * Force creates Gmail Drafts regardless of provider setting (Safe testing & review mode).
 */
function createGmailDraftsForApprovedProspects() {
  processOutreachQueue_(true);
}

/**
 * Core outreach processor.
 * Resolves settings per prospect row, ensuring complete campaign isolation.
 * 
 * @param {boolean} forceDrafts When true, routes all approved prospects to Gmail Drafts
 */
function processOutreachQueue_(forceDrafts) {
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Running in headless or background context
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) {
    Logger.log('Error: Active spreadsheet not accessible.');
    return;
  }

  const sheet = spreadsheet.getSheetByName('Prospect Review Queue');
  if (!sheet) {
    if (ui) ui.alert('Prospect Review Queue sheet not found.');
    else Logger.log('Prospect Review Queue sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    if (ui) ui.alert('No prospects found in Prospect Review Queue.');
    else Logger.log('No prospects found in Prospect Review Queue.');
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

    const isApproved = (humanAppr.startsWith('approved') || reviewStatus === 'approved for outreach');
    const isAlreadySent = (
      sendStatus === 'sent' || 
      sendStatus === 'draft created' || 
      sendStatus === 'draft created in gmail' || 
      sendStatus === 'pushed to sequencer'
    );

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
    const msg = 'No approved prospects are pending outreach.\n\n' +
      'To approve a prospect:\n' +
      '1. Set Human Approval Status to "Approved"\n' +
      '2. Ensure Send Status is empty (not "Sent" or "Draft Created")\n' +
      '3. Ensure Business Email, Draft Subject, and Draft Email are filled.';
    if (ui) ui.alert('No Prospects Ready', msg, ui.ButtonSet.OK);
    else Logger.log(msg);
    return;
  }

  // Pre-fetch unique campaign IDs in the batch
  const uniqueCampaigns = Array.from(new Set(queueToProcess.map(function(item) {
    return item.campaignId || 'DEFAULT';
  })));

  if (ui) {
    const modeLabel = forceDrafts ? 'Safe Mode: Workspace Gmail (Drafts Only)' : 'Configured Providers';
    const confirm = ui.alert(
      'Confirm Outreach Execution',
      'Mode: ' + modeLabel + '\n' +
      'Prospects to process: ' + queueToProcess.length + '\n' +
      'Target Campaign(s): ' + uniqueCampaigns.join(', ') + '\n\n' +
      'Do you want to proceed?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  // In-memory cache to resolve campaign settings once per campaign ID
  const campaignSettingsCache = {};
  let successCount = 0;
  let failCount = 0;
  const now = new Date();
  const nowStr = now.toLocaleString();

  queueToProcess.forEach(function(item) {
    const cId = item.campaignId || 'DEFAULT';
    if (!campaignSettingsCache[cId]) {
      campaignSettingsCache[cId] = getCampaignSenderSettings_(spreadsheet, cId);
    }
    const campaignConfig = campaignSettingsCache[cId];

    // Determine provider for this specific prospect row
    const rawProvider = forceDrafts
      ? 'Workspace Gmail (Drafts Only)'
      : String(campaignConfig.emailSenderProvider || 'Workspace Gmail (Direct Send)').trim();

    try {
      const dispatchResult = dispatchOutreach_(item, rawProvider, campaignConfig, spreadsheet);

      // Write updated status matching SheetSetup.js canonical validation values
      let finalStatus = 'Sent';
      if (rawProvider.toLowerCase().indexOf('draft') >= 0) {
        finalStatus = 'Draft Created in Gmail';
      } else if (
        rawProvider.toLowerCase().indexOf('smartlead') >= 0 ||
        rawProvider.toLowerCase().indexOf('instantly') >= 0 ||
        rawProvider.toLowerCase().indexOf('webhook') >= 0 ||
        rawProvider.toLowerCase().indexOf('custom') >= 0
      ) {
        finalStatus = 'Pushed to Sequencer';
      }

      sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue(finalStatus);
      if (sentAtCol >= 0) sheet.getRange(item.rowIndex, sentAtCol + 1).setValue(nowStr);
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue('');
      successCount++;
    } catch (err) {
      failCount++;
      Logger.log('Dispatch failure for row ' + item.rowIndex + ' (' + item.email + '): ' + err.message);
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue(err.message);
      if (sendStatusCol >= 0) sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue('Draft Failed');
    }
  });

  SpreadsheetApp.flush();

  const completionMsg = 'Outreach Execution Completed!\n\n' +
    '• Successfully Processed: ' + successCount + '\n' +
    (failCount > 0 ? ('• Failed: ' + failCount + ' (Check "Send Error" column for details)') : '• All approved records processed cleanly.');

  if (ui) {
    ui.alert('Outreach Completed', completionMsg, ui.ButtonSet.OK);
  } else {
    Logger.log(completionMsg);
  }
}

/**
 * Universal Dispatcher: Routes to Google Workspace, Smartlead, Instantly, or Custom Webhook.
 * Stamps the configured Sender Display Name onto drafts and direct emails.
 */
function dispatchOutreach_(item, provider, config, spreadsheet) {
  const p = String(provider || '').toLowerCase();
  const senderName = String(config.senderName || '').trim();

  // 1. Google Workspace Drafts (Creates real draft with sender display name)
  if (p.indexOf('draft') >= 0) {
    const draftOptions = {};
    if (senderName) draftOptions.name = senderName;

    GmailApp.createDraft(item.email, item.subject, item.body, draftOptions);
    return { status: 'draft_created' };
  }

  // 2. Google Workspace Direct Send (Dispatches direct email with sender display name)
  if (p.indexOf('gmail') >= 0 || p.indexOf('workspace') >= 0 || p.indexOf('direct') >= 0) {
    const sendOptions = {};
    if (senderName) sendOptions.name = senderName;

    GmailApp.sendEmail(item.email, item.subject, item.body, sendOptions);
    return { status: 'sent' };
  }

  // 3. Smartlead Sequencer Integration
  if (p.indexOf('smartlead') >= 0) {
    const apiKey = getSenderApiKey_('SMARTLEAD_API_KEY', spreadsheet);
    const campaignId = String(config.emailSenderEndpoint || config['Smartlead Campaign ID'] || '').trim();
    if (!apiKey) throw new Error('SMARTLEAD_API_KEY not found in Project Settings or Agent Settings.');
    if (!campaignId) throw new Error('Smartlead Campaign ID is missing in Email Sender Endpoint.');

    const url = 'https://server.smartlead.ai/api/v1/campaigns/' + encodeURIComponent(campaignId) + '/leads?api_key=' + apiKey;
    const nameParts = (item.leadName || '').trim().split(' ');
    const firstName = nameParts[0] || '';
    const lastName = nameParts.slice(1).join(' ') || '';

    const payload = {
      lead_list: [{
        email: item.email,
        first_name: firstName,
        last_name: lastName,
        company_name: item.companyName,
        custom_fields: {
          draft_subject: item.subject,
          draft_body: item.body,
          sender_name: senderName
        }
      }]
    };

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() >= 300) {
      throw new Error('Smartlead API error (' + res.getResponseCode() + '): ' + res.getContentText());
    }
    return { status: 'pushed_smartlead' };
  }

  // 4. Instantly Sequencer Integration
  if (p.indexOf('instantly') >= 0) {
    const apiKey = getSenderApiKey_('INSTANTLY_API_KEY', spreadsheet);
    const campaignId = String(config.emailSenderEndpoint || config['Instantly Campaign ID'] || '').trim();
    if (!apiKey) throw new Error('INSTANTLY_API_KEY not found in Project Settings or Agent Settings.');
    if (!campaignId) throw new Error('Instantly Campaign ID is missing in Email Sender Endpoint.');

    const url = 'https://api.instantly.ai/api/v1/lead/add';
    const nameParts = (item.leadName || '').trim().split(' ');
    const firstName = nameParts[0] || '';
    const lastName = nameParts.slice(1).join(' ') || '';

    const payload = {
      api_key: apiKey,
      campaign_id: campaignId,
      email: item.email,
      first_name: firstName,
      last_name: lastName,
      company_name: item.companyName,
      custom_variables: {
        draft_subject: item.subject,
        draft_body: item.body,
        sender_name: senderName
      }
    };

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() >= 300) {
      throw new Error('Instantly API error (' + res.getResponseCode() + '): ' + res.getContentText());
    }
    return { status: 'pushed_instantly' };
  }

  // 5. Custom HTTPS Webhook (Zapier, Make, n8n, CRM)
  if (p.indexOf('webhook') >= 0 || p.indexOf('custom') >= 0) {
    const webhookUrl = String(config.emailSenderEndpoint || '').trim();
    if (!webhookUrl || !/^https:\/\//i.test(webhookUrl)) {
      throw new Error('A valid HTTPS Webhook Endpoint is required in Email Sender Endpoint.');
    }

    const payload = {
      campaignId: item.campaignId,
      email: item.email,
      leadName: item.leadName,
      companyName: item.companyName,
      subject: item.subject,
      body: item.body,
      senderName: senderName,
      senderTitle: config.senderTitle || '',
      senderCompany: config.senderCompany || '',
      timestamp: new Date().toISOString()
    };

    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() >= 300) {
      throw new Error('Webhook error (' + res.getResponseCode() + '): ' + res.getContentText());
    }
    return { status: 'pushed_webhook' };
  }

  throw new Error('Unsupported email sender provider: ' + provider);
}

/**
 * Dynamically resolves sender profile and provider configuration per Campaign ID.
 * Scans Campaigns (28 cols), Persona Mapper Settings (39 cols), and Campaign Configuration.
 * 
 * @param {Spreadsheet} spreadsheet Active spreadsheet instance
 * @param {string} campaignId Target campaign identifier
 * @return {Object} Resolved settings object
 */
function getCampaignSenderSettings_(spreadsheet, campaignId) {
  const settings = {
    campaignId: campaignId,
    emailSenderProvider: 'Workspace Gmail (Drafts Only)',
    emailSenderEndpoint: '',
    senderName: '',
    senderTitle: '',
    senderCompany: ''
  };

  if (!spreadsheet || !campaignId) return settings;

  // 1. Try reading from Campaigns tab (Locked 28 columns)
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const campData = campSheet.getDataRange().getValues();
    const headers = campData[0].map(function(h) {
      return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    });

    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const provIdx = headers.indexOf('EMAILSENDERPROVIDER');
    const endIdx = headers.indexOf('EMAILSENDERENDPOINT');
    const nameIdx = headers.indexOf('SENDERNAME');
    const titleIdx = headers.indexOf('SENDERTITLE');
    const compIdx = headers.indexOf('SENDERCOMPANY');
    const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON');

    for (let r = 1; r < campData.length; r++) {
      if (String(campData[r][idIdx] || '').trim().toLowerCase() === campaignId.toLowerCase()) {
        // Read JSON backup if present
        if (jsonIdx >= 0 && campData[r][jsonIdx]) {
          try {
            const parsed = JSON.parse(campData[r][jsonIdx]);
            if (parsed && typeof parsed === 'object') {
              if (parsed.emailSenderProvider) settings.emailSenderProvider = parsed.emailSenderProvider;
              if (parsed.emailSenderEndpoint) settings.emailSenderEndpoint = parsed.emailSenderEndpoint;
              if (parsed.senderName) settings.senderName = parsed.senderName;
              if (parsed.senderTitle) settings.senderTitle = parsed.senderTitle;
              if (parsed.senderCompany) settings.senderCompany = parsed.senderCompany;
            }
          } catch (e) {}
        }
        // Direct column overrides
        if (provIdx >= 0 && campData[r][provIdx]) settings.emailSenderProvider = String(campData[r][provIdx]).trim();
        if (endIdx >= 0 && campData[r][endIdx]) settings.emailSenderEndpoint = String(campData[r][endIdx]).trim();
        if (nameIdx >= 0 && campData[r][nameIdx]) settings.senderName = String(campData[r][nameIdx]).trim();
        if (titleIdx >= 0 && campData[r][titleIdx]) settings.senderTitle = String(campData[r][titleIdx]).trim();
        if (compIdx >= 0 && campData[r][compIdx]) settings.senderCompany = String(campData[r][compIdx]).trim();
        return settings;
      }
    }
  }

  // 2. Try reading from Persona Mapper Settings (Locked 39 columns)
  const pSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let r = 1; r < pData.length; r++) {
      if (String(pData[r][0] || '').trim().toLowerCase() === campaignId.toLowerCase()) {
        if (pData[r][37]) settings.emailSenderProvider = String(pData[r][37]).trim(); // Col 38
        if (pData[r][38]) settings.emailSenderEndpoint = String(pData[r][38]).trim(); // Col 39
        break;
      }
    }
  }

  // 3. Fallback to Campaign Configuration if matching
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  if (configSheet && configSheet.getLastRow() > 1) {
    const cfgData = configSheet.getDataRange().getValues();
    const cfgDict = {};
    for (let r = 1; r < cfgData.length; r++) {
      const k = String(cfgData[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (k) cfgDict[k] = cfgData[r][1];
    }
    const currentActiveId = String(cfgDict['CAMPAIGNID'] || '').trim();
    if (currentActiveId && currentActiveId.toLowerCase() === campaignId.toLowerCase()) {
      if (cfgDict['EMAILSENDERPROVIDER']) settings.emailSenderProvider = String(cfgDict['EMAILSENDERPROVIDER']).trim();
      if (cfgDict['EMAILSENDERENDPOINT']) settings.emailSenderEndpoint = String(cfgDict['EMAILSENDERENDPOINT']).trim();
      if (cfgDict['SENDERNAME']) settings.senderName = String(cfgDict['SENDERNAME']).trim();
      if (cfgDict['SENDERTITLE']) settings.senderTitle = String(cfgDict['SENDERTITLE']).trim();
      if (cfgDict['SENDERCOMPANY']) settings.senderCompany = String(cfgDict['SENDERCOMPANY']).trim();
    }
  }

  return settings;
}

/**
 * Universal API Key resolver that checks:
 * 1. Script Properties
 * 2. User Properties
 * 3. Tab 10 (Agent Settings: Category | Setting Key | Value | Status / Description)
 * 
 * @param {string} keyName The key identifier (e.g., 'SMARTLEAD_API_KEY', 'INSTANTLY_API_KEY')
 * @param {Spreadsheet} spreadsheet Active spreadsheet instance
 * @return {string} Configured API key string or empty string
 */
function getSenderApiKey_(keyName, spreadsheet) {
  const targetKey = String(keyName || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  // 1. Check Script Properties
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  for (let k in scriptProps) {
    if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetKey) {
      const val = String(scriptProps[k] || '').trim();
      if (val && !val.includes('PASTE_')) return val;
    }
  }

  // 2. Check User Properties
  const userProps = PropertiesService.getUserProperties().getProperties();
  for (let k in userProps) {
    if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetKey) {
      const val = String(userProps[k] || '').trim();
      if (val && !val.includes('PASTE_')) return val;
    }
  }

  // 3. Check Tab 10: Agent Settings (Category | Setting Key | Value | Status / Description)
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const data = agentSheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        // Col B (index 1) is 'Setting Key', Col A (index 0) is 'Category', Col C (index 2) is 'Value'
        const settingKey = String(data[r][1] || data[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (settingKey === targetKey || settingKey.includes(targetKey)) {
          const val = String(data[r][2] || '').trim();
          if (val && !val.includes('PASTE_')) return val;
        }
      }
    }
  }

  return '';
}