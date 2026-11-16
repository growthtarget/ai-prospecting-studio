/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * PLUGGABLE EMAIL SENDER & CLIENT-ISOLATED OUTREACH DISPATCH ENGINE
 * ============================================================================
 * File: EmailSender.gs / EmailSender.js
 * 
 * CAPABILITIES:
 * 1. Multi-Provider Dispatch: Google Workspace (Local & Client Relay), Smartlead,
 *    Instantly, Custom Webhooks.
 * 2. Strict Client Isolation: External clients strictly require CLIENT_{NAME}_*
 *    credentials in Script Properties (zero fallback to local keys or personal Gmail).
 * 3. Campaign Isolation: Resolves sender profile and provider dynamically per prospect Campaign ID.
 * 4. Sender Identity & Mailbox Stamping: Attaches Sender Display Name, Sender Email/Alias,
 *    and Reply-To Email address to drafts and outreach emails.
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
 * Resolves settings per prospect row, ensuring complete campaign and client isolation.
 * 
 * @param {boolean} forceDrafts When true, routes all approved prospects to Gmail Drafts
 */
function processOutreachQueue_(forceDrafts) {
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {}

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
    const emptyMsg = 'No prospects found in Prospect Review Queue.';
    if (ui) ui.alert(emptyMsg);
    else Logger.log(emptyMsg);
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
    const body = String(row[bodyCol] || '').trim();
    const humanAppr = String(row[humanApprCol] || '').trim().toLowerCase();
    const reviewStatus = String(row[reviewStatusCol] || '').trim().toLowerCase();
    const sendStatus = String(row[sendStatusCol] || '').trim().toLowerCase();
    const campaignId = String(row[campIdCol] || '').trim();

    const isRejected = (humanAppr.startsWith('rejected') || reviewStatus === 'rejected');
    const isApproved = !isRejected && (
      humanAppr.startsWith('approved') ||
      reviewStatus === 'approved' ||
      reviewStatus === 'approved for outreach'
    );
    const isAlreadySent = (
      sendStatus === 'sent' ||
      sendStatus === 'draft created' ||
      sendStatus === 'draft created in gmail' ||
      sendStatus === 'pushed to sequencer' ||
      sendStatus === 'suppressed - not sent'
    );

    if (email && body && isApproved && !isAlreadySent) {
      queueToProcess.push({
        rowIndex: i + 2,
        campaignId: campaignId,
        email: email,
        subject: String(row[subjectCol] || 'Introduction'),
        body: body,
        leadName: String(row[leadNameCol] || ''),
        companyName: String(row[companyCol] || '')
      });
    }
  }

  if (queueToProcess.length === 0) {
    const msg = 'No approved prospects are pending outreach.\n\n' +
      'To approve a prospect:\n' +
      '1. Set Human Approval Status or Review Status to "Approved"\n' +
      '2. Ensure Send Status is not already "Sent", "Draft Created in Gmail", or "Suppressed - Not Sent"\n' +
      '3. Ensure Business Email, Draft Subject, and Draft Email are filled.';
    if (ui) ui.alert('No Prospects Ready', msg, ui.ButtonSet.OK);
    else Logger.log(msg);
    return;
  }

  const campaignSettingsCache = {};
  const uniqueCampaigns = Array.from(new Set(queueToProcess.map(function(item) {
    return item.campaignId || 'DEFAULT';
  })));

  uniqueCampaigns.forEach(function(cId) {
    campaignSettingsCache[cId] = getCampaignSenderSettings_(spreadsheet, cId);
  });

  const uniqueClients = Array.from(new Set(uniqueCampaigns.map(function(cId) {
    return campaignSettingsCache[cId].clientName || 'Growth Target';
  })));

  if (ui) {
    const modeLabel = forceDrafts ? 'Safe Mode: Workspace Gmail (Drafts Only)' : 'Configured Providers';
    const confirm = ui.alert(
      'Confirm Outreach Execution',
      'Mode: ' + modeLabel + '\n' +
      'Prospects to process: ' + queueToProcess.length + '\n' +
      'Target Campaign(s): ' + uniqueCampaigns.join(', ') + '\n' +
      'Client / Owner(s): ' + uniqueClients.join(', ') + '\n\n' +
      'Do you want to proceed?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  let successCount = 0;
  let failCount = 0;
  const nowStr = new Date().toLocaleString();

  queueToProcess.forEach(function(item) {
    const cId = item.campaignId || 'DEFAULT';
    const campaignConfig = campaignSettingsCache[cId] || getCampaignSenderSettings_(spreadsheet, cId);

    const rawProvider = forceDrafts
      ? 'Workspace Gmail (Drafts Only)'
      : String(campaignConfig.emailSenderProvider || 'Workspace Gmail (Drafts Only)').trim();

    try {
      dispatchOutreach_(item, rawProvider, campaignConfig, spreadsheet);

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

      if (sendStatusCol >= 0) sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue(finalStatus);
      if (sentAtCol >= 0 && (finalStatus === 'Sent' || finalStatus === 'Pushed to Sequencer')) {
        sheet.getRange(item.rowIndex, sentAtCol + 1).setValue(nowStr);
      }
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

  if (typeof updateDashboardForActiveCampaign === 'function' && typeof getStudioActiveCampaignId_ === 'function') {
    try {
      updateDashboardForActiveCampaign(getStudioActiveCampaignId_());
    } catch (dashErr) {
      Logger.log('Dashboard refresh notice: ' + dashErr.message);
    }
  }

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
 * Universal Dispatcher: Routes to Google Workspace (local or client relay), Smartlead, Instantly, or Custom Webhook.
 */
function dispatchOutreach_(item, provider, config, spreadsheet) {
  const p = String(provider || '').toLowerCase();
  const senderName = String(config.senderName || '').trim();
  const senderEmail = String(config.senderEmail || '').trim();
  const replyToEmail = String(config.replyToEmail || senderEmail || '').trim();
  const clientName = String(config.clientName || 'Growth Target').trim();
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;

  // 1. Google Workspace Drafts & Direct Send
  if (p.indexOf('draft') >= 0 || p.indexOf('gmail') >= 0 || p.indexOf('workspace') >= 0 || p.indexOf('direct') >= 0) {
    const isDraftMode = p.indexOf('draft') >= 0;

    if (isLocal) {
      // Local Agency ("Growth Target"): Use your logged-in Google Workspace GmailApp
      const mailOptions = {};
      if (senderName) mailOptions.name = senderName;
      if (replyToEmail && replyToEmail.indexOf('@') > 0) mailOptions.replyTo = replyToEmail;

      if (senderEmail && senderEmail.indexOf('@') > 0) {
        try {
          const currentUser = Session.getActiveUser().getEmail();
          if (currentUser && senderEmail.toLowerCase() !== currentUser.toLowerCase()) {
            mailOptions.from = senderEmail;
          }
        } catch (e) {}
      }

      if (isDraftMode) {
        try {
          GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
        } catch (aliasErr) {
          if (mailOptions.from) {
            delete mailOptions.from;
            GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
          } else {
            throw aliasErr;
          }
        }
        return { status: 'draft_created' };
      }

      try {
        GmailApp.sendEmail(item.email, item.subject, item.body, mailOptions);
      } catch (aliasErr) {
        if (mailOptions.from) {
          delete mailOptions.from;
          GmailApp.sendEmail(item.email, item.subject, item.body, mailOptions);
        } else {
          throw aliasErr;
        }
      }
      return { status: 'sent' };
    }

    // External Client Workspace Relay (Option A): Strictly routes via client's own Relay Web App
    const relayUrl = getSenderApiKey_('GMAIL_RELAY_URL', spreadsheet, config);
    const relaySecret = getSenderApiKey_('GMAIL_RELAY_SECRET', spreadsheet, config);

    if (!/^https:\/\//i.test(relayUrl)) {
      throw new Error('Client Gmail Relay URL for "' + clientName + '" must start with https://.');
    }

    const relayPayload = {
      secret: relaySecret,
      action: isDraftMode ? 'createDraft' : 'sendEmail',
      to: item.email,
      subject: item.subject,
      body: item.body,
      senderName: senderName,
      senderEmail: senderEmail,
      replyTo: replyToEmail,
      campaignId: item.campaignId,
      clientName: clientName
    };

    const relayRes = UrlFetchApp.fetch(relayUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(relayPayload),
      muteHttpExceptions: true
    });

    const relayCode = relayRes.getResponseCode();
    const relayText = relayRes.getContentText() || '';

    if (relayCode >= 300) {
      throw new Error('Client Gmail Relay HTTP ' + relayCode + ': ' + relayText.slice(0, 200));
    }

    let parsedRelay;
    try {
      parsedRelay = JSON.parse(relayText);
    } catch (parseErr) {
      throw new Error(
        'Client Gmail Relay for "' + clientName + '" returned non-JSON response. ' +
        'Ensure the client\'s Relay Web App is deployed with "Execute as: Me" and "Who has access: Anyone".'
      );
    }

    if (!parsedRelay || parsedRelay.success !== true) {
      throw new Error('Client Gmail Relay error: ' + (parsedRelay && parsedRelay.error ? parsedRelay.error : relayText.slice(0, 200)));
    }

    return { status: isDraftMode ? 'draft_created_client_relay' : 'sent_client_relay' };
  }

  // 2. Smartlead Sequencer Integration (Option B)
  if (p.indexOf('smartlead') >= 0) {
    const apiKey = getSenderApiKey_('SMARTLEAD_API_KEY', spreadsheet, config);
    const campaignId = String(config.emailSenderEndpoint || config['Smartlead Campaign ID'] || '').trim();
    if (!apiKey) throw new Error('SMARTLEAD_API_KEY not found in Project Settings > Script Properties.');
    if (!campaignId) throw new Error('Smartlead Campaign ID is missing in Email Sender Endpoint.');

    const url = 'https://server.smartlead.ai/api/v1/campaigns/' + encodeURIComponent(campaignId) + '/leads?api_key=' + encodeURIComponent(apiKey);
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
          sender_name: senderName,
          sender_email: senderEmail,
          reply_to: replyToEmail,
          client_name: clientName
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

  // 3. Instantly Sequencer Integration (Option B)
  if (p.indexOf('instantly') >= 0) {
    const apiKey = getSenderApiKey_('INSTANTLY_API_KEY', spreadsheet, config);
    const campaignId = String(config.emailSenderEndpoint || config['Instantly Campaign ID'] || '').trim();
    if (!apiKey) throw new Error('INSTANTLY_API_KEY not found in Project Settings > Script Properties.');
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
        sender_name: senderName,
        sender_email: senderEmail,
        reply_to: replyToEmail,
        client_name: clientName
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

  // 4. Custom HTTPS Webhook (Zapier, Make, n8n, CRM)
  if (p.indexOf('webhook') >= 0 || p.indexOf('custom') >= 0) {
    let webhookUrl = String(config.emailSenderEndpoint || '').trim();
    if (!webhookUrl && typeof resolveClientScopedSecret_ === 'function') {
      const urlRes = resolveClientScopedSecret_(
        ['WEBHOOK_URL', 'EMAIL_SENDER_WEBHOOK_URL'],
        { campaignId: config.campaignId, clientName: clientName },
        { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: spreadsheet }
      );
      if (urlRes.key) webhookUrl = urlRes.key;
    }

    if (!webhookUrl || !/^https:\/\//i.test(webhookUrl)) {
      throw new Error('A valid HTTPS Webhook Endpoint is required in Email Sender Endpoint.');
    }

    const headers = {};
    if (typeof resolveClientScopedSecret_ === 'function') {
      const secretRes = resolveClientScopedSecret_(
        ['WEBHOOK_SECRET'],
        { campaignId: config.campaignId, clientName: clientName },
        { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: spreadsheet }
      );
      if (secretRes.key) {
        headers['X-Webhook-Secret'] = secretRes.key;
      }
    }

    const payload = {
      campaignId: item.campaignId,
      clientName: clientName,
      email: item.email,
      leadName: item.leadName,
      companyName: item.companyName,
      subject: item.subject,
      body: item.body,
      senderName: senderName,
      senderTitle: config.senderTitle || '',
      senderCompany: config.senderCompany || '',
      senderEmail: senderEmail,
      replyToEmail: replyToEmail,
      timestamp: new Date().toISOString()
    };

    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
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
 * Dynamically resolves sender profile, provider configuration, and Client / Owner per Campaign ID.
 * Scans Campaigns (28 cols), Persona Mapper Settings (39 cols), and Campaign Configuration.
 * 
 * @param {Spreadsheet} spreadsheet Active spreadsheet instance
 * @param {string} campaignId Target campaign identifier
 * @return {Object} Resolved settings object
 */
function getCampaignSenderSettings_(spreadsheet, campaignId) {
  const settings = {
    campaignId: campaignId,
    clientName: '',
    emailSenderProvider: '',
    emailSenderEndpoint: '',
    senderName: '',
    senderTitle: '',
    senderCompany: '',
    senderEmail: '',
    replyToEmail: ''
  };

  if (!spreadsheet || !campaignId) {
    settings.clientName = 'Growth Target';
    settings.emailSenderProvider = 'Workspace Gmail (Drafts Only)';
    return settings;
  }

  // 1. Read from Campaigns tab (Locked 28 columns)
  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const campData = campSheet.getDataRange().getValues();
    const headers = campData[0].map(function(h) {
      return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    });

    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const deptIdx = headers.indexOf('DEPARTMENT') >= 0 ? headers.indexOf('DEPARTMENT') : 2;
    const provIdx = headers.indexOf('EMAILSENDERPROVIDER');
    const endIdx = headers.indexOf('EMAILSENDERENDPOINT');
    const nameIdx = headers.indexOf('SENDERNAME');
    const titleIdx = headers.indexOf('SENDERTITLE');
    const compIdx = headers.indexOf('SENDERCOMPANY');
    const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON');

    for (let r = 1; r < campData.length; r++) {
      if (String(campData[r][idIdx] || '').trim().toLowerCase() === campaignId.toLowerCase()) {
        if (jsonIdx >= 0 && campData[r][jsonIdx]) {
          try {
            const parsed = JSON.parse(campData[r][jsonIdx]);
            if (parsed && typeof parsed === 'object') {
              if (parsed.department) settings.clientName = String(parsed.department).trim();
              if (parsed.emailSenderProvider) settings.emailSenderProvider = parsed.emailSenderProvider;
              if (parsed.emailSenderEndpoint) settings.emailSenderEndpoint = parsed.emailSenderEndpoint;
              if (parsed.senderName) settings.senderName = parsed.senderName;
              if (parsed.senderTitle) settings.senderTitle = parsed.senderTitle;
              if (parsed.senderCompany) settings.senderCompany = parsed.senderCompany;
              if (parsed.senderEmail) settings.senderEmail = String(parsed.senderEmail).trim();
              if (parsed.replyToEmail) settings.replyToEmail = String(parsed.replyToEmail).trim();
            }
          } catch (e) {}
        }
        if (deptIdx >= 0 && campData[r][deptIdx]) settings.clientName = String(campData[r][deptIdx]).trim();
        if (provIdx >= 0 && campData[r][provIdx]) settings.emailSenderProvider = String(campData[r][provIdx]).trim();
        if (endIdx >= 0 && campData[r][endIdx]) settings.emailSenderEndpoint = String(campData[r][endIdx]).trim();
        if (nameIdx >= 0 && campData[r][nameIdx]) settings.senderName = String(campData[r][nameIdx]).trim();
        if (titleIdx >= 0 && campData[r][titleIdx]) settings.senderTitle = String(campData[r][titleIdx]).trim();
        if (compIdx >= 0 && campData[r][compIdx]) settings.senderCompany = String(campData[r][compIdx]).trim();
        break;
      }
    }
  }

  // 2. Fill missing fields from Persona Mapper Settings (Locked 39 columns)
  const pSheet = spreadsheet.getSheetByName('Persona Mapper Settings');
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let r = 1; r < pData.length; r++) {
      if (String(pData[r][0] || '').trim().toLowerCase() === campaignId.toLowerCase()) {
        if (!settings.emailSenderProvider && pData[r][37]) settings.emailSenderProvider = String(pData[r][37]).trim();
        if (!settings.emailSenderEndpoint && pData[r][38]) settings.emailSenderEndpoint = String(pData[r][38]).trim();
        break;
      }
    }
  }

  // 3. Fill remaining blanks from Campaign Configuration if matching
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
      if (!settings.clientName && (cfgDict['DEPARTMENT'] || cfgDict['CLIENTOWNER'] || cfgDict['CLIENT'])) {
        settings.clientName = String(cfgDict['DEPARTMENT'] || cfgDict['CLIENTOWNER'] || cfgDict['CLIENT']).trim();
      }
      if (!settings.emailSenderProvider && (cfgDict['EMAILSENDERPROVIDER'] || cfgDict['DISPATCHMODE'])) {
        settings.emailSenderProvider = String(cfgDict['EMAILSENDERPROVIDER'] || cfgDict['DISPATCHMODE']).trim();
      }
      if (!settings.emailSenderEndpoint && cfgDict['EMAILSENDERENDPOINT']) settings.emailSenderEndpoint = String(cfgDict['EMAILSENDERENDPOINT']).trim();
      if (!settings.senderName && cfgDict['SENDERNAME']) settings.senderName = String(cfgDict['SENDERNAME']).trim();
      if (!settings.senderTitle && cfgDict['SENDERTITLE']) settings.senderTitle = String(cfgDict['SENDERTITLE']).trim();
      if (!settings.senderCompany && cfgDict['SENDERCOMPANY']) settings.senderCompany = String(cfgDict['SENDERCOMPANY']).trim();
      if (!settings.senderEmail && (cfgDict['SENDEREMAIL'] || cfgDict['SENDERMAILBOX'])) settings.senderEmail = String(cfgDict['SENDEREMAIL'] || cfgDict['SENDERMAILBOX']).trim();
      if (!settings.replyToEmail && (cfgDict['REPLYTOEMAIL'] || cfgDict['REPLYTO'])) settings.replyToEmail = String(cfgDict['REPLYTOEMAIL'] || cfgDict['REPLYTO']).trim();
    }
  }

  if (!settings.clientName) {
    settings.clientName = typeof resolveCampaignClientName_ === 'function'
      ? resolveCampaignClientName_(campaignId, '', spreadsheet)
      : 'Growth Target';
  }

  if (!settings.emailSenderProvider) {
    settings.emailSenderProvider = 'Workspace Gmail (Drafts Only)';
  }

  return settings;
}

function isValidSenderSecretValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const upper = s.toUpperCase();
  if (upper.indexOf('PASTE_') >= 0) return false;
  if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
  if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
  return true;
}

/**
 * Strict Client-Scoped Sender Credential Resolver.
 * For external clients, strictly requires CLIENT_{NAME}_{keyName} (zero local fallback).
 */
function getSenderApiKey_(keyName, spreadsheet, config) {
  const ctx = {
    campaignId: (config && config.campaignId) ? config.campaignId : '',
    clientName: (config && config.clientName) ? config.clientName : ''
  };

  if (typeof resolveClientScopedSecret_ === 'function') {
    const resolved = resolveClientScopedSecret_(
      [keyName],
      ctx,
      {
        serviceLabel: 'Email Sender (' + keyName + ')',
        throwOnMissing: true,
        allowAgentSettingsForLocal: true,
        spreadsheet: spreadsheet
      }
    );
    return resolved.key || '';
  }

  const targetKey = String(keyName || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  for (let k in scriptProps) {
    if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetKey) {
      if (isValidSenderSecretValue_(scriptProps[k])) return String(scriptProps[k]).trim();
    }
  }

  const userProps = PropertiesService.getUserProperties().getProperties();
  for (let k in userProps) {
    if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetKey) {
      if (isValidSenderSecretValue_(userProps[k])) return String(userProps[k]).trim();
    }
  }

  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const data = agentSheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        const settingKey = String(data[r][1] || data[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (settingKey === targetKey || settingKey.includes(targetKey)) {
          const val = String(data[r][2] || '').trim();
          if (isValidSenderSecretValue_(val)) return val;
        }
      }
    }
  }

  return '';
}