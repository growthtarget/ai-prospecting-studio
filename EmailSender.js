/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * PLUGGABLE EMAIL SENDER & CLIENT-ISOLATED OUTREACH DISPATCH ENGINE
 * ============================================================================
 * File: EmailSender.gs / EmailSender.js
 *
 * CAPABILITIES:
 * 1. Multi-Provider Dispatch: Google Workspace (Local & Client Relay),
 *    Microsoft 365 / Exchange Online (Drafts Only & Direct Send via Graph API),
 *    Smartlead, Instantly, Custom Webhooks.
 * 2. Strict Client Isolation: External clients strictly require CLIENT_{NAME}_*
 *    credentials in Script Properties (zero fallback to local keys or personal Gmail).
 * 3. Campaign Isolation: Resolves sender profile and provider dynamically per prospect Campaign ID.
 * 4. Sender Identity & Mailbox Stamping: Attaches Sender Display Name, Sender Email/Mailbox,
 *    and Reply-To Email address to drafts and outreach emails.
 * 5. Deliverability Event Capture: Captures Draft / Message IDs in Prospect Review Queue
 *    and records inline "Sent" events to EmailEvents via EmailEventsDeliverability.js.
 * 6. Headless Safe: Gracefully executes in both UI dialog and silent background/sidebar modes.
 * ============================================================================
 */

/**
 * Universal sender: Dispatches outreach based on each campaign's configured provider.
 */
function sendApprovedOutreachEmails() {
  processOutreachQueue_(false);
}

/**
 * Force creates Drafts regardless of provider Direct Send setting (Safe testing & review mode).
 * Routes Microsoft 365 campaigns to Microsoft 365 Drafts and all other campaigns to Workspace Gmail Drafts.
 */
function createGmailDraftsForApprovedProspects() {
  processOutreachQueue_(true);
}

/**
 * Helper: Finds the first matching column index from a list of header aliases.
 */
function findSenderHeaderCol_(headers, candidates) {
  const normHeaders = headers.map(function(h) {
    return String(h || '').trim().toLowerCase();
  });
  for (let i = 0; i < candidates.length; i++) {
    const idx = normHeaders.indexOf(String(candidates[i]).toLowerCase());
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * Core outreach processor.
 * Resolves settings per prospect row, ensuring complete campaign and client isolation.
 *
 * @param {boolean} forceDrafts When true, routes all approved prospects to Drafts Only mode
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

  const campIdCol = findSenderHeaderCol_(headers, ['Campaign ID']);
  const emailCol = findSenderHeaderCol_(headers, ['Business Email']);
  const approvedSubjectCol = findSenderHeaderCol_(headers, ['Approved Email Subject']);
  const subjectCol = findSenderHeaderCol_(headers, ['Draft Subject', 'Email Subject']);
  const approvedBodyCol = findSenderHeaderCol_(headers, ['Approved Email Body']);
  const bodyCol = findSenderHeaderCol_(headers, ['Draft Email', 'Email Body']);
  const humanApprCol = findSenderHeaderCol_(headers, ['Human Approval Status']);
  const reviewStatusCol = findSenderHeaderCol_(headers, ['Review Status']);
  const sendStatusCol = findSenderHeaderCol_(headers, ['Send Status']);
  const sentAtCol = findSenderHeaderCol_(headers, ['Sent At']);
  const dispatchModeCol = findSenderHeaderCol_(headers, ['Dispatch Mode']);
  const threadMsgIdCol = findSenderHeaderCol_(headers, ['Thread / Message ID', 'Message ID']);
  const draftCreatedAtCol = findSenderHeaderCol_(headers, ['Draft Created At']);
  const sendErrorCol = findSenderHeaderCol_(headers, ['Send Error', 'Last Error']);
  const leadNameCol = findSenderHeaderCol_(headers, ['Name', 'Persona Name']);
  const companyCol = findSenderHeaderCol_(headers, ['Company Name']);

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const queueToProcess = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const email = emailCol >= 0 ? String(row[emailCol] || '').trim() : '';
    const approvedBody = approvedBodyCol >= 0 ? String(row[approvedBodyCol] || '').trim() : '';
    const baseBody = bodyCol >= 0 ? String(row[bodyCol] || '').trim() : '';
    const body = approvedBody || baseBody;

    const approvedSubj = approvedSubjectCol >= 0 ? String(row[approvedSubjectCol] || '').trim() : '';
    const baseSubj = subjectCol >= 0 ? String(row[subjectCol] || '').trim() : '';
    const subject = approvedSubj || baseSubj || 'Introduction';

    const humanAppr = humanApprCol >= 0 ? String(row[humanApprCol] || '').trim().toLowerCase() : '';
    const reviewStatus = reviewStatusCol >= 0 ? String(row[reviewStatusCol] || '').trim().toLowerCase() : '';
    const sendStatus = sendStatusCol >= 0 ? String(row[sendStatusCol] || '').trim().toLowerCase() : '';
    const campaignId = campIdCol >= 0 ? String(row[campIdCol] || '').trim() : '';

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
        subject: subject,
        body: body,
        leadName: leadNameCol >= 0 ? String(row[leadNameCol] || '') : '',
        companyName: companyCol >= 0 ? String(row[companyCol] || '') : ''
      });
    }
  }

  if (queueToProcess.length === 0) {
    const msg = 'No approved prospects are pending outreach.\n\n' +
      'To approve a prospect:\n' +
      '1. Set Human Approval Status or Review Status to "Approved"\n' +
      '2. Ensure Send Status is not already "Sent", "Draft Created", "Draft Created in Gmail", or "Suppressed - Not Sent"\n' +
      '3. Ensure Business Email and Email Subject/Body are filled.';
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
    const modeLabel = forceDrafts ? 'Safe Mode: Drafts Only' : 'Configured Providers';
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
  const nowDate = new Date();
  const nowStr = nowDate.toLocaleString();
  const msTokenCache = {};

  queueToProcess.forEach(function(item) {
    const cId = item.campaignId || 'DEFAULT';
    const campaignConfig = campaignSettingsCache[cId] || getCampaignSenderSettings_(spreadsheet, cId);

    const configuredProvider = String(campaignConfig.emailSenderProvider || 'Workspace Gmail (Drafts Only)').trim();
    let rawProvider = configuredProvider;

    if (forceDrafts) {
      const cfgLower = configuredProvider.toLowerCase();
      if (isMicrosoft365Provider_(cfgLower)) {
        rawProvider = 'Microsoft 365 (Drafts Only)';
      } else {
        rawProvider = 'Workspace Gmail (Drafts Only)';
      }
    }

    try {
      const dispatchResult = dispatchOutreach_(item, rawProvider, campaignConfig, spreadsheet, msTokenCache) || {};

      const lowerProv = rawProvider.toLowerCase();
      let finalStatus = 'Sent';
      if (isMicrosoft365Provider_(lowerProv)) {
        finalStatus = lowerProv.indexOf('draft') >= 0 ? 'Draft Created' : 'Sent';
      } else if (lowerProv.indexOf('draft') >= 0) {
        finalStatus = 'Draft Created in Gmail';
      } else if (
        lowerProv.indexOf('smartlead') >= 0 ||
        lowerProv.indexOf('instantly') >= 0 ||
        lowerProv.indexOf('webhook') >= 0 ||
        lowerProv.indexOf('custom') >= 0
      ) {
        finalStatus = 'Pushed to Sequencer';
      }

      if (sendStatusCol >= 0) sheet.getRange(item.rowIndex, sendStatusCol + 1).setValue(finalStatus);
      if (dispatchModeCol >= 0) sheet.getRange(item.rowIndex, dispatchModeCol + 1).setValue(rawProvider);

      const trackedId = String(dispatchResult.queueTrackId || dispatchResult.providerMessageId || dispatchResult.draftId || '').trim();
      if (threadMsgIdCol >= 0 && trackedId) {
        sheet.getRange(item.rowIndex, threadMsgIdCol + 1).setValue(trackedId);
      }

      if (draftCreatedAtCol >= 0 && (finalStatus === 'Draft Created' || finalStatus === 'Draft Created in Gmail')) {
        sheet.getRange(item.rowIndex, draftCreatedAtCol + 1).setValue(nowStr);
      }
      if (sentAtCol >= 0 && (finalStatus === 'Sent' || finalStatus === 'Pushed to Sequencer')) {
        sheet.getRange(item.rowIndex, sentAtCol + 1).setValue(nowStr);
      }
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue('');

      // Record Sent events inline to EmailEvents for real-time deliverability tracking
      if (finalStatus === 'Sent' &&
          (dispatchResult.provider === 'Gmail' || dispatchResult.provider === 'Microsoft 365') &&
          typeof recordEmailEvent_ === 'function') {
        try {
          recordEmailEvent_({
            client_id: campaignConfig.clientName || 'Growth Target',
            campaign_id: item.campaignId,
            provider: dispatchResult.provider,
            provider_message_id: dispatchResult.providerMessageId || trackedId,
            event_type: 'Sent',
            event_date: new Date(),
            recipient: item.email,
            smtp_code: '250 2.0.0',
            reason: 'Dispatched via ' + rawProvider
          }, spreadsheet);
        } catch (evtErr) {
          Logger.log('EmailEvents Sent logging notice: ' + evtErr.message);
        }
      }

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
    (failCount > 0 ? ('• Failed: ' + failCount + ' (Check "Send Error" / "Last Error" column for details)') : '• All approved records processed cleanly.');

  if (ui) {
    ui.alert('Outreach Completed', completionMsg, ui.ButtonSet.OK);
  } else {
    Logger.log(completionMsg);
  }
}

/**
 * Helper: Identifies Microsoft 365 / Exchange Online provider labels.
 */
function isMicrosoft365Provider_(providerLower) {
  const p = String(providerLower || '').toLowerCase();
  return (
    p.indexOf('microsoft') >= 0 ||
    p.indexOf('m365') >= 0 ||
    p.indexOf('exchange') >= 0 ||
    p.indexOf('outlook') >= 0
  );
}

/**
 * Universal Dispatcher: Routes to Microsoft 365 (Graph API), Google Workspace (local or client relay),
 * Smartlead, Instantly, or Custom Webhook.
 */
function dispatchOutreach_(item, provider, config, spreadsheet, msTokenCache) {
  const p = String(provider || '').toLowerCase();
  const senderName = String(config.senderName || '').trim();
  const senderEmail = String(config.senderEmail || '').trim();
  const replyToEmail = String(config.replyToEmail || senderEmail || '').trim();
  const clientName = String(config.clientName || 'Growth Target').trim();
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;

  // 1. Microsoft 365 / Exchange Online (Check BEFORE Workspace 'draft'/'direct' match)
  if (isMicrosoft365Provider_(p)) {
    return dispatchMicrosoft365Outreach_(item, p, config, spreadsheet, msTokenCache);
  }

  // 2. Google Workspace Drafts & Direct Send
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
        let draft;
        try {
          draft = GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
        } catch (aliasErr) {
          if (mailOptions.from) {
            delete mailOptions.from;
            draft = GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
          } else {
            throw aliasErr;
          }
        }
        const draftId = draft ? draft.getId() : '';
        return {
          status: 'draft_created',
          provider: 'Gmail',
          draftId: draftId,
          providerMessageId: draftId,
          queueTrackId: draftId
        };
      }

      // Direct Send via createDraft().send() so we capture GmailMessage ID & RFC Message-ID
      let draftToSend;
      try {
        draftToSend = GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
      } catch (aliasErr) {
        if (mailOptions.from) {
          delete mailOptions.from;
          draftToSend = GmailApp.createDraft(item.email, item.subject, item.body, mailOptions);
        } else {
          throw aliasErr;
        }
      }

      const sentMsg = draftToSend.send();
      let sentMsgId = '';
      let rfcMsgId = '';
      try {
        sentMsgId = sentMsg ? (sentMsg.getId() || '') : '';
        rfcMsgId = sentMsg ? (sentMsg.getHeader('Message-ID') || '') : '';
      } catch (idErr) {}

      const finalMsgId = rfcMsgId || sentMsgId || ('gmail-' + Utilities.getUuid());
      return {
        status: 'sent',
        provider: 'Gmail',
        messageId: sentMsgId,
        rfcMessageId: rfcMsgId,
        providerMessageId: finalMsgId,
        queueTrackId: finalMsgId
      };
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

    const relayDraftId = parsedRelay.draftId || '';
    const relayMsgId = parsedRelay.rfcMessageId || parsedRelay.messageId || relayDraftId || '';
    return {
      status: isDraftMode ? 'draft_created_client_relay' : 'sent_client_relay',
      provider: 'Gmail',
      draftId: relayDraftId,
      providerMessageId: relayMsgId,
      queueTrackId: isDraftMode ? relayDraftId : relayMsgId
    };
  }

  // 3. Smartlead Sequencer Integration (Option B)
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

  // 4. Instantly Sequencer Integration (Option B)
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

  // 5. Custom HTTPS Webhook (Zapier, Make, n8n, CRM)
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
 * Microsoft 365 / Exchange Online Dispatcher (Microsoft Graph API v1.0).
 * Strictly isolated per Client / Owner using MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET.
 * Uses ImmutableId + Draft-then-Send on Direct Send to capture internetMessageId & conversationId.
 */
function dispatchMicrosoft365Outreach_(item, providerLower, config, spreadsheet, msTokenCache) {
  const clientName = String(config.clientName || 'Growth Target').trim();
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
  const prefix = (!isLocal && typeof normalizeClientPropertyPrefix_ === 'function')
    ? normalizeClientPropertyPrefix_(clientName)
    : '';

  const tenantId = getSenderApiKey_('MS_TENANT_ID', spreadsheet, config);
  const clientId = getSenderApiKey_('MS_CLIENT_ID', spreadsheet, config);
  const clientSecret = getSenderApiKey_('MS_CLIENT_SECRET', spreadsheet, config);

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'Missing Microsoft 365 credentials for "' + clientName + '". ' +
      'Set "' + prefix + 'MS_TENANT_ID", "' + prefix + 'MS_CLIENT_ID", and "' + prefix + 'MS_CLIENT_SECRET" ' +
      'in Project Settings > Script Properties.'
    );
  }

  let senderEmail = String(config.senderEmail || '').trim();
  if (!senderEmail && typeof resolveClientScopedSecret_ === 'function') {
    const emailFallback = resolveClientScopedSecret_(
      ['MS_SENDER_EMAIL'],
      { campaignId: config.campaignId, clientName: clientName },
      { throwOnMissing: false, allowAgentSettingsForLocal: false, spreadsheet: spreadsheet }
    );
    if (emailFallback && emailFallback.key) {
      senderEmail = String(emailFallback.key).trim();
    }
  }

  if (!senderEmail || senderEmail.indexOf('@') < 1) {
    throw new Error(
      'Microsoft 365 requires a valid "Sender Email / Mailbox" (Exchange Online mailbox address, e.g. sales@clientdomain.com) ' +
      'in the Campaign Control Panel Sender Profile.'
    );
  }

  const senderName = String(config.senderName || '').trim();
  const replyToEmail = String(config.replyToEmail || senderEmail).trim();
  const isDraftMode = String(providerLower || '').toLowerCase().indexOf('draft') >= 0;

  const accessToken = getMicrosoftGraphAccessToken_(tenantId, clientId, clientSecret, clientName, msTokenCache);

  const toRecipient = {
    emailAddress: {
      address: item.email
    }
  };
  if (item.leadName) {
    toRecipient.emailAddress.name = item.leadName;
  }

  const messageObj = {
    subject: item.subject,
    body: {
      contentType: 'Text',
      content: item.body
    },
    toRecipients: [toRecipient]
  };

  if (replyToEmail && replyToEmail.indexOf('@') > 0) {
    const replyAddr = { address: replyToEmail };
    if (senderName) replyAddr.name = senderName;
    messageObj.replyTo = [{ emailAddress: replyAddr }];
  }

  // Create draft first (works for Drafts Only AND captures internetMessageId for Direct Send)
  const createDraftPayload = Object.assign({}, messageObj, { isDraft: true });
  const draftUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(senderEmail) + '/messages';
  const draftRes = UrlFetchApp.fetch(draftUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + accessToken,
      Prefer: 'IdType="ImmutableId"'
    },
    payload: JSON.stringify(createDraftPayload),
    muteHttpExceptions: true
  });

  const draftCode = draftRes.getResponseCode();
  const draftText = draftRes.getContentText() || '';

  if (draftCode >= 200 && draftCode < 300) {
    let parsedDraft = {};
    try {
      parsedDraft = JSON.parse(draftText || '{}');
    } catch (e) {}

    const draftId = parsedDraft.id || '';
    const internetMessageId = parsedDraft.internetMessageId || '';
    const conversationId = parsedDraft.conversationId || '';

    if (isDraftMode) {
      return {
        status: 'draft_created_microsoft365',
        provider: 'Microsoft 365',
        draftId: draftId,
        providerMessageId: internetMessageId || draftId,
        queueTrackId: draftId,
        conversationId: conversationId
      };
    }

    // Direct Send: Send the created draft (saves automatically to Sent Items)
    const sendDraftUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(senderEmail) +
      '/messages/' + encodeURIComponent(draftId) + '/send';
    const sendDraftRes = UrlFetchApp.fetch(sendDraftUrl, {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + accessToken
      },
      payload: '',
      muteHttpExceptions: true
    });

    const sendDraftCode = sendDraftRes.getResponseCode();
    const sendDraftText = sendDraftRes.getContentText() || '';
    if (sendDraftCode < 200 || sendDraftCode >= 300) {
      throw new Error(parseMicrosoftGraphError_(sendDraftCode, sendDraftText, 'send created draft', senderEmail));
    }

    const finalMsgId = internetMessageId || draftId;
    return {
      status: 'sent_microsoft365',
      provider: 'Microsoft 365',
      draftId: draftId,
      providerMessageId: finalMsgId,
      queueTrackId: finalMsgId,
      conversationId: conversationId
    };
  }

  // If Drafts Only failed, throw error immediately
  if (isDraftMode) {
    throw new Error(parseMicrosoftGraphError_(draftCode, draftText, 'create draft', senderEmail));
  }

  // If Direct Send failed on /messages with 403 (client only granted Mail.Send without Mail.ReadWrite), fallback to /sendMail
  if (draftCode === 403) {
    const sendUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(senderEmail) + '/sendMail';
    const sendRes = UrlFetchApp.fetch(sendUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + accessToken
      },
      payload: JSON.stringify({
        message: messageObj,
        saveToSentItems: true
      }),
      muteHttpExceptions: true
    });

    const sendCode = sendRes.getResponseCode();
    const sendText = sendRes.getContentText() || '';
    if (sendCode < 200 || sendCode >= 300) {
      throw new Error(parseMicrosoftGraphError_(sendCode, sendText, 'send email', senderEmail));
    }

    const fallbackId = 'm365-' + Utilities.getUuid();
    return {
      status: 'sent_microsoft365',
      provider: 'Microsoft 365',
      providerMessageId: fallbackId,
      queueTrackId: fallbackId
    };
  }

  throw new Error(parseMicrosoftGraphError_(draftCode, draftText, 'create message before send', senderEmail));
}

/**
 * Retrieves and caches an App-Only Microsoft Graph OAuth2 token via client_credentials flow.
 */
function getMicrosoftGraphAccessToken_(tenantId, clientId, clientSecret, clientName, msTokenCache) {
  const cache = msTokenCache || {};
  const cacheKey = tenantId + '::' + clientId;
  const cached = cache[cacheKey];

  if (cached && cached.token && cached.expiresAt > Date.now() + 60000) {
    return cached.token;
  }

  const tokenUrl = 'https://login.microsoftonline.com/' + encodeURIComponent(tenantId) + '/oauth2/v2.0/token';
  const res = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let parsed = {};
  try {
    parsed = JSON.parse(text || '{}');
  } catch (e) {}

  if (code < 200 || code >= 300 || !parsed.access_token) {
    const oauthErr = parsed.error_description || parsed.error || text.slice(0, 220);
    throw new Error(
      'Microsoft 365 OAuth authentication failed for "' + clientName + '" (HTTP ' + code + '): ' +
      String(oauthErr).replace(/\s+/g, ' ').trim()
    );
  }

  const expiresInSec = Number(parsed.expires_in) || 3599;
  cache[cacheKey] = {
    token: parsed.access_token,
    expiresAt: Date.now() + expiresInSec * 1000
  };

  return parsed.access_token;
}

/**
 * Formats Microsoft Graph HTTP errors into clear operational messages without exposing secrets.
 */
function parseMicrosoftGraphError_(statusCode, bodyText, actionLabel, senderEmail) {
  let errCode = '';
  let errMsg = String(bodyText || '').slice(0, 250);
  try {
    const parsed = JSON.parse(bodyText || '{}');
    if (parsed && parsed.error) {
      errCode = String(parsed.error.code || '');
      errMsg = String(parsed.error.message || errMsg);
    }
  } catch (e) {}

  let hint = '';
  if (statusCode === 403 || errCode === 'ErrorAccessDenied') {
    hint = ' Ensure Microsoft Graph Application permissions (Mail.ReadWrite for drafts, Mail.Send for direct send) have Admin Consent in Entra ID.';
  } else if (statusCode === 404 || errCode === 'ResourceNotFound' || errCode === 'ErrorInvalidUser' || errCode === 'MailboxNotEnabledForRESTAPI') {
    hint = ' Verify "' + senderEmail + '" is an active Exchange Online primary mailbox / UPN.';
  }

  return 'Microsoft Graph failed to ' + actionLabel + ' (HTTP ' + statusCode + (errCode ? ' - ' + errCode : '') + '): ' + errMsg + hint;
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

  const clientName = ctx.clientName || 'Growth Target';
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
  if (!isLocal) {
    const prefix = typeof normalizeClientPropertyPrefix_ === 'function'
      ? normalizeClientPropertyPrefix_(clientName)
      : ('CLIENT_' + clientName.toUpperCase().replace(/[^A-Z0-9]+/g, '_') + '_');
    const targetClientKey = (prefix + keyName).toUpperCase().replace(/[^A-Z0-9]/g, '');
    const sp = PropertiesService.getScriptProperties().getProperties();
    for (let k in sp) {
      if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetClientKey && isValidSenderSecretValue_(sp[k])) {
        return String(sp[k]).trim();
      }
    }
    throw new Error('Missing client credential "' + prefix + keyName + '" in Project Settings > Script Properties.');
  }

  const targetKey = String(keyName || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  for (let k in scriptProps) {
    if (String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
    if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === targetKey) {
      if (isValidSenderSecretValue_(scriptProps[k])) return String(scriptProps[k]).trim();
    }
  }

  const userProps = PropertiesService.getUserProperties().getProperties();
  for (let k in userProps) {
    if (String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
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