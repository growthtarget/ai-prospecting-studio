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
 * 3. Active Campaign Scope Guard: Processes ONLY the Active Campaign so approved
 *    leads from other clients/campaigns are never dispatched accidentally.
 * 4. Runtime Suppression, Duplicate-Send & Daily Send-Limit Guardrails + Pacing.
 * 5. CASL / B2B Opt-Out Compliance Footer: Automatically appends client-scoped
 *    sender identification and an opt-out instruction when not already present.
 * 6. Deliverability Event Capture: Captures Draft / Message IDs in Prospect Review Queue
 *    (AT:AV) and records inline "Sent" events to EmailEvents.
 * 7. Headless Safe: Gracefully executes in both UI dialog and silent background/sidebar modes.
 * ============================================================================
 */

/**
 * Universal sender: Dispatches outreach for the Active Campaign based on its configured provider.
 */
function sendApprovedOutreachEmails() {
  processOutreachQueue_(false);
}

/**
 * Force creates Drafts for the Active Campaign regardless of provider Direct Send setting (Safe testing & review mode).
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
 * Ensures trailing tracking columns ("Dispatch Mode", "Thread / Message ID", "Draft Created At")
 * exist at the end of "Prospect Review Queue" without shifting existing columns A:AS.
 */
function ensureReviewQueueTrackingColumns_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  const trackingHeaders = ['Dispatch Mode', 'Thread / Message ID', 'Draft Created At'];
  const missing = trackingHeaders.filter(function(h) {
    return existingHeaders.indexOf(h) < 0;
  });
  if (missing.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    if (sheet.getMaxColumns() < startCol + missing.length - 1) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), (startCol + missing.length - 1) - sheet.getMaxColumns());
    }
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  }
}

/**
 * Loads active suppression rules from "Lists & Settings" for runtime safety checks.
 */
function loadSenderSuppressionRules_(ss) {
  const rules = { domains: {}, emails: {} };
  if (!ss) return rules;
  const sheet = ss.getSheetByName('Lists & Settings');
  if (!sheet || sheet.getLastRow() <= 1) return rules;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const listName = String(data[i][0] || '').trim().toLowerCase();
    const val = String(data[i][1] || '').trim().toLowerCase();
    const status = String(data[i][2] || '').trim().toLowerCase();
    const isActive = (status === 'active' || status === 'true' || status === '');
    const isSuppression = (
      listName.indexOf('suppression') >= 0 ||
      listName.indexOf('blacklist') >= 0 ||
      listName.indexOf('opt-out') >= 0 ||
      listName.indexOf('opt out') >= 0 ||
      listName.indexOf('unsubscribe') >= 0 ||
      listName.indexOf('competitor') >= 0
    );
    if (isActive && isSuppression && val) {
      if (val.indexOf('@') >= 0) {
        rules.emails[val] = true;
      } else {
        rules.domains[val.replace(/^@/, '')] = true;
      }
    }
  }
  return rules;
}

/**
 * Collects already-dispatched (campaignId::email) keys from Prospect Review Queue & EmailEvents.
 */
function loadAlreadyDispatchedLookup_(ss, queueData, campIdCol, emailCol, sendStatusCol) {
  const dispatched = {};

  for (let i = 0; i < queueData.length; i++) {
    const row = queueData[i];
    const cid = campIdCol >= 0 ? String(row[campIdCol] || '').trim().toLowerCase() : '';
    const email = emailCol >= 0 ? String(row[emailCol] || '').trim().toLowerCase() : '';
    const st = sendStatusCol >= 0 ? String(row[sendStatusCol] || '').trim().toLowerCase() : '';
    if (!email) continue;
    if (st === 'sent' || st === 'draft created' || st === 'draft created in gmail' || st === 'pushed to sequencer') {
      dispatched[cid + '::' + email] = true;
    }
  }

  const eventsSheet = ss ? ss.getSheetByName('EmailEvents') : null;
  if (eventsSheet && eventsSheet.getLastRow() > 1) {
    const eData = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 8).getValues();
    for (let i = 0; i < eData.length; i++) {
      const cid = String(eData[i][2] || '').trim().toLowerCase();
      const eType = String(eData[i][5] || '').trim().toLowerCase();
      const recip = String(eData[i][7] || '').trim().toLowerCase();
      if (recip && (eType === 'sent' || eType === 'delivered')) {
        dispatched[cid + '::' + recip] = true;
      }
    }
  }

  return dispatched;
}

/**
 * Reads per-client DAILY_SEND_LIMIT from Agent Settings / Script Properties (defaults to 25).
 */
function getClientDailySendLimit_(ss, clientName) {
  const cleanClient = String(clientName || 'Growth Target').trim();
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(cleanClient) : true;
  const prefix = (!isLocal && typeof normalizeClientPropertyPrefix_ === 'function')
    ? normalizeClientPropertyPrefix_(cleanClient)
    : '';
  const targetKey = isLocal ? 'DAILY_SEND_LIMIT' : (prefix + 'DAILY_SEND_LIMIT');
  const normTarget = targetKey.toUpperCase().replace(/[^A-Z0-9]/g, '');

  function parseLimit(raw) {
    if (raw === '' || raw === null || raw === undefined || raw instanceof Date) return NaN;
    const n = parseInt(String(raw).replace(/[,\s]/g, '').trim(), 10);
    return isFinite(n) && n >= 0 ? n : NaN;
  }

  if (ss) {
    const agentSheet = ss.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const rows = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, 3).getValues();
      for (let i = 0; i < rows.length; i++) {
        const k = String(rows[i][1] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (k === normTarget) {
          const val = parseLimit(rows[i][2]);
          if (!isNaN(val)) return val;
        }
      }
    }
  }

  const spVal = PropertiesService.getScriptProperties().getProperty(targetKey);
  const parsedSp = parseLimit(spVal);
  if (!isNaN(parsedSp)) return parsedSp;

  return 25;
}

/**
 * Counts how many live sends occurred today for clientName across EmailEvents & Prospect Review Queue.
 */
function countClientSendsToday_(ss, clientName, queueData, campIdCol, emailCol, sendStatusCol, sentAtCol) {
  const tz = (ss && ss.getSpreadsheetTimeZone) ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  const todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const clientLow = String(clientName || 'Growth Target').trim().toLowerCase();
  const seenSent = {};

  const eventsSheet = ss ? ss.getSheetByName('EmailEvents') : null;
  if (eventsSheet && eventsSheet.getLastRow() > 1) {
    const eData = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 8).getValues();
    for (let i = 0; i < eData.length; i++) {
      const rowClient = String(eData[i][1] || '').trim().toLowerCase();
      const rowCid = String(eData[i][2] || '').trim().toLowerCase();
      const eType = String(eData[i][5] || '').trim().toLowerCase();
      const rawDate = eData[i][6];
      const recip = String(eData[i][7] || '').trim().toLowerCase();
      if (eType !== 'sent' || rowClient !== clientLow || !rawDate) continue;
      const d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') === todayIso) {
        seenSent[rowCid + '::' + recip] = true;
      }
    }
  }

  if (queueData && sentAtCol >= 0 && sendStatusCol >= 0) {
    for (let i = 0; i < queueData.length; i++) {
      const row = queueData[i];
      const st = String(row[sendStatusCol] || '').trim().toLowerCase();
      if (st !== 'sent' && st !== 'pushed to sequencer') continue;
      const rawSentAt = row[sentAtCol];
      if (!rawSentAt) continue;
      const d = (rawSentAt instanceof Date) ? rawSentAt : new Date(rawSentAt);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== todayIso) continue;
      const cid = campIdCol >= 0 ? String(row[campIdCol] || '').trim() : '';
      const rowClient = (typeof resolveCampaignClientName_ === 'function')
        ? resolveCampaignClientName_(cid, '', ss).toLowerCase()
        : 'growth target';
      if (rowClient === clientLow) {
        const email = emailCol >= 0 ? String(row[emailCol] || '').trim().toLowerCase() : ('row-' + i);
        seenSent[cid.toLowerCase() + '::' + email] = true;
      }
    }
  }

  return Object.keys(seenSent).length;
}

/**
 * Resolves an optional client-scoped non-throwing configuration value from Script Properties.
 * Strictly respects Client / Owner boundaries (external clients never fall back to local defaults).
 */
function getOptionalClientSenderProp_(keyCandidates, config, spreadsheet) {
  const clientName = (config && config.clientName) ? String(config.clientName).trim() : 'Growth Target';
  const campaignId = (config && config.campaignId) ? String(config.campaignId).trim() : '';

  if (typeof resolveClientScopedSecret_ === 'function') {
    const res = resolveClientScopedSecret_(
      keyCandidates,
      { campaignId: campaignId, clientName: clientName },
      { throwOnMissing: false, allowAgentSettingsForLocal: true, spreadsheet: spreadsheet }
    );
    if (res && res.key && isValidSenderSecretValue_(res.key)) {
      return String(res.key).trim();
    }
    return '';
  }

  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
  const prefix = (!isLocal && typeof normalizeClientPropertyPrefix_ === 'function')
    ? normalizeClientPropertyPrefix_(clientName)
    : (!isLocal ? ('CLIENT_' + clientName.toUpperCase().replace(/[^A-Z0-9]+/g, '_') + '_') : '');
  const sp = PropertiesService.getScriptProperties().getProperties();

  for (let i = 0; i < keyCandidates.length; i++) {
    const rawKey = isLocal ? keyCandidates[i] : (prefix + keyCandidates[i]);
    const normTarget = String(rawKey).toUpperCase().replace(/[^A-Z0-9]/g, '');
    for (let k in sp) {
      if (isLocal && String(k).toUpperCase().indexOf('CLIENT_') === 0) continue;
      if (k.toUpperCase().replace(/[^A-Z0-9]/g, '') === normTarget && isValidSenderSecretValue_(sp[k])) {
        return String(sp[k]).trim();
      }
    }
  }

  return '';
}

/**
 * Checks whether CASL / B2B Opt-Out Footer injection is enabled for this client.
 * Defaults to true for compliance; can be disabled per-client via [CLIENT_{NAME}_]OPT_OUT_FOOTER_ENABLED = false.
 */
function isClientOptOutFooterEnabled_(config, spreadsheet) {
  const flag = getOptionalClientSenderProp_(
    ['OPT_OUT_FOOTER_ENABLED', 'CASL_FOOTER_ENABLED'],
    config,
    spreadsheet
  );
  if (!flag) return true;
  const low = flag.toLowerCase();
  return !(low === 'false' || low === 'no' || low === '0' || low === 'off' || low === 'disabled');
}

/**
 * Checks if an email body already includes an opt-out or unsubscribe instruction
 * so footers are never appended twice.
 */
function hasExistingOptOutNotice_(bodyText) {
  const text = String(bodyText || '');
  if (!text) return false;
  return /\b(unsubscribe|opt[\s-]?out|remove\s+(?:you|your\s+(?:email|address)|me)\s+from|stop\s+receiving|prefer\s+not\s+to\s+(?:hear|receive)|don't\s+want\s+to\s+receive|future\s+outreach)\b/i.test(text);
}

/**
 * Builds a clean, CASL-aligned B2B sender identification and opt-out footer.
 * Never leaks local agency details into external client emails.
 */
function buildCaslOptOutFooter_(config, spreadsheet) {
  const cfg = config || {};
  const customFooter = getOptionalClientSenderProp_(
    ['OPT_OUT_FOOTER_TEXT', 'CASL_FOOTER_TEXT'],
    cfg,
    spreadsheet
  );
  if (customFooter) {
    return customFooter;
  }

  const senderName = String(cfg.senderName || '').trim();
  const orgName = String(cfg.senderCompany || cfg.clientName || '').trim();
  const contactEmail = String(cfg.replyToEmail || cfg.senderEmail || '').trim();
  const mailingAddress = getOptionalClientSenderProp_(
    ['CASL_MAILING_ADDRESS', 'COMPANY_MAILING_ADDRESS', 'MAILING_ADDRESS'],
    cfg,
    spreadsheet
  );

  const idParts = [];
  if (senderName && orgName && senderName.toLowerCase() !== orgName.toLowerCase()) {
    idParts.push(senderName + ', ' + orgName);
  } else if (orgName) {
    idParts.push(orgName);
  } else if (senderName) {
    idParts.push(senderName);
  }
  if (mailingAddress) {
    idParts.push(mailingAddress);
  }
  if (contactEmail && contactEmail.indexOf('@') > 0) {
    idParts.push(contactEmail);
  }

  const lines = ['--'];
  if (idParts.length > 0) {
    lines.push(idParts.join(' | '));
  }
  lines.push('If this isn\'t relevant or you\'d prefer not to receive future emails from us, simply reply "Unsubscribe" or "Opt out" and we will remove your address promptly.');

  return lines.join('\n');
}

/**
 * Appends the CASL / B2B Opt-Out Footer to outbound copy if enabled and not already present.
 */
function appendCaslOptOutFooter_(bodyText, config, spreadsheet) {
  const rawBody = String(bodyText || '').replace(/\s+$/, '');
  if (!rawBody) return rawBody;
  if (!isClientOptOutFooterEnabled_(config, spreadsheet)) return rawBody;
  if (hasExistingOptOutNotice_(rawBody)) return rawBody;

  const footer = buildCaslOptOutFooter_(config, spreadsheet);
  if (!footer) return rawBody;
  return rawBody + '\n\n' + footer;
}

/**
 * Core outreach processor.
 * Enforces Active Campaign scoping, runtime suppression checks, duplicate-send protection,
 * daily send limits, CASL/opt-out footer compliance, and pacing between dispatches.
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
  if (lastRow < 2) {
    const emptyMsg = 'No prospects found in Prospect Review Queue.';
    if (ui) ui.alert(emptyMsg);
    else Logger.log(emptyMsg);
    return;
  }

  ensureReviewQueueTrackingColumns_(sheet);
  const lastCol = sheet.getLastColumn();
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
  const complianceCol = findSenderHeaderCol_(headers, ['Compliance Status']);
  const suppressionCol = findSenderHeaderCol_(headers, ['Suppression Check']);
  const verifStatusCol = findSenderHeaderCol_(headers, ['Email Verification Status']);

  const activeCampaignId = (typeof getStudioActiveCampaignId_ === 'function')
    ? String(getStudioActiveCampaignId_() || '').trim()
    : '';
  const activeCampLow = activeCampaignId.toLowerCase();

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const suppressionRules = loadSenderSuppressionRules_(spreadsheet);
  const dispatchedLookup = loadAlreadyDispatchedLookup_(spreadsheet, data, campIdCol, emailCol, sendStatusCol);
  const batchSeenKeys = {};
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@', 'office@', 'jobs@', 'careers@'];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  const queueToProcess = [];
  let otherCampaignReadyCount = 0;
  let blockedSuppressedCount = 0;
  let duplicateBlockedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 2;
    const email = emailCol >= 0 ? String(row[emailCol] || '').trim() : '';
    const emailLow = email.toLowerCase();
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

    if (!email || !body || !isApproved || isAlreadySent) continue;

    // 1. Active Campaign Scope Guard
    if (activeCampLow && campaignId.toLowerCase() !== activeCampLow) {
      otherCampaignReadyCount++;
      continue;
    }

    // 2. Runtime Suppression & Invalid-Email Guard
    const compStatus = complianceCol >= 0 ? String(row[complianceCol] || '').trim().toLowerCase() : '';
    const suppCheck = suppressionCol >= 0 ? String(row[suppressionCol] || '').trim().toLowerCase() : '';
    const verifStatus = verifStatusCol >= 0 ? String(row[verifStatusCol] || '').trim().toLowerCase() : '';
    const emailDomain = emailLow.indexOf('@') >= 0 ? emailLow.split('@')[1] : '';

    let blockReason = '';
    if (!emailRegex.test(emailLow)) {
      blockReason = 'Blocked: Invalid email syntax';
    } else if (compStatus.indexOf('failed') >= 0 || suppCheck.indexOf('blocked') >= 0) {
      blockReason = 'Blocked: Compliance/Suppression status failed';
    } else if (verifStatus === 'invalid') {
      blockReason = 'Blocked: Email verification status is invalid';
    } else if (rolePrefixes.some(function(pre) { return emailLow.indexOf(pre) === 0; })) {
      blockReason = 'Blocked: Role-based email prefix';
    } else if (suppressionRules.emails[emailLow]) {
      blockReason = 'Blocked: Email is on active suppression list in Lists & Settings';
    } else if (emailDomain && suppressionRules.domains[emailDomain]) {
      blockReason = 'Blocked: Domain is on active suppression list in Lists & Settings';
    }

    if (blockReason) {
      blockedSuppressedCount++;
      if (sendStatusCol >= 0) sheet.getRange(rowIndex, sendStatusCol + 1).setValue('Suppressed - Not Sent');
      if (sendErrorCol >= 0) sheet.getRange(rowIndex, sendErrorCol + 1).setValue(blockReason);
      continue;
    }

    // 3. Duplicate-Send Protection (per Campaign + Recipient)
    const dedupKey = campaignId.toLowerCase() + '::' + emailLow;
    if (dispatchedLookup[dedupKey] || batchSeenKeys[dedupKey]) {
      duplicateBlockedCount++;
      if (sendStatusCol >= 0) sheet.getRange(rowIndex, sendStatusCol + 1).setValue('Suppressed - Not Sent');
      if (sendErrorCol >= 0) sheet.getRange(rowIndex, sendErrorCol + 1).setValue('Blocked: Duplicate recipient already dispatched in this campaign');
      continue;
    }
    batchSeenKeys[dedupKey] = true;

    queueToProcess.push({
      rowIndex: rowIndex,
      campaignId: campaignId,
      email: email,
      subject: subject,
      body: body,
      leadName: leadNameCol >= 0 ? String(row[leadNameCol] || '') : '',
      companyName: companyCol >= 0 ? String(row[companyCol] || '') : ''
    });
  }

  if (queueToProcess.length === 0) {
    let msg = 'No approved prospects are pending outreach for active campaign "' + (activeCampaignId || 'Active') + '".\n\n';
    if (otherCampaignReadyCount > 0) {
      msg += '• Note: ' + otherCampaignReadyCount + ' approved prospect(s) exist in other campaigns and were safely excluded. Switch campaigns in the Control Panel to process them.\n';
    }
    if (blockedSuppressedCount > 0) {
      msg += '• Blocked by Suppression/Invalid guard: ' + blockedSuppressedCount + '\n';
    }
    if (duplicateBlockedCount > 0) {
      msg += '• Blocked as duplicates already sent in this campaign: ' + duplicateBlockedCount + '\n';
    }
    msg += '\nTo approve a prospect:\n' +
      '1. Set Human Approval Status or Review Status to "Approved"\n' +
      '2. Ensure Send Status is "Queued" or "Draft Failed"\n' +
      '3. Ensure Business Email and Email Subject/Body are filled.';
    SpreadsheetApp.flush();
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

  const primaryCamp = uniqueCampaigns[0];
  const primaryConfig = campaignSettingsCache[primaryCamp] || {};
  const primaryClient = primaryConfig.clientName || 'Growth Target';
  const configuredProv = String(primaryConfig.emailSenderProvider || 'Workspace Gmail (Drafts Only)').trim();
  const effectiveMode = forceDrafts
    ? (isMicrosoft365Provider_(configuredProv.toLowerCase()) ? 'Microsoft 365 (Drafts Only)' : 'Workspace Gmail (Drafts Only)')
    : configuredProv;
  const isLiveSendRun = effectiveMode.toLowerCase().indexOf('draft') < 0;

  const clientSendLimits = {};
  const clientSendsUsedToday = {};
  uniqueCampaigns.forEach(function(cId) {
    const cName = (campaignSettingsCache[cId] && campaignSettingsCache[cId].clientName) || 'Growth Target';
    if (clientSendLimits[cName] === undefined) {
      clientSendLimits[cName] = getClientDailySendLimit_(spreadsheet, cName);
      clientSendsUsedToday[cName] = countClientSendsToday_(
        spreadsheet, cName, data, campIdCol, emailCol, sendStatusCol, sentAtCol
      );
    }
  });

  if (ui) {
    const modeLabel = forceDrafts ? 'Safe Mode: Drafts Only' : ('Configured Channel: ' + effectiveMode);
    let confirmMsg =
      'Mode: ' + modeLabel + '\n' +
      'Active Campaign: ' + (activeCampaignId || uniqueCampaigns.join(', ')) + '\n' +
      'Client / Owner: ' + primaryClient + '\n' +
      (primaryConfig.senderEmail ? ('Sender Mailbox: ' + primaryConfig.senderEmail + '\n') : '') +
      'Prospects to process: ' + queueToProcess.length + '\n';

    if (isLiveSendRun) {
      const used = clientSendsUsedToday[primaryClient] || 0;
      const cap = clientSendLimits[primaryClient];
      confirmMsg += 'Daily Send Limit: ' + used + ' / ' + cap + ' sent today (' + Math.max(0, cap - used) + ' remaining)\n';
    }
    if (otherCampaignReadyCount > 0) {
      confirmMsg += '• Excluded from other campaigns: ' + otherCampaignReadyCount + '\n';
    }
    if (blockedSuppressedCount > 0) {
      confirmMsg += '• Blocked (Suppressed/Invalid): ' + blockedSuppressedCount + '\n';
    }
    if (duplicateBlockedCount > 0) {
      confirmMsg += '• Blocked (Duplicates): ' + duplicateBlockedCount + '\n';
    }

    confirmMsg += '\nDo you want to proceed?';

    const confirm = ui.alert(
      isLiveSendRun ? 'Confirm LIVE Outreach Dispatch' : 'Confirm Safe Draft Creation',
      confirmMsg,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) {
      SpreadsheetApp.flush();
      return;
    }
  }

  let successCount = 0;
  let failCount = 0;
  let pausedByLimitCount = 0;
  let liveSendsInBatch = 0;
  let draftsInBatch = 0;
  const nowStr = new Date().toLocaleString();
  const msTokenCache = {};

  queueToProcess.forEach(function(item) {
    const cId = item.campaignId || 'DEFAULT';
    const campaignConfig = campaignSettingsCache[cId] || getCampaignSenderSettings_(spreadsheet, cId);
    const clientName = campaignConfig.clientName || 'Growth Target';

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

    const lowerProv = rawProvider.toLowerCase();
    const isRowLiveSend = lowerProv.indexOf('draft') < 0;

    // Enforce Daily Send Limit on live sends / sequencer pushes
    if (isRowLiveSend) {
      const dailyLimit = clientSendLimits[clientName] !== undefined
        ? clientSendLimits[clientName]
        : getClientDailySendLimit_(spreadsheet, clientName);
      const usedToday = clientSendsUsedToday[clientName] !== undefined
        ? clientSendsUsedToday[clientName]
        : countClientSendsToday_(spreadsheet, clientName, data, campIdCol, emailCol, sendStatusCol, sentAtCol);

      if (usedToday >= dailyLimit) {
        pausedByLimitCount++;
        if (sendErrorCol >= 0) {
          sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue(
            'Paused: Daily Send Limit (' + dailyLimit + '/day) reached for ' + clientName
          );
        }
        return;
      }

      if (liveSendsInBatch > 0) {
        Utilities.sleep(1500);
      }
    } else if (draftsInBatch > 0) {
      Utilities.sleep(250);
    }

    const compliantBody = appendCaslOptOutFooter_(item.body, campaignConfig, spreadsheet);
    const dispatchItem = {
      rowIndex: item.rowIndex,
      campaignId: item.campaignId,
      email: item.email,
      subject: item.subject,
      body: compliantBody,
      leadName: item.leadName,
      companyName: item.companyName
    };

    try {
      const dispatchResult = dispatchOutreach_(dispatchItem, rawProvider, campaignConfig, spreadsheet, msTokenCache) || {};

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
        const trackCell = sheet.getRange(item.rowIndex, threadMsgIdCol + 1);
        trackCell.setNumberFormat('@');
        trackCell.setValue(trackedId);
      }

      if (draftCreatedAtCol >= 0 && (finalStatus === 'Draft Created' || finalStatus === 'Draft Created in Gmail')) {
        sheet.getRange(item.rowIndex, draftCreatedAtCol + 1).setValue(nowStr);
      }
      if (sentAtCol >= 0 && (finalStatus === 'Sent' || finalStatus === 'Pushed to Sequencer')) {
        sheet.getRange(item.rowIndex, sentAtCol + 1).setValue(nowStr);
      }
      if (sendErrorCol >= 0) sheet.getRange(item.rowIndex, sendErrorCol + 1).setValue('');

      if (isRowLiveSend) {
        liveSendsInBatch++;
        clientSendsUsedToday[clientName] = (clientSendsUsedToday[clientName] || 0) + 1;
      } else {
        draftsInBatch++;
      }

      // Record Sent events inline to EmailEvents for real-time deliverability tracking
      if (finalStatus === 'Sent' &&
          (dispatchResult.provider === 'Gmail' || dispatchResult.provider === 'Microsoft 365') &&
          typeof recordEmailEvent_ === 'function') {
        try {
          recordEmailEvent_({
            client_id: clientName,
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

  let completionMsg = 'Outreach Execution Completed (' + (activeCampaignId || primaryCamp) + ')!\n\n' +
    '• Successfully Processed: ' + successCount + '\n';
  if (failCount > 0) {
    completionMsg += '• Failed: ' + failCount + ' (Check "Send Error" column for details)\n';
  }
  if (pausedByLimitCount > 0) {
    completionMsg += '• Paused by Daily Send Limit: ' + pausedByLimitCount + ' (Remaining rows stay Queued for tomorrow)\n';
  }
  if (blockedSuppressedCount > 0) {
    completionMsg += '• Blocked by Suppression/Invalid Guard: ' + blockedSuppressedCount + '\n';
  }
  if (duplicateBlockedCount > 0) {
    completionMsg += '• Blocked Duplicates: ' + duplicateBlockedCount + '\n';
  }
  if (otherCampaignReadyCount > 0) {
    completionMsg += '• Excluded from other campaigns: ' + otherCampaignReadyCount + '\n';
  }
  if (failCount === 0 && pausedByLimitCount === 0) {
    completionMsg += '• All eligible active-campaign records processed cleanly.';
  }

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

  if (isDraftMode) {
    throw new Error(parseMicrosoftGraphError_(draftCode, draftText, 'create draft', senderEmail));
  }

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