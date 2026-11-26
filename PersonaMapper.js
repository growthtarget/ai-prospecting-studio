/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AGENT 2: PERSONA IDENTIFICATION, RESEARCH & EMAIL CURATION (CLIENT-ISOLATED)
 * ============================================================================
 * File: PersonaMapper.gs / PersonaMapper.js
 *
 * CAPABILITIES:
 * 1. Multi-AI Support: Routes dynamically to GEMINI (Grounding) or OPENAI (GPT-4o)
 *    using strict Client-Scoped API keys (zero local fallback for clients).
 * 2. Strict Verifier Isolation: Resolves ZeroBounce / MillionVerifier / NeverBounce /
 *    Custom HTTPS keys strictly per Client / Owner (validated pre-flight before AI calls).
 * 3. Campaign-Level Anti-Duplication:
 *    - Automatically skips companies in the same Campaign ID that already have mapped personas.
 *    - Deduplicates personas per Campaign ID across normalized LinkedIn URLs,
 *      Business Emails, and normalized Person Name + Company before spending verifier credits.
 * 4. Manual Queue Verification: Supports verifying selected visible rows in
 *    "Prospect Review Queue" on demand and syncing results to "Persona Leads".
 * 5. Real-Time Suppression: Cross-references Lists & Settings for competitor domains,
 *    opt-out emails, and role-based accounts before verification & queuing.
 * 6. Dual-Tab Intelligence: Enriches "Persona Leads" (29 cols) and "Prospect Review Queue" (45 base cols).
 * 7. Grounding Evidence Logging: Persists verified decision-maker signals into "Research Evidence" (9 cols).
 * 8. API Usage & Cost Tracking: Logs LLM token costs and email verifier credit usage to "API_Usage_Log".
 * ============================================================================
 */

const PERSONA_MAPPER_HEADERS_ = {
  companies: ['Campaign ID','Campaign Name','Company ID','Company Name','Country','Industry','Company Size','ESG Frameworks','Reporting Obligations','ESG Reporting Maturity','Supply Chain Profile','Scope 3 Exposure','Sustainability Data Needs','Digital Transformation Signals','ESG Governance','Business Triggers','eValuater Solution Fit','Company Opportunity Score','Company Opportunity','Evidence URLs','Evidence Checked Date','Verification Status','Human Review Status','Last Enriched At','Company Score Rationale'],
  personas: ['Campaign ID','Campaign Name','Company ID','Company Name','Industry','Persona ID','Name','Title','Persona Type','LinkedIn URL','Signal','Signal Date','Evidence URL','Buying Signal','Verification Status','Confidence','Persona Fit Score','Personalized Hook','Review Status','Last Verified At','Business Email','Email Domain','Email Source URL','Email Discovery Method','Email Verification Status','Email Verification Date','Email Verification Provider','Email Identity Confidence','Email Verification Notes'],
  settings: ['Campaign ID','Campaign Name','Persona Mapper Enabled','AI Provider','AI Model','Persona Titles','Search Keywords','Campaign Focus','Evidence Period Months','Research Depth','Include ESG Reporting Obligations','Include ESG Reporting Maturity','Include Supply Chain Complexity','Include Scope 3 Exposure','Include Sustainability Data Challenges','Include Digital Transformation','Include ESG Governance','Include Business Triggers','Include eValuater Solution Fit','Weight ESG Reporting Relevance','Weight Supply Chain Complexity','Weight Sustainability Data Needs','Weight Digital Transformation Relevance','Weight Business Triggers','Weight eValuater Solution Fit','Minimum Company Opportunity Score','Minimum Persona Confidence','Minimum Email Priority Score','Human Approval Required','Auto Send Enabled','Email Priority Weight Company Score','Email Priority Weight Persona Fit','Email Priority Weight Buying Signal','Notes','Max Companies Per Run','Email Verification Provider','Email Verification Endpoint','Email Sender Provider','Email Sender Endpoint'],
  queue: ['Campaign ID','Campaign Name','Company ID','Company Name','Persona ID','Name','Title','Business Email','Company Opportunity Score','Persona Fit Score','Buying Signal','Priority Score','Priority Tier','Priority Score Rationale','Draft Subject','Draft Email','Evidence URLs','Review Status','Reviewer','Reviewer Notes','Reviewed At','Human Approval Status','Approved By','Approved At','Send Status','Sent At','Send Error','Country','Industry','Company Opportunity','Persona Type','LinkedIn URL','Email Domain','Email Source URL','Email Discovery Method','Email Verification Status','Email Verification Date','Email Verification Provider','Email Identity Confidence','Email Verification Notes','Queue Status','Compliance Status','Suppression Check','Date Added','Personalized Hook']
};

const PERSONA_MAPPER_SHEETS_ = {
  companies: 'Company Intelligence',
  personas: 'Persona Leads',
  settings: 'Persona Mapper Settings',
  queue: 'Prospect Review Queue',
  evidence: 'Research Evidence'
};

/**
 * Validates budget cap & external client verifier keys BEFORE burning paid AI tokens.
 */
function validatePersonaPreRunGuardrails_(spreadsheet, campaignId, settings) {
  const clientName = (settings && settings['Department']) || 'Growth Target';

  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(campaignId, clientName);
  }

  const verifierProvider = String((settings && settings['Email Verification Provider']) || 'ZeroBounce').trim();
  if (verifierProvider !== 'None') {
    resolveVerifierApiKey_(spreadsheet, settings);
    if (verifierProvider.toLowerCase().indexOf('custom') >= 0) {
      const endpoint = String((settings && settings['Email Verification Endpoint']) || '').trim();
      const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
      if (!isLocal && (!endpoint || !/^https:\/\//i.test(endpoint))) {
        throw new Error(
          'Custom HTTPS email verifier for "' + clientName + '" requires a valid https:// Email Verification Endpoint before running Persona Mapper.'
        );
      }
    }
  }
}

/**
 * Manual Queue Email Verifier:
 * Verifies selected visible rows in "Prospect Review Queue" outside the AI agent workflow
 * using the campaign's configured verifier (ZeroBounce, MillionVerifier, NeverBounce, Custom HTTPS),
 * syncs results to "Persona Leads", auto-suppresses invalid emails, and warns if an invalid
 * email already has a staged draft in Gmail / Microsoft 365.
 */
function verifySelectedProspectEmails() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  const qSheet = ss.getActiveSheet ? ss.getActiveSheet() : null;
  if (!qSheet || qSheet.getName() !== PERSONA_MAPPER_SHEETS_.queue) {
    const msg = 'Please switch to the "Prospect Review Queue" tab and select one or more visible prospect rows before running email verification.';
    if (ui) ui.alert(msg);
    else Logger.log(msg);
    return;
  }

  const lastRow = qSheet.getLastRow();
  const lastCol = qSheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) {
    const emptyMsg = 'No prospect rows found in Prospect Review Queue.';
    if (ui) ui.alert(emptyMsg);
    else Logger.log(emptyMsg);
    return;
  }

  let ranges = [];
  try {
    const rangeList = qSheet.getActiveRangeList ? qSheet.getActiveRangeList() : null;
    if (rangeList) ranges = rangeList.getRanges() || [];
  } catch (e) {}
  if (!ranges || ranges.length === 0) {
    try {
      const single = qSheet.getActiveRange ? qSheet.getActiveRange() : null;
      if (single) ranges = [single];
    } catch (e) {}
  }
  if (!ranges || ranges.length === 0) {
    try {
      const activeCell = qSheet.getActiveCell ? qSheet.getActiveCell() : null;
      if (activeCell) ranges = [activeCell];
    } catch (e) {}
  }

  const allRows = qSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowSet = {};

  for (let i = 0; i < ranges.length; i++) {
    const rng = ranges[i];
    if (!rng) continue;
    const startR = Math.max(2, rng.getRow());
    const endR = Math.min(lastRow, rng.getRow() + rng.getNumRows() - 1);
    for (let r = startR; r <= endR; r++) {
      if (rowSet[r]) continue;
      let isHidden = false;
      try {
        if ((typeof qSheet.isRowHiddenByFilter === 'function' && qSheet.isRowHiddenByFilter(r)) ||
            (typeof qSheet.isRowHiddenByUser === 'function' && qSheet.isRowHiddenByUser(r))) {
          isHidden = true;
        }
      } catch (e) {}
      if (isHidden) continue;

      const rowPreview = (allRows[r - 2] || []).slice(0, Math.min(lastCol, 18));
      const hasData = rowPreview.some(function(cell) {
        return cell !== '' && cell !== null && cell !== undefined;
      });
      if (hasData) rowSet[r] = true;
    }
  }

  const targetRows = Object.keys(rowSet).map(Number).sort(function(a, b) { return a - b; });
  if (targetRows.length === 0) {
    const selMsg = 'Please select one or more visible prospect rows (row 2 or below) in Prospect Review Queue.';
    if (ui) ui.alert(selMsg);
    else Logger.log(selMsg);
    return;
  }

  const headers = qSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  const normHeaders = headers.map(function(h) {
    return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  });

  function colIdx(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = headers.indexOf(candidates[i]);
      if (idx >= 0) return idx;
    }
    for (let i = 0; i < candidates.length; i++) {
      const normTarget = String(candidates[i]).toUpperCase().replace(/[^A-Z0-9]/g, '');
      const nIdx = normHeaders.indexOf(normTarget);
      if (nIdx >= 0) return nIdx;
    }
    return -1;
  }

  const campCol = colIdx(['Campaign ID']);
  const personaIdCol = colIdx(['Persona ID']);
  const nameCol = colIdx(['Name', 'Persona Name']);
  const emailCol = colIdx(['Business Email', 'Email']);
  const domainCol = colIdx(['Email Domain']);
  const reviewCol = colIdx(['Review Status']);
  const approvalCol = colIdx(['Human Approval Status']);
  const sendStatusCol = colIdx(['Send Status']);
  const sendErrorCol = colIdx(['Send Error', 'Last Error']);
  const verifStatusCol = colIdx(['Email Verification Status']);
  const verifDateCol = colIdx(['Email Verification Date']);
  const verifProvCol = colIdx(['Email Verification Provider']);
  const verifNotesCol = colIdx(['Email Verification Notes']);
  const queueStatusCol = colIdx(['Queue Status']);
  const complianceCol = colIdx(['Compliance Status']);
  const suppressionCol = colIdx(['Suppression Check']);

  if (emailCol < 0 || verifStatusCol < 0) {
    const schemaMsg = 'Required columns ("Business Email" / "Email Verification Status") not found in Prospect Review Queue.';
    if (ui) ui.alert(schemaMsg);
    else Logger.log(schemaMsg);
    return;
  }

  const fallbackActiveCid = typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : '';
  const settingsCache = {};
  const uniqueCampaignIds = [];

  targetRows.forEach(function(rNum) {
    const row = allRows[rNum - 2] || [];
    const cid = (campCol >= 0 ? String(row[campCol] || '').trim() : '') || fallbackActiveCid || 'DEFAULT';
    if (!settingsCache[cid]) {
      settingsCache[cid] = getPersonaMapperSettings_(ss, cid);
      uniqueCampaignIds.push(cid);
    }
  });

  // Pre-flight guardrails before making any API call
  try {
    uniqueCampaignIds.forEach(function(cid) {
      const st = settingsCache[cid];
      const clientName = st['Department'] || 'Growth Target';
      const prov = String(st['Email Verification Provider'] || 'ZeroBounce').trim();

      if (prov.toLowerCase() === 'none') {
        throw new Error(
          'Campaign "' + cid + '" currently has Email Verification Provider set to "None".\n\n' +
          'Open the Campaign Control Panel (Section 5), select ZeroBounce (or your verifier), click "Save Campaign Settings", and try again.'
        );
      }

      validatePersonaPreRunGuardrails_(ss, cid, st);

      const resolvedKey = resolveVerifierApiKey_(ss, st);
      if (!resolvedKey) {
        const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
        const propHint = isLocal ? 'VERIFIER_API_KEY (or ZEROBOUNCE_API_KEY)' : ('CLIENT_' + clientName.toUpperCase().replace(/[^A-Z0-9]+/g, '_') + '_VERIFIER_API_KEY');
        throw new Error(
          'No ' + prov + ' API key found for "' + clientName + '".\n\n' +
          'Add ' + propHint + ' in Extensions > Apps Script > Project Settings > Script Properties.'
        );
      }
    });
  } catch (preErr) {
    if (ui) ui.alert('Verification Pre-Check Blocked', preErr.message, ui.ButtonSet.OK);
    else Logger.log('Verification Pre-Check Blocked: ' + preErr.message);
    return;
  }

  const primarySettings = settingsCache[uniqueCampaignIds[0]] || {};
  const primaryProvider = String(primarySettings['Email Verification Provider'] || 'ZeroBounce').trim();
  const primaryClient = String(primarySettings['Department'] || 'Growth Target').trim();

  if (ui) {
    const confirm = ui.alert(
      'Verify Selected Prospect Emails',
      'Verify ' + targetRows.length + ' selected visible row(s) using ' + primaryProvider + '?\n' +
      '• Client / Owner: ' + primaryClient + '\n' +
      '• Campaign(s): ' + uniqueCampaignIds.join(', ') + '\n\n' +
      'This will use 1 verification credit per valid-syntax email and update both Prospect Review Queue & Persona Leads.',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  const todayIso = new Date().toISOString().split('T')[0];
  let validCount = 0;
  let softWarnCount = 0;
  let invalidCount = 0;
  let errorCount = 0;
  const invalidEmailsToSuppress = [];
  const staleDraftWarnings = [];
  const personaSyncUpdates = [];
  const checkedCampaignClients = {};

  targetRows.forEach(function(rowNum) {
    const row = allRows[rowNum - 2] || [];
    const cid = (campCol >= 0 ? String(row[campCol] || '').trim() : '') || fallbackActiveCid || 'DEFAULT';
    const settings = settingsCache[cid];
    const clientName = settings['Department'] || 'Growth Target';
    checkedCampaignClients[cid] = clientName;

    const rawEmail = String(row[emailCol] || '').trim().toLowerCase();
    const personaId = personaIdCol >= 0 ? String(row[personaIdCol] || '').trim() : '';
    const leadName = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
    const currSendStatus = sendStatusCol >= 0 ? String(row[sendStatusCol] || '').trim() : '';
    const currSendLow = currSendStatus.toLowerCase();
    const hadStagedDraft = (currSendLow === 'draft created' || currSendLow === 'draft created in gmail');
    const isAlreadyDispatched = (currSendLow === 'sent' || currSendLow === 'pushed to sequencer');

    const atIdx = rawEmail.lastIndexOf('@');
    const domain = atIdx > 0 ? rawEmail.slice(atIdx + 1) : '';
    const isValidSyntax = Boolean(domain && /^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(rawEmail));
    const isPersonal = Boolean(domain && isPersonalEmailDomain_(domain));

    let normStatus = 'unverified';
    let providerUsed = String(settings['Email Verification Provider'] || 'ZeroBounce').trim();
    let notesText = '';
    let isApiError = false;

    if (!rawEmail || !isValidSyntax || isPersonal) {
      normStatus = 'invalid';
      notesText = !rawEmail
        ? 'Missing Business Email.'
        : (isPersonal ? 'Personal domain rejected; business email required.' : 'Email syntax invalid.');
    } else {
      const verif = callConfiguredEmailVerifier_(rawEmail, settings, ss);
      providerUsed = verif.provider || providerUsed;
      normStatus = normalizeEmailVerificationStatus_(verif.configured ? verif.status : 'unverified');
      notesText = verif.notes || '';

      const nLow = notesText.toLowerCase();
      if (!verif.configured || nLow.indexOf('verification request error') >= 0 || nLow.indexOf(' status ') >= 0 || nLow.indexOf(' http ') >= 0) {
        if (normStatus === 'unverified') {
          isApiError = true;
        }
      }
    }

    if (domainCol >= 0 && domain) {
      qSheet.getRange(rowNum, domainCol + 1).setValue(domain);
    }
    qSheet.getRange(rowNum, verifStatusCol + 1).setValue(normStatus);
    if (verifDateCol >= 0) qSheet.getRange(rowNum, verifDateCol + 1).setValue(todayIso);
    if (verifProvCol >= 0 && providerUsed) qSheet.getRange(rowNum, verifProvCol + 1).setValue(providerUsed);

    if (isApiError) {
      errorCount++;
      if (verifNotesCol >= 0) qSheet.getRange(rowNum, verifNotesCol + 1).setValue(notesText);
    } else if (normStatus === 'invalid') {
      invalidCount++;
      if (verifNotesCol >= 0) qSheet.getRange(rowNum, verifNotesCol + 1).setValue(notesText);
      if (queueStatusCol >= 0) qSheet.getRange(rowNum, queueStatusCol + 1).setValue('Email Verification Failed');

      const isPilot = cid.toUpperCase().indexOf('PILOT-') === 0;
      if (isValidSyntax && !isPersonal && !isPilot) {
        invalidEmailsToSuppress.push(rawEmail);
        if (complianceCol >= 0) qSheet.getRange(rowNum, complianceCol + 1).setValue('Failed - Suppressed');
        if (suppressionCol >= 0) qSheet.getRange(rowNum, suppressionCol + 1).setValue('Blocked: Email in Lists & Settings');
      } else {
        if (complianceCol >= 0) qSheet.getRange(rowNum, complianceCol + 1).setValue('Failed - Invalid Syntax');
        if (suppressionCol >= 0) qSheet.getRange(rowNum, suppressionCol + 1).setValue('Blocked: Bad Email Syntax');
      }

      if (!isAlreadyDispatched) {
        if (reviewCol >= 0) qSheet.getRange(rowNum, reviewCol + 1).setValue('Rejected');
        if (approvalCol >= 0) qSheet.getRange(rowNum, approvalCol + 1).setValue('Rejected (Suppression Guard)');
        if (sendStatusCol >= 0) qSheet.getRange(rowNum, sendStatusCol + 1).setValue('Suppressed - Not Sent');
      }

      if (sendErrorCol >= 0) {
        const errNote = 'Blocked: ' + providerUsed + ' reported status is invalid' +
          (hadStagedDraft ? ' — DELETE staged draft from mailbox!' : '');
        qSheet.getRange(rowNum, sendErrorCol + 1).setValue(errNote);
      }

      if (hadStagedDraft) {
        staleDraftWarnings.push('• Row ' + rowNum + ' (' + rawEmail + ')');
      }
    } else if (normStatus === 'catch-all' || normStatus === 'risky') {
      softWarnCount++;
      const warnNote = 'Soft Warning: Server status is ' + normStatus + (notesText ? (' (' + notesText + ')') : '') + '.';
      if (verifNotesCol >= 0) qSheet.getRange(rowNum, verifNotesCol + 1).setValue(warnNote);
      if (queueStatusCol >= 0) {
        const qCurr = String(row[queueStatusCol] || '').trim();
        if (!qCurr || qCurr === 'Email Verification Failed' || qCurr === 'Email Research Required') {
          qSheet.getRange(rowNum, queueStatusCol + 1).setValue('Ready for Review');
        }
      }
      notesText = warnNote;
    } else {
      validCount++;
      if (verifNotesCol >= 0) qSheet.getRange(rowNum, verifNotesCol + 1).setValue(notesText || (providerUsed + ': valid'));
      if (queueStatusCol >= 0) {
        const qCurr = String(row[queueStatusCol] || '').trim();
        if (!qCurr || qCurr === 'Email Verification Failed' || qCurr === 'Email Research Required') {
          qSheet.getRange(rowNum, queueStatusCol + 1).setValue('Ready for Review');
        }
      }
    }

    personaSyncUpdates.push({
      campaignId: cid,
      personaId: personaId,
      email: rawEmail,
      name: leadName,
      status: normStatus,
      date: todayIso,
      provider: providerUsed,
      notes: notesText,
      rejectPersona: (!isApiError && normStatus === 'invalid')
    });
  });

  // Auto-suppress valid-syntax invalid emails into Lists & Settings
  if (invalidEmailsToSuppress.length > 0) {
    if (typeof appendSuppressedEmailsToLists_ === 'function') {
      appendSuppressedEmailsToLists_(ss, invalidEmailsToSuppress, 'Suppression - Email', 'invalid verified email(s)');
    } else {
      suppressInvalidVerifiedEmails_(ss, invalidEmailsToSuppress);
    }
  }

  // Sync verification statuses back to "Persona Leads"
  syncVerifiedEmailToPersonaLeads_(ss, personaSyncUpdates);

  // Trigger low-balance check per verified client/campaign
  Object.keys(checkedCampaignClients).forEach(function(cid) {
    if (typeof checkVerifierBalanceAndAlert_ === 'function') {
      checkVerifierBalanceAndAlert_(cid, checkedCampaignClients[cid]);
    }
  });

  if (typeof updateDashboardForActiveCampaign === 'function' && typeof getStudioActiveCampaignId_ === 'function') {
    try {
      updateDashboardForActiveCampaign(getStudioActiveCampaignId_());
    } catch (dashErr) {
      Logger.log('Dashboard refresh notice after manual verification: ' + dashErr.message);
    }
  }

  SpreadsheetApp.flush();

  let summary = 'Verified ' + targetRows.length + ' selected prospect email(s) via ' + primaryProvider + ':\n\n' +
    '• Valid: ' + validCount + '\n' +
    '• Soft Warning (catch-all / risky): ' + softWarnCount + '\n' +
    '• Invalid (Blocked & Suppressed): ' + invalidCount + '\n';
  if (errorCount > 0) {
    summary += '• Request Errors (Unverified): ' + errorCount + ' (Check Email Verification Notes)\n';
  }
  if (staleDraftWarnings.length > 0) {
    summary += '\n⚠️ ACTION REQUIRED — Existing Drafts Found:\n' +
      'The following invalid email(s) already had a draft created in your mailbox. Please delete/discard them in Gmail / Microsoft 365 Drafts so they are not sent:\n' +
      staleDraftWarnings.join('\n');
  }

  if (ui) ui.alert('Manual Email Verification Complete', summary, ui.ButtonSet.OK);
  else Logger.log(summary);
}

/**
 * Fallback suppression helper when appending invalid verified emails into Lists & Settings.
 */
function suppressInvalidVerifiedEmails_(ss, emails) {
  if (!ss || !Array.isArray(emails) || emails.length === 0) return 0;
  let sheet = ss.getSheetByName('Lists & Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Lists & Settings');
    sheet.appendRow(['List name', 'Value', 'Active status']);
  }
  const existing = {};
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      const listName = String(data[i][0] || '').trim().toLowerCase();
      const val = String(data[i][1] || '').trim().toLowerCase();
      if (val && (listName.indexOf('suppression') >= 0 || listName.indexOf('opt-out') >= 0 || listName.indexOf('blacklist') >= 0)) {
        existing[val] = true;
      }
    }
  }
  const toAdd = [];
  emails.forEach(function(em) {
    const clean = String(em || '').trim().toLowerCase();
    if (clean && !existing[clean]) {
      existing[clean] = true;
      toAdd.push(['Suppression - Email', clean, 'Active']);
    }
  });
  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
    if (typeof sortListsAndSettingsTab_ === 'function') {
      try { sortListsAndSettingsTab_(ss); } catch (e) {}
    }
  }
  return toAdd.length;
}

/**
 * Syncs manual email verification results back into matching "Persona Leads" rows.
 */
function syncVerifiedEmailToPersonaLeads_(ss, updates) {
  if (!ss || !Array.isArray(updates) || updates.length === 0) return;
  const pSheet = ss.getSheetByName(PERSONA_MAPPER_SHEETS_.personas);
  if (!pSheet || pSheet.getLastRow() < 2) return;

  const lastCol = pSheet.getLastColumn();
  const headers = pSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  const campCol = headers.indexOf('Campaign ID');
  const personaIdCol = headers.indexOf('Persona ID');
  const nameCol = headers.indexOf('Name');
  const emailCol = headers.indexOf('Business Email');
  const reviewCol = headers.indexOf('Review Status');
  const verifStatusCol = headers.indexOf('Email Verification Status');
  const verifDateCol = headers.indexOf('Email Verification Date');
  const verifProvCol = headers.indexOf('Email Verification Provider');
  const verifNotesCol = headers.indexOf('Email Verification Notes');

  if (campCol < 0 || verifStatusCol < 0) return;

  const rows = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, lastCol).getValues();
  for (let u = 0; u < updates.length; u++) {
    const upd = updates[u];
    const targetCid = String(upd.campaignId || '').trim().toLowerCase();
    const targetPid = String(upd.personaId || '').trim().toLowerCase();
    const targetEmail = String(upd.email || '').trim().toLowerCase();
    const targetName = String(upd.name || '').trim().toLowerCase();

    for (let r = 0; r < rows.length; r++) {
      const rCid = String(rows[r][campCol] || '').trim().toLowerCase();
      if (rCid !== targetCid) continue;

      const rPid = personaIdCol >= 0 ? String(rows[r][personaIdCol] || '').trim().toLowerCase() : '';
      const rEmail = emailCol >= 0 ? String(rows[r][emailCol] || '').trim().toLowerCase() : '';
      const rName = nameCol >= 0 ? String(rows[r][nameCol] || '').trim().toLowerCase() : '';

      const matched = (targetPid && rPid && rPid === targetPid) ||
        (targetEmail && rEmail && rEmail === targetEmail) ||
        (targetName && rName && rName === targetName);

      if (!matched) continue;

      const rowNum = r + 2;
      pSheet.getRange(rowNum, verifStatusCol + 1).setValue(upd.status);
      if (verifDateCol >= 0) pSheet.getRange(rowNum, verifDateCol + 1).setValue(upd.date);
      if (verifProvCol >= 0 && upd.provider) pSheet.getRange(rowNum, verifProvCol + 1).setValue(upd.provider);
      if (verifNotesCol >= 0) pSheet.getRange(rowNum, verifNotesCol + 1).setValue(upd.notes);
      if (upd.rejectPersona && reviewCol >= 0) {
        pSheet.getRange(rowNum, reviewCol + 1).setValue('Rejected');
      }
    }
  }
}

/**
 * Called directly from the Sidebar Control Panel (Silent execution, returns JSON)
 */
function runPersonaMapperFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'Persona Mapper is already running. Try again after it finishes.' };
  }
  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const activeId = campaignId || (typeof getStudioActiveCampaignId_ === 'function' ? getStudioActiveCampaignId_() : 'CMP-001');
    const settings = getPersonaMapperSettings_(spreadsheet, activeId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) {
      return { success: false, message: 'Persona Mapper is disabled for campaign ' + activeId + '. Enable it in settings.' };
    }

    const clientName = settings['Department'] || 'Growth Target';
    const providerChoice = String(settings['AI Provider'] || 'GEMINI').toUpperCase();
    const modelChoice = resolvePersonaAiModel_(providerChoice, settings['AI Model']);
    const providerLabel = providerChoice === 'OPENAI' ? 'OpenAI (' + modelChoice + ')' : 'Google Gemini (' + modelChoice + ')';

    const allQualifiedCompanies = getPersonaMapperCompanies_(spreadsheet, activeId, settings, { skipAlreadyMapped: false });
    if (!allQualifiedCompanies.length) {
      return { success: false, message: 'No qualified companies found for campaign ' + activeId + ' in Company Database. Run Step 1 first!' };
    }

    const unmappedCompanies = getPersonaMapperCompanies_(spreadsheet, activeId, settings, { skipAlreadyMapped: true });
    const properties = PropertiesService.getDocumentProperties();
    const cursorKey = 'PERSONA_MAPPER_CURSOR_' + activeId;

    if (!unmappedCompanies.length) {
      properties.deleteProperty(cursorKey);
      return {
        success: true,
        count: 0,
        allMapped: true,
        message: 'All ' + allQualifiedCompanies.length + ' qualified company record(s) in campaign ' + activeId +
          ' already have personas mapped in "Persona Leads". Increase Discovery Limit and run Step 1 to add new companies first.'
      };
    }

    // Fail-fast before making any paid AI or Verifier API calls
    validatePersonaPreRunGuardrails_(spreadsheet, activeId, settings);

    const suppressionRules = loadSuppressionRulesFromSettings_(spreadsheet);
    const existingPersonaTracker = getExistingCampaignPersonas_(spreadsheet, activeId);
    const mappedTracker = getMappedCompanyTracker_(spreadsheet, activeId);
    const batchLimit = Math.max(1, Math.min(10, Number(settings['Max Companies Per Run']) || 3));
    const batchCompanies = unmappedCompanies.slice(0, batchLimit);
    let processed = 0;
    let skippedDuplicatePersonas = 0;

    const minCompanyScore = Number(settings['Minimum Company Opportunity Score']) || 0;
    const minConfidenceRank = resolveMinPersonaConfidenceRank_(settings['Minimum Persona Confidence']);

    for (let index = 0; index < batchCompanies.length; index++) {
      const company = batchCompanies[index];
      const companyId = String(company['Company ID'] || '').trim();
      const companyName = String(company['Company name'] || company['Company Name'] || '').trim();
      const normCompName = normalizePersonaCompanyName_(companyName);

      // Also guard against duplicate company rows within the same batch run
      if ((companyId && mappedTracker.companyIds[companyId.toLowerCase()]) ||
          (normCompName && mappedTracker.companyNames[normCompName])) {
        Logger.log('Skipped duplicate/already-mapped company in batch for ' + activeId + ': ' + companyName + ' (' + companyId + ')');
        continue;
      }

      const evidenceSnippets = getCompanyEvidenceSnippets_(spreadsheet, activeId, companyName);

      const aiResponse = executeAiCall_({
        campaignId: activeId,
        clientName: clientName,
        provider: providerChoice,
        model: modelChoice,
        prompt: buildPersonaMapperPrompt_(company, settings, evidenceSnippets),
        enableWebSearch: true,
        json: true,
        operation: 'PersonaMapping'
      });

      const result = aiResponse.json || {};
      const citations = aiResponse.citations || [];
      const companyScore = scoreCompanyOpportunity_(result.company || {}, settings);
      const companyRecord = makeCompanyIntelligenceRecord_(company, result.company || {}, companyScore, citations, settings);

      upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.companies), PERSONA_MAPPER_HEADERS_.companies, 'company', companyRecord);

      const grounding = (typeof extractGroundingMetadata_ === 'function' && aiResponse)
        ? extractGroundingMetadata_(aiResponse.raw || aiResponse)
        : { citationsByUrl: {}, searchQueries: [] };

      const personasList = Array.isArray(result.personas)
        ? result.personas
        : (result.persona && typeof result.persona === 'object' ? [result.persona] : []);

      personasList.forEach(function(persona) {
        const confidence = titleCaseSignal_(persona.confidence) || 'Medium';
        if (confidence === 'Low' || confidenceRank_(confidence) < minConfidenceRank) return;
        if (companyScore.score < minCompanyScore) return;

        // Deduplicate per Campaign ID BEFORE calling paid email verifier
        if (isDuplicatePersonaInCampaign_(persona, company, existingPersonaTracker)) {
          skippedDuplicatePersonas++;
          Logger.log('Skipped duplicate persona in ' + activeId + ': ' + (persona.name || '') + ' @ ' + companyName);
          return;
        }

        const fitScore = scorePersonaFit_(persona.fitDimensions || {});
        const emailInfo = prepareEmailIntelligence_(persona, company, settings, spreadsheet, suppressionRules);

        if (emailInfo.businessEmail && isDuplicatePersonaInCampaign_(persona, company, existingPersonaTracker, emailInfo.businessEmail)) {
          skippedDuplicatePersonas++;
          Logger.log('Skipped duplicate persona email in ' + activeId + ': ' + emailInfo.businessEmail);
          return;
        }

        const personaRecord = makePersonaRecord_(company, persona, fitScore, emailInfo, settings);
        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.personas), PERSONA_MAPPER_HEADERS_.personas, 'persona', personaRecord);
        const queueRecord = makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings, suppressionRules);
        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue), PERSONA_MAPPER_HEADERS_.queue, 'queue', queueRecord);

        registerPersonaInCampaignTracker_(persona, company, existingPersonaTracker, emailInfo);
        recordPersonaGroundingEvidence_(spreadsheet, activeId, company, persona, grounding, settings['Persona Titles']);
      });

      if (result.company && Array.isArray(result.company.evidence)) {
        result.company.evidence.forEach(function(evItem) {
          if (evItem && evItem.url) {
            recordCompanyGroundingEvidence_(spreadsheet, activeId, company, evItem, grounding, settings['Campaign Focus']);
          }
        });
      }

      if (companyId) mappedTracker.companyIds[companyId.toLowerCase()] = true;
      if (normCompName) mappedTracker.companyNames[normCompName] = true;
      processed++;
    }

    if (typeof checkVerifierBalanceAndAlert_ === 'function') {
      checkVerifierBalanceAndAlert_(activeId, clientName);
    }

    if (typeof updateDashboardForActiveCampaign === 'function' &&
        (typeof getStudioActiveCampaignId_ !== 'function' || activeId === getStudioActiveCampaignId_())) {
      updateDashboardForActiveCampaign(activeId);
    }

    properties.deleteProperty(cursorKey);
    const remainingUnmapped = Math.max(0, unmappedCompanies.length - batchCompanies.length);
    const dupNote = skippedDuplicatePersonas > 0
      ? (' (' + skippedDuplicatePersonas + ' duplicate persona(s) automatically skipped)')
      : '';

    if (remainingUnmapped === 0) {
      return {
        success: true,
        count: processed,
        message: 'Persona Mapper completed ' + processed + ' unmapped company record(s) for "' + clientName + '" using ' + providerLabel + '!' + dupNote + ' Leads staged in Review Queue & citations logged in Research Evidence.'
      };
    } else {
      return {
        success: true,
        count: processed,
        message: 'Processed ' + processed + ' unmapped company record(s) for "' + clientName + '" with ' + providerLabel + '.' + dupNote + ' ' + remainingUnmapped + ' unmapped company record(s) remaining—click Step 2 again to run next batch.'
      };
    }
  } catch (err) {
    Logger.log('Persona Mapper Error: ' + err.stack);
    return { success: false, message: 'Persona Mapper error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Called from the Top Menu (Interactive dialog mode with headless safety)
 */
function runPersonaMapperForSelectedCampaign() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    if (ui) ui.alert('Persona Mapper is already running. Try again after it finishes.');
    else Logger.log('Persona Mapper is already running.');
    return;
  }

  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const campaignId = getPersonaMapperCampaignId_(spreadsheet, ui);
    if (!campaignId) return;
    const settings = getPersonaMapperSettings_(spreadsheet, campaignId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) {
      if (ui) ui.alert('Persona Mapper is disabled for this campaign. Enable it in settings.');
      else Logger.log('Persona Mapper is disabled for campaign ' + campaignId);
      return;
    }

    const clientName = settings['Department'] || 'Growth Target';
    const providerChoice = String(settings['AI Provider'] || 'GEMINI').toUpperCase();
    const modelChoice = resolvePersonaAiModel_(providerChoice, settings['AI Model']);
    const providerLabel = providerChoice === 'OPENAI' ? 'OpenAI (' + modelChoice + ')' : 'Google Gemini (' + modelChoice + ')';

    const allQualifiedCompanies = getPersonaMapperCompanies_(spreadsheet, campaignId, settings, { skipAlreadyMapped: false });
    if (!allQualifiedCompanies.length) {
      if (ui) ui.alert('No qualified companies found under campaign ' + campaignId + ' in Company Database yet. Run Step 1 first.');
      else Logger.log('No qualified companies found for campaign ' + campaignId);
      return;
    }

    const unmappedCompanies = getPersonaMapperCompanies_(spreadsheet, campaignId, settings, { skipAlreadyMapped: true });
    const properties = PropertiesService.getDocumentProperties();
    const cursorKey = 'PERSONA_MAPPER_CURSOR_' + campaignId;

    if (!unmappedCompanies.length) {
      properties.deleteProperty(cursorKey);
      const doneMsg = 'All ' + allQualifiedCompanies.length + ' qualified company record(s) in campaign ' + campaignId +
        ' already have personas mapped in "Persona Leads".\n\nIncrease Discovery Limit and run Step 1 to add new companies first.';
      if (ui) ui.alert('All Companies Already Mapped', doneMsg, ui.ButtonSet.OK);
      else Logger.log(doneMsg);
      return;
    }

    // Fail-fast before making any paid AI or Verifier API calls
    try {
      validatePersonaPreRunGuardrails_(spreadsheet, campaignId, settings);
    } catch (preErr) {
      if (ui) ui.alert('Pre-Flight Check Blocked Run', preErr.message, ui.ButtonSet.OK);
      else Logger.log('Pre-Flight Check Blocked Run: ' + preErr.message);
      return;
    }

    const suppressionRules = loadSuppressionRulesFromSettings_(spreadsheet);
    const existingPersonaTracker = getExistingCampaignPersonas_(spreadsheet, campaignId);
    const mappedTracker = getMappedCompanyTracker_(spreadsheet, campaignId);
    const batchLimit = Math.max(1, Math.min(10, Number(settings['Max Companies Per Run']) || 3));
    const batchCompanies = unmappedCompanies.slice(0, batchLimit);
    let processed = 0;
    let skippedDuplicatePersonas = 0;

    const minCompanyScore = Number(settings['Minimum Company Opportunity Score']) || 0;
    const minConfidenceRank = resolveMinPersonaConfidenceRank_(settings['Minimum Persona Confidence']);

    for (let index = 0; index < batchCompanies.length; index++) {
      const company = batchCompanies[index];
      const companyId = String(company['Company ID'] || '').trim();
      const companyName = String(company['Company name'] || company['Company Name'] || '').trim();
      const normCompName = normalizePersonaCompanyName_(companyName);

      if ((companyId && mappedTracker.companyIds[companyId.toLowerCase()]) ||
          (normCompName && mappedTracker.companyNames[normCompName])) {
        Logger.log('Skipped duplicate/already-mapped company in batch for ' + campaignId + ': ' + companyName + ' (' + companyId + ')');
        continue;
      }

      const evidenceSnippets = getCompanyEvidenceSnippets_(spreadsheet, campaignId, companyName);

      try {
        const aiResponse = executeAiCall_({
          campaignId: campaignId,
          clientName: clientName,
          provider: providerChoice,
          model: modelChoice,
          prompt: buildPersonaMapperPrompt_(company, settings, evidenceSnippets),
          enableWebSearch: true,
          json: true,
          operation: 'PersonaMapping'
        });

        const result = aiResponse.json || {};
        const citations = aiResponse.citations || [];
        const companyScore = scoreCompanyOpportunity_(result.company || {}, settings);
        const companyRecord = makeCompanyIntelligenceRecord_(company, result.company || {}, companyScore, citations, settings);

        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.companies), PERSONA_MAPPER_HEADERS_.companies, 'company', companyRecord);

        const grounding = (typeof extractGroundingMetadata_ === 'function' && aiResponse)
          ? extractGroundingMetadata_(aiResponse.raw || aiResponse)
          : { citationsByUrl: {}, searchQueries: [] };

        const personasList = Array.isArray(result.personas)
          ? result.personas
          : (result.persona && typeof result.persona === 'object' ? [result.persona] : []);

        personasList.forEach(function(persona) {
          const confidence = titleCaseSignal_(persona.confidence) || 'Medium';
          if (confidence === 'Low' || confidenceRank_(confidence) < minConfidenceRank) return;
          if (companyScore.score < minCompanyScore) return;

          if (isDuplicatePersonaInCampaign_(persona, company, existingPersonaTracker)) {
            skippedDuplicatePersonas++;
            Logger.log('Skipped duplicate persona in ' + campaignId + ': ' + (persona.name || '') + ' @ ' + companyName);
            return;
          }

          const fitScore = scorePersonaFit_(persona.fitDimensions || {});
          const emailInfo = prepareEmailIntelligence_(persona, company, settings, spreadsheet, suppressionRules);

          if (emailInfo.businessEmail && isDuplicatePersonaInCampaign_(persona, company, existingPersonaTracker, emailInfo.businessEmail)) {
            skippedDuplicatePersonas++;
            Logger.log('Skipped duplicate persona email in ' + campaignId + ': ' + emailInfo.businessEmail);
            return;
          }

          const personaRecord = makePersonaRecord_(company, persona, fitScore, emailInfo, settings);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.personas), PERSONA_MAPPER_HEADERS_.personas, 'persona', personaRecord);
          const queueRecord = makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings, suppressionRules);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue), PERSONA_MAPPER_HEADERS_.queue, 'queue', queueRecord);

          registerPersonaInCampaignTracker_(persona, company, existingPersonaTracker, emailInfo);
          recordPersonaGroundingEvidence_(spreadsheet, campaignId, company, persona, grounding, settings['Persona Titles']);
        });

        if (result.company && Array.isArray(result.company.evidence)) {
          result.company.evidence.forEach(function(evItem) {
            if (evItem && evItem.url) {
              recordCompanyGroundingEvidence_(spreadsheet, campaignId, company, evItem, grounding, settings['Campaign Focus']);
            }
          });
        }

        if (companyId) mappedTracker.companyIds[companyId.toLowerCase()] = true;
        if (normCompName) mappedTracker.companyNames[normCompName] = true;
        processed++;
      } catch (error) {
        Logger.log('Persona Mapper paused on company: ' + error.message);
        if (ui) ui.alert('Persona Mapper paused: ' + error.message);
        return;
      }
    }

    if (typeof checkVerifierBalanceAndAlert_ === 'function') {
      checkVerifierBalanceAndAlert_(campaignId, clientName);
    }

    if (typeof updateDashboardForActiveCampaign === 'function' &&
        (typeof getStudioActiveCampaignId_ !== 'function' || campaignId === getStudioActiveCampaignId_())) {
      updateDashboardForActiveCampaign(campaignId);
    }

    properties.deleteProperty(cursorKey);
    const remainingUnmapped = Math.max(0, unmappedCompanies.length - batchCompanies.length);
    const dupNote = skippedDuplicatePersonas > 0
      ? (' (' + skippedDuplicatePersonas + ' duplicate persona(s) automatically skipped)')
      : '';

    if (remainingUnmapped === 0) {
      if (ui) ui.alert('Persona Mapper completed ' + processed + ' unmapped company record(s) for "' + clientName + '" using ' + providerLabel + '.' + dupNote + ' Leads staged in Prospect Review Queue & Grounding Evidence recorded.');
    } else {
      if (ui) ui.alert('Persona Mapper processed ' + processed + ' unmapped company record(s) for "' + clientName + '".' + dupNote + ' ' + remainingUnmapped + ' unmapped company record(s) remaining—run it again to continue next batch.');
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Queries Tab 10 (Research Evidence) for existing evidence snippets & email pattern clues
 */
function getCompanyEvidenceSnippets_(spreadsheet, campaignId, companyName) {
  const snippets = [];
  if (!spreadsheet || !companyName) return snippets;

  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');
  if (!evidenceSheet || evidenceSheet.getLastRow() <= 1) return snippets;

  const data = evidenceSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCampId = String(data[i][1] || '').trim();
    const entityName = String(data[i][2] || '').trim().toLowerCase();
    const snippet = String(data[i][6] || '').trim();
    const url = String(data[i][5] || '').trim();

    if (rowCampId === String(campaignId).trim() && entityName.includes(companyName.toLowerCase())) {
      if (snippet) {
        snippets.push('• ' + snippet + (url ? ' (Source: ' + url + ')' : ''));
      }
    }
  }
  return snippets;
}

/**
 * Builds the Multi-Client Persona Research & Outreach Prompt
 */
function buildPersonaMapperPrompt_(company, settings, evidenceSnippets) {
  const clientOwner = settings['Department'] || 'Growth Target';
  const senderName = settings['Sender Name'] || 'Alex Mercer';
  const senderTitle = settings['Sender Title'] || 'Solutions Consultant';
  const senderCompany = settings['Sender Company'] || 'eValuater Solutions';
  const customNeeds = settings['Custom Business Needs'] || settings['Campaign Focus'] || settings['Notes'] || 'operational scalability and automated compliance';
  const snippetsText = (evidenceSnippets && evidenceSnippets.length > 0)
    ? evidenceSnippets.slice(0, 8).join('\n')
    : 'No prior company research snippets recorded yet. Perform fresh discovery.';

  return [
    'Act as the Senior B2B Account Researcher, Executive Persona Mapper, and Outbound Email Copywriter.',
    'Research ONLY the single existing company supplied below. Do not invent or discover other companies.',
    '',
    'OPERATIONAL CONTEXT & SENDER IDENTITY:',
    '• Client / Account Owner: ' + clientOwner,
    '• Sender Name: ' + senderName,
    '• Sender Title: ' + senderTitle,
    '• Sender Company: ' + senderCompany,
    '• Core Business Pain / Buying Triggers to Target: ' + customNeeds,
    '• Research Depth: ' + (settings['Research Depth'] || 'Deep Persona') + '.',
    '• Target Persona Titles: ' + (settings['Persona Titles'] || 'VP, Director, Head of Department, Chief Officer'),
    '',
    'EXISTING RESEARCH EVIDENCE & CLUES (FROM RESEARCH EVIDENCE TAB):',
    snippetsText,
    '',
    'COMPANY PROFILE UNDER RESEARCH:',
    JSON.stringify(company),
    '',
    'RESEARCH INSTRUCTIONS:',
    '1. Identify up to ' + Math.max(1, Math.min(5, Number(settings['Personas Per Company']) || 2)) + ' real decision-makers matching the target roles using live web search when available or verified public knowledge.',
    '2. Separate verified facts from interpretation. Never invent a person, title, recent activity, or email address.',
    '3. Search for or identify the company domain and likely professional business email pattern.',
    '',
    'EMAIL CURATION INSTRUCTIONS (IMPORTANT):',
    '• Draft a peer-to-peer cold outreach message under 150 words from ' + senderName + ' (' + senderTitle + ' at ' + senderCompany + ').',
    '• Open with a personalized hook referencing the decision-maker\'s specific initiative or operational context.',
    '• Frame how ' + senderCompany + ' eliminates the specific pain point (' + customNeeds.slice(0, 150) + ') without slowing sprint cycles.',
    '• Propose a brief, low-friction 7-minute introductory call next week.',
    '• Sign off with ' + senderName + ', ' + senderTitle + ' | ' + senderCompany + '.',
    '',
    'Return ONLY a valid JSON object matching this schema:',
    '{',
    '  "company": {',
    '    "country": "",',
    '    "esgFrameworks": "",',
    '    "reportingObligations": "",',
    '    "reportingMaturity": "One of: Early Stage, Developing / Spreadsheet-reliant, Developing, Mature, Advanced, or N/A",',
    '    "supplyChainProfile": "",',
    '    "scope3Exposure": "",',
    '    "sustainabilityDataNeeds": "",',
    '    "digitalTransformationSignals": "",',
    '    "esgGovernance": "",',
    '    "businessTriggers": "",',
    '    "solutionFit": "One of: High, Medium, Low, or Disqualified - Direct Competitor",',
    '    "dimensions": {',
    '      "esgReportingRelevance": 0,',
    '      "supplyChainComplexity": 0,',
    '      "sustainabilityDataNeeds": 0,',
    '      "digitalTransformationRelevance": 0,',
    '      "businessTriggers": 0,',
    '      "solutionFit": 0',
    '    },',
    '    "evidence": [{"url": "", "claim": ""}]',
    '  },',
    '  "personas": [',
    '    {',
    '      "name": "Full Name",',
    '      "title": "Exact Current Job Title",',
    '      "personaType": "Economic Buyer, Key Decision Maker, or Technical Evaluator",',
    '      "linkedinUrl": "https://linkedin.com/in/...",',
    '      "signal": "Verified recent keynote, post, publication, or role responsibility",',
    '      "signalDate": "YYYY-MM-DD",',
    '      "evidenceUrl": "Direct source URL confirming this person and signal",',
    '      "buyingSignal": "Active, Relevant, or Emerging",',
    '      "verificationStatus": "Verified",',
    '      "confidence": "High, Medium, or Low",',
    '      "fitDimensions": {',
    '        "decisionAuthority": 85,',
    '        "functionalRelevance": 90,',
    '        "opportunityConnection": 85,',
    '        "recentActivity": 80,',
    '        "evidenceQuality": 85',
    '      },',
    '      "personalizedHook": "Specific opening hook referencing their context",',
    '      "draftSubject": "Relevant and concise subject line",',
    '      "draftEmail": "Full curated email under 150 words",',
    '      "businessEmail": "username@companydomain.com",',
    '      "emailDomain": "companydomain.com",',
    '      "emailSourceUrl": "https://...",',
    '      "emailDiscoveryMethod": "Published, Public biography, Pattern Match, or Not Found",',
    '      "emailIdentityConfidence": "High, Medium, or Low"',
    '    }',
    '  ]',
    '}'
  ].join('\n');
}

function resolvePersonaAiModel_(provider, configuredModel) {
  if (configuredModel && String(configuredModel).trim() !== '') {
    return String(configuredModel).trim();
  }
  return provider === 'OPENAI' ? 'gpt-4o' : 'gemini-3.8-flash';
}

function getPersonaMapperSettings_(spreadsheet, campaignId) {
  const settingsSheet = spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.settings);
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  const campaignsSheet = spreadsheet.getSheetByName('Campaigns');
  const agentSheet = spreadsheet.getSheetByName('Agent Settings');

  const merged = {
    'Campaign ID': campaignId || 'CMP-001',
    'Campaign Name': '',
    'Persona Mapper Enabled': true,
    'AI Provider': '',
    'AI Model': '',
    'Department': '',
    'Persona Titles': '',
    'Search Keywords': '',
    'Campaign Focus': '',
    'Custom Business Needs': '',
    'Sender Name': 'Alex Mercer',
    'Sender Title': 'Solutions Consultant',
    'Sender Company': 'eValuater Solutions',
    'Research Depth': 'Deep Persona',
    'Personas Per Company': 2,
    'Max Companies Per Run': 3,
    'Minimum Company Opportunity Score': 75,
    'Minimum Persona Confidence': 85,
    'Minimum Email Priority Score': 80,
    'Email Verification Provider': 'ZeroBounce',
    'Email Verification Endpoint': '',
    'Email Sender Provider': 'Workspace Gmail (Drafts Only)',
    'Email Sender Endpoint': '',
    'Include ESG Reporting Obligations': true,
    'Include ESG Reporting Maturity': true,
    'Include Supply Chain Complexity': true,
    'Include Scope 3 Exposure': true,
    'Include Sustainability Data Challenges': true,
    'Include Digital Transformation': true,
    'Include ESG Governance': true,
    'Include Business Triggers': true,
    'Include eValuater Solution Fit': true,
    'Weight ESG Reporting Relevance': 20,
    'Weight Supply Chain Complexity': 20,
    'Weight Sustainability Data Needs': 20,
    'Weight Digital Transformation Relevance': 15,
    'Weight Business Triggers': 15,
    'Weight eValuater Solution Fit': 10
  };

  // 1. Read Client / Owner & Sender fields from Campaigns master table by Campaign ID
  if (campaignsSheet && campaignsSheet.getLastRow() > 1) {
    const campRecords = getPersonaMapperSheetRecords_(campaignsSheet);
    const campMatch = campRecords.find(function(row) {
      return String(row['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();
    });
    if (campMatch) {
      if (campMatch['Campaign Config JSON']) {
        try {
          const parsed = JSON.parse(campMatch['Campaign Config JSON']);
          if (parsed && typeof parsed === 'object') {
            if (parsed.department) merged['Department'] = String(parsed.department).trim();
            if (parsed.customBusinessNeeds) merged['Custom Business Needs'] = String(parsed.customBusinessNeeds).trim();
            if (parsed.senderName) merged['Sender Name'] = String(parsed.senderName).trim();
            if (parsed.senderTitle) merged['Sender Title'] = String(parsed.senderTitle).trim();
            if (parsed.senderCompany) merged['Sender Company'] = String(parsed.senderCompany).trim();
          }
        } catch (e) {}
      }
      if (campMatch['Campaign Name']) merged['Campaign Name'] = String(campMatch['Campaign Name']).trim();
      if (campMatch['Department']) merged['Department'] = String(campMatch['Department']).trim();
      if (campMatch['Custom Business Needs']) merged['Custom Business Needs'] = String(campMatch['Custom Business Needs']).trim();
      if (campMatch['Sender Name']) merged['Sender Name'] = String(campMatch['Sender Name']).trim();
      if (campMatch['Sender Title']) merged['Sender Title'] = String(campMatch['Sender Title']).trim();
      if (campMatch['Sender Company']) merged['Sender Company'] = String(campMatch['Sender Company']).trim();
      if (campMatch['AI Provider']) merged['AI Provider'] = String(campMatch['AI Provider']).trim().toUpperCase();
      if (campMatch['AI Model']) merged['AI Model'] = String(campMatch['AI Model']).trim();
      if (campMatch['Email Verifier Provider']) merged['Email Verification Provider'] = String(campMatch['Email Verifier Provider']).trim();
      if (campMatch['Email Verifier Endpoint']) merged['Email Verification Endpoint'] = String(campMatch['Email Verifier Endpoint']).trim();
      if (campMatch['Email Sender Provider']) merged['Email Sender Provider'] = String(campMatch['Email Sender Provider']).trim();
      if (campMatch['Email Sender Endpoint']) merged['Email Sender Endpoint'] = String(campMatch['Email Sender Endpoint']).trim();
    }
  }

  // 2. Read from Campaign Configuration when Campaign ID matches
  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    const dict = {};
    for (let r = 1; r < data.length; r++) {
      const k = String(data[r][0] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (k) dict[k] = data[r][1];
    }
    const activeConfigId = String(dict['CAMPAIGNID'] || '').trim();
    if (!activeConfigId || activeConfigId.toLowerCase() === String(campaignId).trim().toLowerCase()) {
      if (dict['CAMPAIGNNAME'] || dict['NAME']) merged['Campaign Name'] = String(dict['CAMPAIGNNAME'] || dict['NAME']);
      if (dict['DEPARTMENT'] || dict['DEPARTMENTOWNER'] || dict['CLIENTOWNER'] || dict['CLIENT']) {
        merged['Department'] = String(dict['DEPARTMENT'] || dict['DEPARTMENTOWNER'] || dict['CLIENTOWNER'] || dict['CLIENT']).trim();
      }
      if (dict['AIPROVIDER'] || dict['PRIMARYAIPROVIDER']) merged['AI Provider'] = String(dict['AIPROVIDER'] || dict['PRIMARYAIPROVIDER']).trim().toUpperCase();
      if (dict['AIMODEL'] || dict['OPENAIMODEL'] || dict['GEMINIMODEL']) merged['AI Model'] = String(dict['AIMODEL'] || dict['OPENAIMODEL'] || dict['GEMINIMODEL']).trim();
      if (dict['TARGETTITLES'] || dict['TARGETPERSONATITLES']) merged['Persona Titles'] = String(dict['TARGETTITLES'] || dict['TARGETPERSONATITLES']);
      if (dict['SEARCHKEYWORDS']) merged['Search Keywords'] = String(dict['SEARCHKEYWORDS']);
      if (dict['CAMPAIGNFOCUS']) merged['Campaign Focus'] = String(dict['CAMPAIGNFOCUS']);
      if (dict['CUSTOMBUSINESSNEEDS']) merged['Custom Business Needs'] = String(dict['CUSTOMBUSINESSNEEDS']);
      if (dict['SENDERNAME']) merged['Sender Name'] = String(dict['SENDERNAME']);
      if (dict['SENDERTITLE']) merged['Sender Title'] = String(dict['SENDERTITLE']);
      if (dict['SENDERCOMPANY']) merged['Sender Company'] = String(dict['SENDERCOMPANY']);
      if (dict['RESEARCHDEPTH']) merged['Research Depth'] = String(dict['RESEARCHDEPTH']);
      if (dict['MINCOMPANYSCORE'] || dict['MINCOMPANYFITSCORE']) merged['Minimum Company Opportunity Score'] = Number(dict['MINCOMPANYSCORE'] || dict['MINCOMPANYFITSCORE']);
      if (dict['MINPERSONASCORE'] || dict['MINPERSONAFITSCORE']) merged['Minimum Persona Confidence'] = Number(dict['MINPERSONASCORE'] || dict['MINPERSONAFITSCORE']);
      if (dict['MINEMAILPRIORITYSCORE']) merged['Minimum Email Priority Score'] = Number(dict['MINEMAILPRIORITYSCORE']);
      if (dict['MAXCOMPANIES']) merged['Max Companies Per Run'] = Number(dict['MAXCOMPANIES']);
      if (dict['EMAILVERIFICATIONPROVIDER']) merged['Email Verification Provider'] = String(dict['EMAILVERIFICATIONPROVIDER']);
      if (dict['EMAILVERIFICATIONENDPOINT']) merged['Email Verification Endpoint'] = String(dict['EMAILVERIFICATIONENDPOINT']);
      if (dict['EMAILSENDERPROVIDER'] || dict['DISPATCHMODE']) merged['Email Sender Provider'] = String(dict['EMAILSENDERPROVIDER'] || dict['DISPATCHMODE']);
      if (dict['EMAILSENDERENDPOINT']) merged['Email Sender Endpoint'] = String(dict['EMAILSENDERENDPOINT']);
    }
  }

  // 3. Read from Persona Mapper Settings (39 cols)
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const records = getPersonaMapperSheetRecords_(settingsSheet);
    const match = records.find(function(row) {
      return String(row['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();
    });
    if (match) {
      Object.keys(match).forEach(function(k) {
        if (match[k] !== '' && match[k] !== null && match[k] !== undefined) {
          merged[k] = match[k];
        }
      });
    }
  }

  if (!merged['Department']) {
    merged['Department'] = typeof resolveCampaignClientName_ === 'function'
      ? resolveCampaignClientName_(campaignId, '', spreadsheet)
      : 'Growth Target';
  }

  if (!merged['Custom Business Needs'] && merged['Notes']) {
    merged['Custom Business Needs'] = String(merged['Notes']);
  }

  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentData = agentSheet.getDataRange().getValues();
    for (let i = 1; i < agentData.length; i++) {
      const k = String(agentData[i][1] || agentData[i][0] || '').trim().toUpperCase();
      const rawVal = agentData[i].length >= 3 ? agentData[i][2] : agentData[i][1];
      const v = String(rawVal || '').trim();
      if (!v) continue;
      if (k === 'PRIMARY_AI_PROVIDER' && !merged['AI Provider']) merged['AI Provider'] = v.toUpperCase();
      if (merged['AI Provider'] === 'OPENAI' && k === 'OPENAI_MODEL' && !merged['AI Model']) merged['AI Model'] = v;
      if (merged['AI Provider'] === 'GEMINI' && k === 'GEMINI_MODEL' && !merged['AI Model']) merged['AI Model'] = v;
    }
  }

  if (!merged['AI Provider']) merged['AI Provider'] = 'GEMINI';
  if (!merged['AI Model']) {
    merged['AI Model'] = String(merged['AI Provider']).toUpperCase() === 'OPENAI' ? 'gpt-4o' : 'gemini-3.8-flash';
  }

  return merged;
}

function executeAiCall_(options) {
  if (typeof callAI === 'function') {
    return callAI(options);
  }

  const provider = String(options.provider || 'GEMINI').toUpperCase();

  if (provider === 'OPENAI') {
    if (typeof requestOpenAIChatCompletion_ === 'function') {
      const rawOpenAi = requestOpenAIChatCompletion_(options);
      const text = rawOpenAi && rawOpenAi.choices && rawOpenAi.choices[0] && rawOpenAi.choices[0].message
        ? rawOpenAi.choices[0].message.content
        : '{}';
      return { text: text, json: JSON.parse(text), citations: [], raw: rawOpenAi };
    }
    throw new Error('OpenAIProvider / AIProvider not loaded.');
  }

  if (provider === 'GEMINI' && typeof callGemini_ === 'function') {
    return callGemini_(options);
  }

  throw new Error('AI router not available for ' + provider);
}

/**
 * Returns qualified companies for the active Campaign ID.
 * By default skips companies that already have mapped personas in "Persona Leads"
 * or "Prospect Review Queue" for the same Campaign ID (Scenario B preserved).
 */
function getPersonaMapperCompanies_(spreadsheet, campaignId, settings, options) {
  const sheet = spreadsheet.getSheetByName('Company Database');
  if (!sheet) throw new Error('Company Database sheet was not found.');

  const skipAlreadyMapped = !(options && options.skipAlreadyMapped === false);
  const mappedTracker = skipAlreadyMapped ? getMappedCompanyTracker_(spreadsheet, campaignId) : null;
  const minScore = Number(settings && settings['Minimum Company Opportunity Score']) || 0;
  const seenInDb = { companyIds: {}, companyNames: {} };

  return getPersonaMapperSheetRecords_(sheet).filter(function(row) {
    if (String(row['Campaign ID'] || '').trim() !== String(campaignId).trim()) return false;

    const qualification = String(row['Status'] || row['Qualification result'] || row['Qualification Result'] || '').trim().toLowerCase();
    if (qualification && (
      qualification.indexOf('disqualified') >= 0 ||
      qualification.indexOf('rejected') >= 0 ||
      !/\b(qualified|approved|pass|passed|verified|enriched)\b/.test(qualification)
    )) {
      return false;
    }

    const score = Number(row['Company Opportunity Score'] || row['Score'] || row['initialScore'] || 0);
    if (minScore && score && score < minScore) return false;

    const companyId = String(row['Company ID'] || '').trim().toLowerCase();
    const companyName = String(row['Company Name'] || row['Company name'] || '').trim();
    const normCompanyName = normalizePersonaCompanyName_(companyName);

    if (skipAlreadyMapped && mappedTracker) {
      if ((companyId && mappedTracker.companyIds[companyId]) ||
          (normCompanyName && mappedTracker.companyNames[normCompanyName])) {
        return false;
      }
    }

    // Prevent processing duplicate company rows inside Company Database in the same Campaign ID
    if ((companyId && seenInDb.companyIds[companyId]) ||
        (normCompanyName && seenInDb.companyNames[normCompanyName])) {
      return false;
    }
    if (companyId) seenInDb.companyIds[companyId] = true;
    if (normCompanyName) seenInDb.companyNames[normCompanyName] = true;

    return true;
  });
}

/**
 * Identifies all companies that already have personas mapped for a given Campaign ID
 * across "Persona Leads" and "Prospect Review Queue".
 */
function getMappedCompanyTracker_(spreadsheet, campaignId) {
  const tracker = { companyIds: {}, companyNames: {} };
  const targetCamp = String(campaignId || '').trim();
  if (!spreadsheet || !targetCamp) return tracker;

  [PERSONA_MAPPER_SHEETS_.personas, PERSONA_MAPPER_SHEETS_.queue].forEach(function(sheetName) {
    const sh = spreadsheet.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return;

    const records = getPersonaMapperSheetRecords_(sh);
    records.forEach(function(row) {
      if (String(row['Campaign ID'] || '').trim() !== targetCamp) return;

      const cid = String(row['Company ID'] || '').trim().toLowerCase();
      const cname = normalizePersonaCompanyName_(row['Company Name'] || row['Company name'] || '');
      if (cid) tracker.companyIds[cid] = true;
      if (cname) tracker.companyNames[cname] = true;
    });
  });

  return tracker;
}

/**
 * Builds lookup sets of already-existing personas in the given Campaign ID across
 * "Persona Leads" and "Prospect Review Queue" to enforce strict campaign-level deduplication.
 */
function getExistingCampaignPersonas_(spreadsheet, campaignId) {
  const tracker = { linkedinUrls: {}, emails: {}, nameCompanyKeys: {} };
  const targetCamp = String(campaignId || '').trim();
  if (!spreadsheet || !targetCamp) return tracker;

  [PERSONA_MAPPER_SHEETS_.personas, PERSONA_MAPPER_SHEETS_.queue].forEach(function(sheetName) {
    const sh = spreadsheet.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return;

    const records = getPersonaMapperSheetRecords_(sh);
    records.forEach(function(row) {
      if (String(row['Campaign ID'] || '').trim() !== targetCamp) return;

      const normLi = normalizeLinkedInUrl_(row['LinkedIn URL'] || '');
      if (normLi) tracker.linkedinUrls[normLi] = true;

      const email = String(row['Business Email'] || '').trim().toLowerCase();
      if (email) tracker.emails[email] = true;

      const normName = normalizePersonaName_(row['Name'] || '');
      const compId = String(row['Company ID'] || '').trim().toLowerCase();
      const normComp = normalizePersonaCompanyName_(row['Company Name'] || row['Company name'] || '');

      if (normName && compId) tracker.nameCompanyKeys[normName + '|id:' + compId] = true;
      if (normName && normComp) tracker.nameCompanyKeys[normName + '|name:' + normComp] = true;
    });
  });

  return tracker;
}

function isDuplicatePersonaInCampaign_(persona, company, tracker, extraEmail) {
  if (!persona || !tracker) return false;

  const normLi = normalizeLinkedInUrl_(persona.linkedinUrl || '');
  if (normLi && tracker.linkedinUrls[normLi]) return true;

  const email = String(extraEmail || persona.businessEmail || '').trim().toLowerCase();
  if (email && tracker.emails[email]) return true;

  const normName = normalizePersonaName_(persona.name || '');
  const compId = String((company && company['Company ID']) || '').trim().toLowerCase();
  const normComp = normalizePersonaCompanyName_((company && (company['Company Name'] || company['Company name'])) || '');

  if (normName && compId && tracker.nameCompanyKeys[normName + '|id:' + compId]) return true;
  if (normName && normComp && tracker.nameCompanyKeys[normName + '|name:' + normComp]) return true;

  return false;
}

function registerPersonaInCampaignTracker_(persona, company, tracker, emailInfo) {
  if (!persona || !tracker) return;

  const normLi = normalizeLinkedInUrl_(persona.linkedinUrl || '');
  if (normLi) tracker.linkedinUrls[normLi] = true;

  const email = String((emailInfo && emailInfo.businessEmail) || persona.businessEmail || '').trim().toLowerCase();
  if (email) tracker.emails[email] = true;

  const normName = normalizePersonaName_(persona.name || '');
  const compId = String((company && company['Company ID']) || '').trim().toLowerCase();
  const normComp = normalizePersonaCompanyName_((company && (company['Company Name'] || company['Company name'])) || '');

  if (normName && compId) tracker.nameCompanyKeys[normName + '|id:' + compId] = true;
  if (normName && normComp) tracker.nameCompanyKeys[normName + '|name:' + normComp] = true;
}

function normalizeLinkedInUrl_(url) {
  let clean = String(url || '').trim().toLowerCase();
  if (!clean || clean.indexOf('linkedin.com') < 0) return '';
  clean = clean.replace(/^https?:\/\//i, '');
  clean = clean.replace(/^[a-z0-9-]+\.linkedin\.com\//i, 'linkedin.com/');
  clean = clean.split('?')[0].split('#')[0].replace(/\/+$/, '');
  return clean;
}

function normalizePersonaName_(name) {
  return String(name || '')
    .trim()
    .toLowerCase()
    .replace(/[.,'"()]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizePersonaCompanyName_(name) {
  let clean = String(name || '').trim().toLowerCase();
  if (!clean) return '';

  clean = clean.replace(/\([^)]*\)/g, ' ');
  clean = clean.replace(/[.,'"&]+/g, ' ').replace(/[-_/]+/g, ' ').replace(/\s+/g, ' ').trim();

  const stopSuffixes = {
    'inc': true, 'incorporated': true, 'corp': true, 'corporation': true,
    'ltd': true, 'limited': true, 'llc': true, 'llp': true, 'lp': true,
    'plc': true, 'gmbh': true, 'sa': true, 'ag': true, 'co': true,
    'company': true, 'companies': true, 'of': true, 'group': true,
    'holdings': true, 'canada': true, 'usa': true, 'us': true,
    'division': true, 'subsidiary': true, 'branch': true
  };

  const words = clean.split(' ').filter(Boolean);
  while (words.length > 1 && stopSuffixes[words[words.length - 1]]) {
    words.pop();
  }

  return words.join(' ').trim();
}

function scoreCompanyOpportunity_(result, settings) {
  const dimensions = result.dimensions || {};
  const map = [
    ['esgReportingRelevance', 'Weight ESG Reporting Relevance', 20],
    ['supplyChainComplexity', 'Weight Supply Chain Complexity', 20],
    ['sustainabilityDataNeeds', 'Weight Sustainability Data Needs', 20],
    ['digitalTransformationRelevance', 'Weight Digital Transformation Relevance', 15],
    ['businessTriggers', 'Weight Business Triggers', 15],
    ['solutionFit', 'Weight eValuater Solution Fit', 10]
  ];
  let weighted = 0, totalWeight = 0;
  map.forEach(function(item) {
    const weight = configuredNumber_(settings, item[1], item[2]);
    weighted += clampScore_(dimensions[item[0]]) * weight;
    totalWeight += weight;
  });
  const score = totalWeight ? Math.round(weighted / totalWeight) : 85;
  const rationale = map.map(function(item) {
    return item[0] + ': ' + clampScore_(dimensions[item[0]]) + '/100';
  }).join('; ');

  return { score: score, label: score >= 75 ? 'High' : score >= 50 ? 'Medium' : 'Low', rationale: rationale };
}

function scorePersonaFit_(dimensions) {
  const weights = [['decisionAuthority', 25], ['functionalRelevance', 25], ['opportunityConnection', 25], ['recentActivity', 15], ['evidenceQuality', 10]];
  return Math.round(weights.reduce(function(sum, item) {
    return sum + clampScore_(dimensions[item[0]]) * item[1];
  }, 0) / 100);
}

function normalizeEsgMaturity_(rawVal) {
  const val = String(rawVal || '').trim();
  const lower = val.toLowerCase();
  const allowed = ['Early Stage', 'Developing / Spreadsheet-reliant', 'Developing', 'Mature', 'Advanced', 'N/A'];
  for (let i = 0; i < allowed.length; i++) {
    if (allowed[i].toLowerCase() === lower) return allowed[i];
  }
  if (lower.indexOf('spreadsheet') >= 0) return 'Developing / Spreadsheet-reliant';
  if (lower.indexOf('early') >= 0 || lower.indexOf('initial') >= 0) return 'Early Stage';
  if (lower.indexOf('develop') >= 0 || lower.indexOf('moderate') >= 0 || lower.indexOf('intermediate') >= 0) return 'Developing';
  if (lower.indexOf('mature') >= 0) return 'Mature';
  if (lower.indexOf('advanc') >= 0 || lower.indexOf('high') >= 0) return 'Advanced';
  return 'N/A';
}

function normalizeSolutionFit_(rawVal, scoreObj) {
  const val = String(rawVal || '').trim();
  const lower = val.toLowerCase();
  const allowed = ['High', 'Medium', 'Low', 'Disqualified - Direct Competitor'];
  for (let i = 0; i < allowed.length; i++) {
    if (allowed[i].toLowerCase() === lower) return allowed[i];
  }
  if (lower.indexOf('competitor') >= 0 || lower.indexOf('disqualified') >= 0) return 'Disqualified - Direct Competitor';
  if (lower.indexOf('high') >= 0) return 'High';
  if (lower.indexOf('medium') >= 0 || lower.indexOf('moderate') >= 0) return 'Medium';
  if (lower.indexOf('low') >= 0) return 'Low';
  return (scoreObj && scoreObj.label && allowed.indexOf(scoreObj.label) >= 0) ? scoreObj.label : 'Medium';
}

function makeCompanyIntelligenceRecord_(source, enrichment, score, citations, settings) {
  const evidence = Array.isArray(enrichment.evidence) ? enrichment.evidence : [];
  const urls = citations.concat(evidence.map(function(item) { return item && item.url ? item.url : ''; })).filter(Boolean);

  return {
    'Campaign ID': source['Campaign ID'],
    'Campaign Name': settings['Campaign Name'] || '',
    'Company ID': source['Company ID'] || '',
    'Company Name': source['Company name'] || source['Company Name'] || '',
    'Country': enrichment.country || source['Country'] || source['Headquarters'] || '',
    'Industry': source['Industry'] || '',
    'Company Size': source['Employee Size'] || source['Company size'] || '',
    'ESG Frameworks': enrichment.esgFrameworks || 'N/A',
    'Reporting Obligations': enrichment.reportingObligations || 'Standard',
    'ESG Reporting Maturity': normalizeEsgMaturity_(enrichment.reportingMaturity),
    'Supply Chain Profile': enrichment.supplyChainProfile || 'Enterprise Stack',
    'Scope 3 Exposure': enrichment.scope3Exposure || 'N/A',
    'Sustainability Data Needs': enrichment.sustainabilityDataNeeds || 'Operational monitoring',
    'Digital Transformation Signals': enrichment.digitalTransformationSignals || 'Cloud architecture',
    'ESG Governance': enrichment.esgGovernance || 'Director of Sustainability',
    'Business Triggers': enrichment.businessTriggers || settings['Custom Business Needs'] || '',
    'eValuater Solution Fit': normalizeSolutionFit_(enrichment.solutionFit, score),
    'Company Opportunity Score': score.score,
    'Company Opportunity': score.label,
    'Company Score Rationale': score.rationale,
    'Evidence URLs': Array.from(new Set(urls)).join('\n'),
    'Evidence Checked Date': new Date().toISOString().split('T')[0],
    'Verification Status': urls.length ? 'Verified' : 'Needs review',
    'Human Review Status': 'Accepted',
    'Last Enriched At': new Date().toLocaleString()
  };
}

function makePersonaRecord_(company, persona, fitScore, emailInfo, settings) {
  return {
    'Campaign ID': company['Campaign ID'],
    'Campaign Name': settings['Campaign Name'] || '',
    'Company ID': company['Company ID'] || '',
    'Company Name': company['Company name'] || company['Company Name'] || '',
    'Industry': company['Industry'] || '',
    'Persona ID': persona.linkedinUrl || [persona.name, company['Company ID'] || company['Company name'], company['Campaign ID']].join('|'),
    'Name': persona.name || '',
    'Title': persona.title || '',
    'Persona Type': persona.personaType || 'Key Decision Maker',
    'LinkedIn URL': persona.linkedinUrl || '',
    'Signal': persona.signal || '',
    'Signal Date': persona.signalDate || new Date().toISOString().split('T')[0],
    'Evidence URL': persona.evidenceUrl || '',
    'Buying Signal': titleCaseSignal_(persona.buyingSignal) || 'Active',
    'Verification Status': persona.verificationStatus || 'Verified',
    'Confidence': titleCaseSignal_(persona.confidence) || 'High',
    'Persona Fit Score': fitScore,
    'Personalized Hook': persona.personalizedHook || '',
    'Review Status': 'Ready for Outreach',
    'Last Verified At': new Date().toLocaleString(),
    'Business Email': emailInfo.businessEmail,
    'Email Domain': emailInfo.emailDomain,
    'Email Source URL': emailInfo.emailSourceUrl,
    'Email Discovery Method': emailInfo.emailDiscoveryMethod,
    'Email Verification Status': emailInfo.emailVerificationStatus,
    'Email Verification Date': emailInfo.emailVerificationDate,
    'Email Verification Provider': emailInfo.emailVerificationProvider,
    'Email Identity Confidence': emailInfo.emailIdentityConfidence,
    'Email Verification Notes': emailInfo.emailVerificationNotes
  };
}

function makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings, suppressionRules) {
  const signal = titleCaseSignal_(persona.buyingSignal) || 'Active';
  const signalScore = signal === 'Active' ? 100 : signal === 'Relevant' ? 70 : 40;
  const weights = [
    configuredNumber_(settings, 'Email Priority Weight Company Score', 50),
    configuredNumber_(settings, 'Email Priority Weight Persona Fit', 30),
    configuredNumber_(settings, 'Email Priority Weight Buying Signal', 20)
  ];
  const total = weights.reduce(function(sum, v) { return sum + v; }, 0) || 100;
  const priority = Math.round((companyScore.score * weights[0] + fitScore * weights[1] + signalScore * weights[2]) / total);
  const name = persona.name || '';
  const companyName = company['Company name'] || company['Company Name'] || '';
  const urls = citations.concat([persona.evidenceUrl || '', emailInfo.emailSourceUrl || '']).filter(Boolean);
  const defaultSubject = (settings['Campaign Name'] || 'Collaboration') + ' for ' + companyName;

  const email = String(emailInfo.businessEmail || '').trim().toLowerCase();
  const domain = String(emailInfo.emailDomain || '').trim().toLowerCase();
  const rules = suppressionRules || { domains: [], emails: [] };
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@', 'office@', 'jobs@', 'careers@'];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  let complianceStatus = 'Passed';
  let suppressionCheck = 'Clear';
  let humanApprovalStatus = 'Pending Review';
  let sendStatus = 'Queued';
  let sendError = '';

  const minFitForReview = resolveNumericPersonaThreshold_(settings['Minimum Persona Confidence'], 85);

  if (!email || !emailRegex.test(email)) {
    complianceStatus = 'Failed - Invalid Syntax';
    suppressionCheck = 'Blocked: Bad Email Syntax';
    humanApprovalStatus = 'Rejected (Suppression Guard)';
    sendStatus = 'Suppressed - Not Sent';
    sendError = suppressionCheck;
  } else if (rolePrefixes.some(function(pre) { return email.startsWith(pre); })) {
    complianceStatus = 'Failed - Role Based';
    suppressionCheck = 'Blocked: Role-Based Account';
    humanApprovalStatus = 'Rejected (Suppression Guard)';
    sendStatus = 'Suppressed - Not Sent';
    sendError = suppressionCheck;
  } else if (rules.emails && rules.emails.includes(email)) {
    complianceStatus = 'Failed - Suppressed';
    suppressionCheck = 'Blocked: Email in Lists & Settings';
    humanApprovalStatus = 'Rejected (Suppression Guard)';
    sendStatus = 'Suppressed - Not Sent';
    sendError = suppressionCheck;
  } else if (rules.domains && rules.domains.includes(domain)) {
    complianceStatus = 'Failed - Suppressed';
    suppressionCheck = 'Blocked: Domain in Lists & Settings';
    humanApprovalStatus = 'Rejected (Suppression Guard)';
    sendStatus = 'Suppressed - Not Sent';
    sendError = suppressionCheck;
  } else if (fitScore >= minFitForReview) {
    humanApprovalStatus = 'Needs Review';
  }

  return {
    'Campaign ID': company['Campaign ID'],
    'Campaign Name': settings['Campaign Name'] || '',
    'Company ID': company['Company ID'] || '',
    'Company Name': companyName,
    'Persona ID': persona.linkedinUrl || [name, company['Company ID'] || companyName, company['Campaign ID']].join('|'),
    'Name': name,
    'Title': persona.title || '',
    'Business Email': emailInfo.businessEmail,
    'Company Opportunity Score': companyScore.score,
    'Persona Fit Score': fitScore,
    'Buying Signal': signal,
    'Priority Score': priority,
    'Priority Tier': priority >= 80 ? 'Tier 1' : priority >= 60 ? 'Tier 2' : 'Tier 3',
    'Priority Score Rationale': 'Company score ' + companyScore.score + ' | Persona fit ' + fitScore + ' | Signal ' + signal,
    'Draft Subject': persona.draftSubject || defaultSubject,
    'Draft Email': persona.draftEmail || persona.personalizedHook || '',
    'Evidence URLs': Array.from(new Set(urls)).join('\n'),
    'Review Status': complianceStatus === 'Passed' ? 'Needs Review' : 'Rejected',
    'Reviewer': '',
    'Reviewer Notes': suppressionCheck !== 'Clear' ? suppressionCheck : '',
    'Reviewed At': '',
    'Human Approval Status': humanApprovalStatus,
    'Approved By': '',
    'Approved At': '',
    'Send Status': sendStatus,
    'Sent At': '',
    'Send Error': sendError,
    'Country': emailInfo.country || company['Country'] || company['Headquarters'] || '',
    'Industry': company['Industry'] || '',
    'Company Opportunity': companyScore.label,
    'Persona Type': persona.personaType || 'Key Decision Maker',
    'LinkedIn URL': persona.linkedinUrl || '',
    'Email Domain': emailInfo.emailDomain,
    'Email Source URL': emailInfo.emailSourceUrl,
    'Email Discovery Method': emailInfo.emailDiscoveryMethod,
    'Email Verification Status': emailInfo.emailVerificationStatus,
    'Email Verification Date': emailInfo.emailVerificationDate,
    'Email Verification Provider': emailInfo.emailVerificationProvider,
    'Email Identity Confidence': emailInfo.emailIdentityConfidence,
    'Email Verification Notes': emailInfo.emailVerificationNotes,
    'Queue Status': emailInfo.queueStatus,
    'Compliance Status': complianceStatus,
    'Suppression Check': suppressionCheck,
    'Date Added': new Date().toISOString().split('T')[0],
    'Personalized Hook': persona.personalizedHook || ''
  };
}

function normalizeEmailVerificationStatus_(rawStatus) {
  const s = String(rawStatus || '').trim().toLowerCase();
  const allowed = ['valid', 'risky', 'invalid', 'catch-all', 'unverified'];
  if (allowed.indexOf(s) >= 0) return s;
  if (s === 'ok' || s === 'deliverable') return 'valid';
  if (s === 'bouncable' || s === 'bounce' || s === 'disposable') return 'invalid';
  if (s === 'catchall' || s === 'accept_all') return 'catch-all';
  if (s === 'unknown' || s === 'not found' || s === '') return 'unverified';
  return 'risky';
}

function prepareEmailIntelligence_(persona, company, settings, spreadsheet, suppressionRules) {
  const candidate = String(persona.businessEmail || '').trim().toLowerCase();
  const method = String(persona.emailDiscoveryMethod || (candidate ? 'Public web source' : 'Not Found')).trim();
  const sourceUrl = String(persona.emailSourceUrl || '').trim();
  const empty = {
    businessEmail: '',
    emailDomain: '',
    emailSourceUrl: sourceUrl,
    emailDiscoveryMethod: 'Not Found',
    emailVerificationStatus: '',
    emailVerificationDate: '',
    emailVerificationProvider: '',
    emailIdentityConfidence: 'None',
    emailVerificationNotes: 'No suitable business email found.',
    queueStatus: 'Email Research Required',
    country: persona.country || ''
  };

  if (!candidate) return empty;

  const at = candidate.lastIndexOf('@');
  const domain = at > 0 ? candidate.slice(at + 1) : '';
  const normalizedMethod = /inferred/i.test(method) ? 'Pattern Match' : method;

  if (!domain || !/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(candidate)) {
    return Object.assign({}, empty, {
      businessEmail: candidate,
      emailDomain: domain,
      emailDiscoveryMethod: normalizedMethod,
      emailVerificationStatus: 'invalid',
      emailVerificationNotes: 'Email syntax invalid.',
      queueStatus: 'Email Verification Failed'
    });
  }

  if (isPersonalEmailDomain_(domain)) {
    return Object.assign({}, empty, {
      emailDomain: domain,
      emailVerificationStatus: 'invalid',
      emailVerificationNotes: 'Personal domain rejected (Gmail/Yahoo/Outlook); professional domain required.',
      queueStatus: 'Email Research Required'
    });
  }

  const companyDomain = extractPersonaDomainFromUrl_(company['Domain'] || company['Website'] || '');
  const domainMatches = Boolean(companyDomain && (domain === companyDomain || domain.endsWith('.' + companyDomain)));
  const confidence = titleCaseSignal_(persona.emailIdentityConfidence) || (domainMatches ? 'High' : 'Medium');

  // Skip paid verifier API calls if role-based or already suppressed in Lists & Settings
  const rolePrefixes = ['info@', 'support@', 'sales@', 'admin@', 'billing@', 'contact@', 'help@', 'office@', 'jobs@', 'careers@'];
  const isRoleEmail = rolePrefixes.some(function(pre) { return candidate.startsWith(pre); });
  const rules = suppressionRules || { domains: [], emails: [] };
  const isSuppressedEmail = Boolean(rules.emails && rules.emails.indexOf(candidate) >= 0);
  const isSuppressedDomain = Boolean(rules.domains && rules.domains.indexOf(domain) >= 0);

  if (isRoleEmail || isSuppressedEmail || isSuppressedDomain) {
    return {
      businessEmail: candidate,
      emailDomain: domain,
      emailSourceUrl: sourceUrl,
      emailDiscoveryMethod: normalizedMethod,
      emailVerificationStatus: 'unverified',
      emailVerificationDate: new Date().toISOString().split('T')[0],
      emailVerificationProvider: settings['Email Verification Provider'] || 'ZeroBounce',
      emailIdentityConfidence: confidence,
      emailVerificationNotes: isRoleEmail
        ? 'Skipped verification: role-based email blocked by Suppression Guard.'
        : 'Skipped verification: email/domain already blocked in Lists & Settings.',
      queueStatus: 'Email Verification Failed',
      country: persona.country || ''
    };
  }

  const verifier = callConfiguredEmailVerifier_(candidate, settings, spreadsheet);
  const rawStatus = verifier.configured ? verifier.status : 'unverified';
  const normalizedStatus = normalizeEmailVerificationStatus_(rawStatus);

  let queueStatus = 'Ready for Review';
  let verificationNotes = verifier.notes || 'Domain verified';

  if (normalizedStatus === 'invalid') {
    queueStatus = 'Email Verification Failed';
  } else if (normalizedStatus === 'catch-all' || normalizedStatus === 'risky' || normalizedStatus === 'unverified') {
    if (verifier.configured) {
      verificationNotes = 'Soft Warning: Server status is ' + normalizedStatus + '. Queued for review in Gmail Drafts.';
    }
  }

  return {
    businessEmail: candidate,
    emailDomain: domain,
    emailSourceUrl: sourceUrl,
    emailDiscoveryMethod: normalizedMethod,
    emailVerificationStatus: normalizedStatus,
    emailVerificationDate: new Date().toISOString().split('T')[0],
    emailVerificationProvider: verifier.provider || settings['Email Verification Provider'] || 'ZeroBounce',
    emailIdentityConfidence: confidence,
    emailVerificationNotes: verificationNotes,
    queueStatus: queueStatus,
    country: persona.country || ''
  };
}

function isValidVerifierKeyValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const upper = s.toUpperCase();
  if (upper.indexOf('PASTE_') >= 0) return false;
  if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
  if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
  return true;
}

/**
 * Resolves the email verifier API key strictly scoped to the campaign's Client / Owner.
 * For external clients, enforces strict CLIENT_{NAME}_{PROVIDER}_API_KEY / _VERIFIER_API_KEY
 * with ZERO local fallback when verification is enabled.
 */
function resolveVerifierApiKey_(spreadsheet, settings) {
  const provider = String((settings && settings['Email Verification Provider']) || 'ZeroBounce').trim();
  if (provider === 'None') return '';

  const pLower = provider.toLowerCase();
  const candidateKeys = [];
  if (pLower.indexOf('zerobounce') >= 0) candidateKeys.push('ZEROBOUNCE_API_KEY');
  else if (pLower.indexOf('millionverifier') >= 0) candidateKeys.push('MILLIONVERIFIER_API_KEY');
  else if (pLower.indexOf('neverbounce') >= 0) candidateKeys.push('NEVERBOUNCE_API_KEY');
  else if (pLower.indexOf('custom') >= 0) candidateKeys.push('CUSTOM_VERIFIER_API_KEY');
  candidateKeys.push('VERIFIER_API_KEY', 'EMAIL_VERIFICATION_API_KEY');

  const campaignId = (settings && settings['Campaign ID']) || '';
  const clientName = (settings && settings['Department']) || '';

  if (typeof resolveClientScopedSecret_ === 'function') {
    const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
    const resolved = resolveClientScopedSecret_(
      candidateKeys,
      { campaignId: campaignId, clientName: clientName },
      {
        serviceLabel: provider + ' Email Verifier',
        throwOnMissing: !isLocal,
        allowAgentSettingsForLocal: isLocal,
        spreadsheet: spreadsheet
      }
    );
    return resolved.key || '';
  }

  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();
  for (let i = 0; i < candidateKeys.length; i++) {
    const val = scriptProps[candidateKeys[i]];
    if (isValidVerifierKeyValue_(val)) return String(val).trim();
  }
  for (let i = 0; i < candidateKeys.length; i++) {
    const val = userProps[candidateKeys[i]];
    if (isValidVerifierKeyValue_(val)) return String(val).trim();
  }
  return '';
}

/**
 * Verifies emails via ZeroBounce, MillionVerifier, NeverBounce, or Custom HTTPS
 * with strict client key isolation and ApiTracker budget guardrails.
 */
function callConfiguredEmailVerifier_(email, settings, spreadsheet) {
  const provider = String(settings['Email Verification Provider'] || 'ZeroBounce').trim();
  if (provider === 'None') {
    return {
      configured: false,
      status: 'unverified',
      provider: provider,
      notes: 'Email verification set to None; passed syntax/domain check.'
    };
  }

  const apiKey = resolveVerifierApiKey_(spreadsheet, settings);
  if (!apiKey) {
    return {
      configured: false,
      status: 'unverified',
      provider: provider,
      notes: 'Verifier key not set; passed syntax/domain check (unverified for local free-tier testing).'
    };
  }

  const campaignId = settings['Campaign ID'] || 'CMP-DEFAULT';
  const clientName = settings['Department'] || 'Growth Target';

  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(campaignId, clientName);
  }

  const p = provider.toLowerCase();
  try {
    if (p.indexOf('zerobounce') >= 0) {
      const url = 'https://api.zerobounce.net/v2/validate?api_key=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email) + '&ip_address=';
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = response.getResponseCode();

      if (code >= 300) {
        if (typeof logApiUsage_ === 'function') {
          logApiUsage_({
            campaignId: campaignId,
            clientName: clientName,
            serviceName: 'ZeroBounce',
            operation: 'EmailVerification',
            httpStatus: code,
            inputUnits: 1,
            outputUnits: 0,
            status: 'ERROR',
            details: response.getContentText().slice(0, 200),
            modelName: 'ZeroBounce'
          });
        }
        return { configured: true, status: 'unverified', provider: 'ZeroBounce', notes: 'ZeroBounce status ' + code };
      }

      const body = JSON.parse(response.getContentText() || '{}');
      const st = String(body.status || '').toLowerCase();
      const status = normalizeEmailVerificationStatus_(st);

      if (typeof logApiUsage_ === 'function') {
        logApiUsage_({
          campaignId: campaignId,
          clientName: clientName,
          serviceName: 'ZeroBounce',
          operation: 'EmailVerification',
          httpStatus: code,
          inputUnits: 1,
          outputUnits: 0,
          status: 'SUCCESS',
          details: 'Email: ' + email + ' | Status: ' + st,
          modelName: 'ZeroBounce'
        });
      }

      return { configured: true, status: status, provider: 'ZeroBounce', notes: 'ZeroBounce: ' + st };
    }

    if (p.indexOf('millionverifier') >= 0) {
      const url = 'https://api.millionverifier.com/api/v3/?api=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = response.getResponseCode();

      if (code >= 300) {
        if (typeof logApiUsage_ === 'function') {
          logApiUsage_({
            campaignId: campaignId,
            clientName: clientName,
            serviceName: 'MillionVerifier',
            operation: 'EmailVerification',
            httpStatus: code,
            inputUnits: 1,
            outputUnits: 0,
            status: 'ERROR',
            details: response.getContentText().slice(0, 200),
            modelName: 'MillionVerifier'
          });
        }
        return { configured: true, status: 'unverified', provider: 'MillionVerifier', notes: 'MillionVerifier status ' + code };
      }

      const body = JSON.parse(response.getContentText() || '{}');
      const res = String(body.result || '').toLowerCase();
      const status = normalizeEmailVerificationStatus_(res);

      if (typeof logApiUsage_ === 'function') {
        logApiUsage_({
          campaignId: campaignId,
          clientName: clientName,
          serviceName: 'MillionVerifier',
          operation: 'EmailVerification',
          httpStatus: code,
          inputUnits: 1,
          outputUnits: 0,
          status: 'SUCCESS',
          details: 'Email: ' + email + ' | Result: ' + res,
          modelName: 'MillionVerifier'
        });
      }

      return { configured: true, status: status, provider: 'MillionVerifier', notes: 'MillionVerifier: ' + res };
    }

    if (p.indexOf('neverbounce') >= 0) {
      const url = 'https://api.neverbounce.com/v4/single/check?key=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = response.getResponseCode();

      if (code >= 300) {
        if (typeof logApiUsage_ === 'function') {
          logApiUsage_({
            campaignId: campaignId,
            clientName: clientName,
            serviceName: 'NeverBounce',
            operation: 'EmailVerification',
            httpStatus: code,
            inputUnits: 1,
            outputUnits: 0,
            status: 'ERROR',
            details: response.getContentText().slice(0, 200),
            modelName: 'NeverBounce'
          });
        }
        return { configured: true, status: 'unverified', provider: 'NeverBounce', notes: 'NeverBounce status ' + code };
      }

      const body = JSON.parse(response.getContentText() || '{}');
      const res = String(body.result || '').toLowerCase();
      const status = normalizeEmailVerificationStatus_(res);

      if (typeof logApiUsage_ === 'function') {
        logApiUsage_({
          campaignId: campaignId,
          clientName: clientName,
          serviceName: 'NeverBounce',
          operation: 'EmailVerification',
          httpStatus: code,
          inputUnits: 1,
          outputUnits: 0,
          status: 'SUCCESS',
          details: 'Email: ' + email + ' | Result: ' + res,
          modelName: 'NeverBounce'
        });
      }

      return { configured: true, status: status, provider: 'NeverBounce', notes: 'NeverBounce: ' + res };
    }

    if (p.indexOf('custom') >= 0) {
      const customEndpoint = String(settings['Email Verification Endpoint'] || '').trim();
      if (!customEndpoint || !/^https:\/\//i.test(customEndpoint)) {
        return { configured: false, status: 'unverified', provider: 'Custom HTTPS', notes: 'Custom HTTPS endpoint missing; queued unverified.' };
      }
      const sep = customEndpoint.indexOf('?') >= 0 ? '&' : '?';
      const url = customEndpoint + sep + 'email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + apiKey },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code >= 300) {
        if (typeof logApiUsage_ === 'function') {
          logApiUsage_({
            campaignId: campaignId,
            clientName: clientName,
            serviceName: 'Custom HTTPS',
            operation: 'EmailVerification',
            httpStatus: code,
            inputUnits: 1,
            outputUnits: 0,
            status: 'ERROR',
            details: response.getContentText().slice(0, 200),
            modelName: 'Custom HTTPS'
          });
        }
        return { configured: true, status: 'unverified', provider: 'Custom HTTPS', notes: 'Custom verifier HTTP ' + code };
      }
      const body = JSON.parse(response.getContentText() || '{}');
      const rawRes = String(body.status || body.result || 'unverified').toLowerCase();
      const status = normalizeEmailVerificationStatus_(rawRes);
      if (typeof logApiUsage_ === 'function') {
        logApiUsage_({
          campaignId: campaignId,
          clientName: clientName,
          serviceName: 'Custom HTTPS',
          operation: 'EmailVerification',
          httpStatus: code,
          inputUnits: 1,
          outputUnits: 0,
          status: 'SUCCESS',
          details: 'Email: ' + email + ' | Status: ' + rawRes,
          modelName: 'Custom HTTPS'
        });
      }
      return { configured: true, status: status, provider: 'Custom HTTPS', notes: 'Custom HTTPS: ' + rawRes };
    }

    return { configured: false, status: 'unverified', provider: provider, notes: 'Verifier not mapped; queued unverified.' };
  } catch (err) {
    if (String(err && err.message ? err.message : err).indexOf('RUNAWAY LOOP ABORTED') >= 0) {
      throw err;
    }
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: campaignId,
        clientName: clientName,
        serviceName: provider,
        operation: 'EmailVerification',
        httpStatus: 500,
        inputUnits: 1,
        outputUnits: 0,
        status: 'ERROR',
        details: err.toString(),
        modelName: provider
      });
    }
    return { configured: true, status: 'unverified', provider: provider, notes: 'Verification request error: ' + err.message };
  }
}

function loadSuppressionRulesFromSettings_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Lists & Settings');
  if (!sheet || sheet.getLastRow() <= 1) return { domains: [], emails: [] };
  const data = sheet.getDataRange().getValues();
  const rules = { domains: [], emails: [] };

  for (let i = 1; i < data.length; i++) {
    const listName = String(data[i][0] || '').trim().toLowerCase();
    const value = String(data[i][1] || '').trim().toLowerCase();
    const status = String(data[i][2] || '').trim().toLowerCase();

    const isActive = (status === 'active' || status === 'true' || status === '');
    const isSuppression = listName.includes('suppression') ||
                          listName.includes('blacklist') ||
                          listName.includes('opt-out') ||
                          listName.includes('competitor');

    if (isActive && isSuppression && value) {
      if (value.includes('@')) {
        rules.emails.push(value);
      } else {
        rules.domains.push(value.replace(/^@/, ''));
      }
    }
  }
  return rules;
}

function recordPersonaGroundingEvidence_(spreadsheet, campaignId, company, persona, grounding, defaultQuery) {
  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');
  if (!evidenceSheet) return;

  const evidenceUrl = String(persona.evidenceUrl || persona.linkedinUrl || persona.emailSourceUrl || '').trim();
  if (!evidenceUrl) return;

  const personaName = String(persona.name || '').trim();
  const companyName = String(company['Company name'] || company['Company Name'] || '').trim();
  const entityName = personaName ? (personaName + ' (' + companyName + ')') : companyName;

  const normUrl = typeof normalizeSourceUrl_ === 'function'
    ? normalizeSourceUrl_(evidenceUrl)
    : evidenceUrl.replace(/\/+$/, '');
  const isGrounded = Boolean(grounding.citationsByUrl && grounding.citationsByUrl[normUrl]);
  const conf = titleCaseSignal_(persona.confidence);

  // Only award 'High' Verification Score when live search grounding verified the URL
  const verificationScore = isGrounded
    ? 'High'
    : ((conf === 'High' || conf === 'Medium') ? 'Medium' : 'Unknown');

  const snippet = String(persona.signal || persona.personalizedHook || persona.buyingSignal || (persona.title + ' at ' + companyName)).trim();
  const searchQueries = grounding.searchQueries || [];
  const queryUsed = searchQueries.length > 0
    ? searchQueries.join('; ')
    : (defaultQuery || ('Persona search: ' + (persona.title || 'Leadership') + ' at ' + companyName));

  const evidenceRow = [
    'EVD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
    campaignId,
    entityName,
    'Persona',
    queryUsed,
    evidenceUrl,
    snippet,
    verificationScore,
    new Date().toLocaleString()
  ];

  if (typeof savePreparedPersonaEvidenceRows_ === 'function') {
    try {
      savePreparedPersonaEvidenceRows_(campaignId, [evidenceRow]);
      return;
    } catch (e) {
      Logger.log('savePreparedPersonaEvidenceRows_ notice: ' + e.message);
    }
  }

  evidenceSheet.appendRow(evidenceRow);
}

function recordCompanyGroundingEvidence_(spreadsheet, campaignId, company, evItem, grounding, defaultQuery) {
  const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');
  if (!evidenceSheet || !evItem || !evItem.url) return;

  const evidenceUrl = String(evItem.url).trim();
  const companyName = String(company['Company name'] || company['Company Name'] || '').trim();
  const normUrl = typeof normalizeSourceUrl_ === 'function' ? normalizeSourceUrl_(evidenceUrl) : evidenceUrl.replace(/\/+$/, '');
  const isGrounded = Boolean(grounding.citationsByUrl && grounding.citationsByUrl[normUrl]);
  const verificationScore = isGrounded ? 'High' : 'Medium';
  const snippet = String(evItem.claim || evItem.extracted_information || 'Account research trigger').trim();
  const searchQueries = grounding.searchQueries || [];
  const queryUsed = searchQueries.length > 0 ? searchQueries.join('; ') : (defaultQuery || 'Deep company persona research');

  const evidenceRow = [
    'EVD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
    campaignId,
    companyName,
    'Company',
    queryUsed,
    evidenceUrl,
    snippet,
    verificationScore,
    new Date().toLocaleString()
  ];

  if (typeof savePreparedPersonaEvidenceRows_ === 'function') {
    try {
      savePreparedPersonaEvidenceRows_(campaignId, [evidenceRow]);
      return;
    } catch (e) {
      Logger.log('savePreparedPersonaEvidenceRows_ notice: ' + e.message);
    }
  }

  evidenceSheet.appendRow(evidenceRow);
}

/**
 * Extracts a clean bare domain from either a full URL or bare domain string.
 */
function extractPersonaDomainFromUrl_(url) {
  if (!url) return '';
  let clean = String(url).trim().toLowerCase();
  clean = clean.replace(/^https?:\/\//i, '').replace(/^www\./i, '');
  return clean.split('/')[0].split('?')[0].split('#')[0];
}

function isPersonalEmailDomain_(domain) {
  return ['gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'live.com', 'yahoo.com', 'icloud.com', 'me.com', 'aol.com', 'proton.me', 'protonmail.com'].indexOf(String(domain).toLowerCase()) >= 0;
}

function getPersonaMapperCampaignId_(spreadsheet, ui) {
  const sheet = spreadsheet.getActiveSheet();
  if (['Campaigns', 'Company Database', 'Persona Mapper Settings', 'Campaign Configuration', 'Company Intelligence'].indexOf(sheet.getName()) >= 0) {
    const range = sheet.getActiveRange();
    if (range && range.getRow() > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      const values = sheet.getRange(range.getRow(), 1, 1, headers.length).getValues()[0];
      const record = {};
      headers.forEach(function(h, idx) { record[h] = values[idx]; });
      if (record['Campaign ID']) return String(record['Campaign ID']).trim();
    }
  }
  if (typeof getStudioActiveCampaignId_ === 'function') {
    const activeId = getStudioActiveCampaignId_();
    if (activeId) return activeId;
  }
  if (ui && typeof ui.prompt === 'function') {
    const response = ui.prompt('Persona Mapper', 'Enter the Campaign ID to enrich (e.g. CMP-001):', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      return String(response.getResponseText() || '').trim();
    }
  }
  return '';
}

function getPersonaMapperSheetRecords_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = values[0].map(function(value) { return String(value || '').trim(); });
  return values.slice(1).filter(function(row) {
    return row.some(function(v) { return v !== '' && v !== null; });
  }).map(function(row) {
    const record = {};
    headers.forEach(function(h, idx) { if (h) record[h] = row[idx]; });
    return record;
  });
}

function ensurePersonaMapperSheets_(spreadsheet) {
  let legacy = spreadsheet.getSheetByName('Email Review Queue');
  let queue = spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue);
  if (!queue && legacy) { legacy.setName(PERSONA_MAPPER_SHEETS_.queue); queue = legacy; }

  [
    [PERSONA_MAPPER_SHEETS_.companies, PERSONA_MAPPER_HEADERS_.companies],
    [PERSONA_MAPPER_SHEETS_.personas, PERSONA_MAPPER_HEADERS_.personas],
    [PERSONA_MAPPER_SHEETS_.settings, PERSONA_MAPPER_HEADERS_.settings],
    [PERSONA_MAPPER_SHEETS_.queue, PERSONA_MAPPER_HEADERS_.queue]
  ].forEach(function(d) {
    let sh = spreadsheet.getSheetByName(d[0]);
    if (!sh) sh = spreadsheet.insertSheet(d[0]);
    if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
      if (sh.getMaxColumns() < d[1].length) sh.insertColumnsAfter(sh.getMaxColumns(), d[1].length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, d[1].length).setValues([d[1]]);
      sh.getRange(1, 1, 1, d[1].length)
        .setFontWeight('bold')
        .setFontColor('#174ea6')
        .setBackground('#e8f0fe')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sh.setFrozenRows(1);
      return;
    }
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
    const missing = d[1].filter(function(h) { return headers.indexOf(h) < 0; });
    if (missing.length) {
      const col = sh.getLastColumn();
      if (sh.getMaxColumns() < col + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), col + missing.length - sh.getMaxColumns());
      sh.getRange(1, col + 1, 1, missing.length).setValues([missing]);
    }
  });

  const evSheet = spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.evidence);
  if (evSheet && evSheet.getLastColumn() > 0) {
    evSheet.getRange(1, 1, 1, evSheet.getLastColumn())
      .setFontWeight('bold')
      .setFontColor('#174ea6')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    evSheet.setFrozenRows(1);
  }
}

function upsertPreservingNonEmpty_(sheet, headers, keyType, record) {
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];

  function cellVal(row, header) {
    const index = headers.indexOf(header);
    return index >= 0 ? row[index] : '';
  }

  function same(row, header, value) {
    return String(cellVal(row, header) || '').trim().toLowerCase() === String(value || '').trim().toLowerCase();
  }

  function sameNormalizedCompany(row) {
    const recCompId = String(record['Company ID'] || '').trim().toLowerCase();
    const rowCompId = String(cellVal(row, 'Company ID') || '').trim().toLowerCase();
    if (recCompId && rowCompId && recCompId === rowCompId) return true;

    const recCompName = normalizePersonaCompanyName_(record['Company Name'] || '');
    const rowCompName = normalizePersonaCompanyName_(cellVal(row, 'Company Name') || '');
    return Boolean(recCompName && rowCompName && recCompName === rowCompName);
  }

  function matches(row) {
    if (!same(row, 'Campaign ID', record['Campaign ID'])) return false;

    if (keyType === 'company') {
      return sameNormalizedCompany(row);
    }

    const recNormLi = normalizeLinkedInUrl_(record['LinkedIn URL'] || '');
    const rowNormLi = normalizeLinkedInUrl_(cellVal(row, 'LinkedIn URL') || '');
    if (recNormLi && rowNormLi && recNormLi === rowNormLi) return true;

    const recEmail = String(record['Business Email'] || '').trim().toLowerCase();
    const rowEmail = String(cellVal(row, 'Business Email') || '').trim().toLowerCase();
    if (recEmail && rowEmail && recEmail === rowEmail) return true;

    if (keyType === 'queue') {
      const recPid = String(record['Persona ID'] || '').trim();
      const rowPid = String(cellVal(row, 'Persona ID') || '').trim();
      if (recPid && rowPid) {
        const normRecPid = normalizeLinkedInUrl_(recPid) || recPid.toLowerCase();
        const normRowPid = normalizeLinkedInUrl_(rowPid) || rowPid.toLowerCase();
        if (normRecPid === normRowPid) return true;
      }
    }

    const recNormName = normalizePersonaName_(record['Name'] || '');
    const rowNormName = normalizePersonaName_(cellVal(row, 'Name') || '');
    if (recNormName && rowNormName && recNormName === rowNormName && sameNormalizedCompany(row)) {
      return true;
    }

    return false;
  }

  let targetRow = 0;
  for (let i = 0; i < rows.length; i++) {
    if (matches(rows[i])) { targetRow = i + 2; break; }
  }

  if (!targetRow) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([headers.map(function(h) { return record[h] === undefined ? '' : record[h]; })]);
    return;
  }

  const existing = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
  const updated = existing.slice();
  headers.forEach(function(h, index) {
    const incoming = record[h];
    if ((existing[index] === '' || existing[index] === null) && incoming !== undefined && incoming !== '' && incoming !== null) {
      updated[index] = incoming;
    }
  });
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([updated]);
}

function configuredNumber_(settings, key, fallback) {
  const val = settings[key];
  if (val === '' || val === null || val === undefined) return fallback;
  if (typeof val === 'number') {
    if (!isFinite(val)) return fallback;
    return (val > 0 && val <= 1) ? Math.round(val * 100) : val;
  }
  const str = String(val).trim();
  if (!str) return fallback;
  const hadPercent = str.endsWith('%');
  const num = Number(str.replace(/%$/, '').trim());
  if (!isFinite(num)) return fallback;
  if (!hadPercent && num > 0 && num <= 1) return Math.round(num * 100);
  return num;
}

function resolveMinPersonaConfidenceRank_(val) {
  const str = String(val || '').trim();
  const num = Number(str);
  if (isFinite(num) && str !== '') {
    return num >= 90 ? 2 : 1;
  }
  return confidenceRank_(str || 'Medium');
}

function resolveNumericPersonaThreshold_(val, fallback) {
  const num = Number(val);
  return (isFinite(num) && val !== '' && val !== null) ? num : fallback;
}

function toBoolean_(value) {
  if (value === true) return true;
  return ['true', 'yes', '1', 'enabled'].indexOf(String(value || '').trim().toLowerCase()) >= 0;
}

function clampScore_(value) {
  const num = Number(value);
  if (!isFinite(num)) return 0;
  return Math.max(0, Math.min(100, Math.round(num)));
}

function confidenceRank_(value) {
  const norm = titleCaseSignal_(value);
  return norm === 'High' ? 2 : norm === 'Medium' ? 1 : 0;
}

function titleCaseSignal_(value) {
  const norm = String(value || '').trim().toLowerCase();
  return norm ? norm.charAt(0).toUpperCase() + norm.slice(1) : '';
}