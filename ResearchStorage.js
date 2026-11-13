/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * RESEARCH STORAGE & GROUNDING PERSISTENCE ENGINE
 * ============================================================================
 * File: ResearchStorage.js
 * 
 * Manages atomic persistence and deduplication of grounded research rows into:
 * - Tab 8: Company Database (10 columns, Column J = Campaign ID)
 * - Tab 9: Research Evidence (9 columns, Column B = Campaign ID)
 * ============================================================================
 */

/**
 * Builds a validated save plan and performs Campaign-scoped deduplication.
 * 
 * @param {Array<Array<string>>} existingCompanyRows Existing Tab 8 rows
 * @param {Object} prepared Prepared object from ResearchResults.js
 * @return {Object} Save plan with action ('save' | 'skip_duplicate')
 */
function buildResearchSavePlan_(existingCompanyRows, prepared) {
  if (!prepared || !Array.isArray(prepared.companyRow) || prepared.companyRow.length !== 10) {
    throw new Error('Prepared company data does not match the 10-column Company Database schema.');
  }
  if (!Array.isArray(prepared.evidenceRows)) {
    throw new Error('Prepared evidence data is missing or invalid.');
  }

  const companyRow = prepared.companyRow.slice();
  const companyId = companyRow[0];
  const companyName = companyRow[1];
  const campaignId = companyRow[9];

  if (!campaignId || !companyName) {
    throw new Error('Campaign ID and Company Name are required before saving.');
  }

  // Validate each evidence row matches the 9-column Tab 9 schema
  prepared.evidenceRows.forEach(function(row) {
    if (!Array.isArray(row) || row.length !== 9 || row[1] !== campaignId) {
      throw new Error('Prepared evidence does not match the 9-column Research Evidence schema.');
    }
  });

  // Check Scenario B campaign-scoped deduplication
  const duplicate = findResearchDuplicate_(existingCompanyRows, companyRow);
  if (duplicate) {
    return {
      action: 'skip_duplicate',
      existingCompanyId: duplicate.companyId,
      matchReason: duplicate.reason
    };
  }

  // Flag if company was discovered without grounded evidence
  if (prepared.evidenceRows.length === 0) {
    companyRow[8] = 'Needs Research Evidence'; // Column I: Status
  }

  return {
    action: 'save',
    companyRow: companyRow,
    evidenceRows: prepared.evidenceRows
  };
}

/**
 * Scenario B Campaign-scoped duplicate detector.
 * Allows targeting the same company across different campaigns, but prevents duplicates in the same campaign.
 * 
 * @param {Array<Array<string>>} existingCompanyRows Existing 10-column rows
 * @param {Array<string>} newCompanyRow New 10-column row
 * @return {Object|null} Duplicate info or null
 */
function findResearchDuplicate_(existingCompanyRows, newCompanyRow) {
  const targetCampaignId = String(newCompanyRow[9] || '').trim(); // Col J
  const targetDomain = normalizeCompanyWebsite_(newCompanyRow[2]); // Col C
  const targetName = normalizeCompanyName_(newCompanyRow[1]);       // Col B

  for (let index = 0; index < existingCompanyRows.length; index++) {
    const row = existingCompanyRows[index];
    const rowCampaignId = String(row[9] || '').trim();

    // Isolated per Campaign ID
    if (rowCampaignId !== targetCampaignId) continue;

    const existingDomain = normalizeCompanyWebsite_(row[2]);
    if (targetDomain && existingDomain && targetDomain === existingDomain) {
      return { companyId: row[0], reason: 'website domain' };
    }

    const existingName = normalizeCompanyName_(row[1]);
    if (targetName && existingName && targetName === existingName) {
      return { companyId: row[0], reason: 'company name' };
    }
  }

  return null;
}

/**
 * Normalizes website/domain for accurate duplicate comparisons.
 * 
 * @param {string} website Raw URL or domain
 * @return {string} Clean normalized domain
 */
function normalizeCompanyWebsite_(website) {
  return String(website || '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .split(/[?#]/)[0]
    .replace(/\/+$/, '');
}

/**
 * Normalizes company names for phonetic and punctuation consistency.
 * 
 * @param {string} name Company Name
 * @return {string} Clean normalized string
 */
function normalizeCompanyName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * Atomically persists prepared company and grounding evidence rows to Sheets under ScriptLock.
 * 
 * @param {Object} prepared Object from prepareCompanyResearchRows_
 * @return {Object} Result summary status
 */
function savePreparedResearchRows_(prepared) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  let companySheet = null;
  let companyRowNumber = 0;
  let companyWritten = false;

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('The master spreadsheet is not available.');

    const companyDefinition = APP_CONFIG.sheetDefinitions.find(function(definition) {
      return definition.name === 'Company Database';
    });
    const evidenceDefinition = APP_CONFIG.sheetDefinitions.find(function(definition) {
      return definition.name === 'Research Evidence';
    });

    companySheet = spreadsheet.getSheetByName('Company Database');
    const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');

    if (!companySheet || !evidenceSheet) {
      throw new Error('Required research tabs (Company Database or Research Evidence) are missing.');
    }

    assertResearchHeaders_(companySheet, companyDefinition);
    assertResearchHeaders_(evidenceSheet, evidenceDefinition);

    const lastCompanyRow = companySheet.getLastRow();
    const existingRows = lastCompanyRow > 1
      ? companySheet.getRange(2, 1, lastCompanyRow - 1, 10).getDisplayValues()
      : [];

    const plan = buildResearchSavePlan_(existingRows, prepared);
    if (plan.action === 'skip_duplicate') {
      return {
        status: 'duplicate_skipped',
        existingCompanyId: plan.existingCompanyId,
        matchReason: plan.matchReason
      };
    }

    // Persist company row (10 columns)
    companyRowNumber = companySheet.getLastRow() + 1;
    companySheet.getRange(companyRowNumber, 1, 1, 10).setValues([plan.companyRow]);
    SpreadsheetApp.flush();
    companyWritten = true;

    // Persist evidence rows (9 columns)
    if (plan.evidenceRows.length > 0) {
      const firstEvidenceRow = evidenceSheet.getLastRow() + 1;
      evidenceSheet
        .getRange(firstEvidenceRow, 1, plan.evidenceRows.length, 9)
        .setValues(plan.evidenceRows);
      SpreadsheetApp.flush();
    }

    return {
      status: plan.evidenceRows.length > 0 ? 'saved_for_review' : 'saved_needs_evidence',
      companyId: plan.companyRow[0],
      evidenceCount: plan.evidenceRows.length
    };
  } catch (error) {
    if (companyWritten && companySheet && companyRowNumber) {
      try {
        // Mark status (Col I / col 9) as incomplete
        companySheet.getRange(companyRowNumber, 9).setValue('Evidence Incomplete - Manual Review Required');
        SpreadsheetApp.flush();
      } catch (statusError) {
        // Suppress secondary status failure
      }
    }
    throw new Error('Research rows were not fully saved: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Atomically persists persona grounding evidence rows to Tab 9: Research Evidence under ScriptLock.
 * 
 * @param {string} campaignId Active Campaign ID
 * @param {Array<Array<string>>} evidenceRows Array of 9-column evidence rows
 * @return {number} Number of rows appended
 */
function savePreparedPersonaEvidenceRows_(campaignId, evidenceRows) {
  if (!evidenceRows || evidenceRows.length === 0) return 0;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('The master spreadsheet is not available.');

    const evidenceDefinition = APP_CONFIG.sheetDefinitions.find(function(definition) {
      return definition.name === 'Research Evidence';
    });
    const evidenceSheet = spreadsheet.getSheetByName('Research Evidence');
    if (!evidenceSheet) throw new Error('Research Evidence tab is missing.');

    assertResearchHeaders_(evidenceSheet, evidenceDefinition);

    // Validate 9 columns
    evidenceRows.forEach(function(row) {
      if (!Array.isArray(row) || row.length !== 9) {
        throw new Error('Persona evidence row does not match the 9-column schema.');
      }
    });

    const firstEvidenceRow = evidenceSheet.getLastRow() + 1;
    evidenceSheet.getRange(firstEvidenceRow, 1, evidenceRows.length, 9).setValues(evidenceRows);
    SpreadsheetApp.flush();

    return evidenceRows.length;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validates that actual sheet headers match configuration, ignoring subtle whitespace differences.
 * 
 * @param {Sheet} sheet Google Sheet instance
 * @param {Object} definition Definition object from APP_CONFIG
 */
function assertResearchHeaders_(sheet, definition) {
  if (!definition) throw new Error('Sheet definition is missing in configuration.');
  const actual = sheet.getRange(1, 1, 1, definition.headers.length).getDisplayValues()[0];
  for (let index = 0; index < definition.headers.length; index++) {
    const expectedHeader = String(definition.headers[index] || '').trim();
    const actualHeader = String(actual[index] || '').trim();
    if (actualHeader !== expectedHeader) {
      throw new Error(
        'Header mismatch in tab "' + sheet.getName() + '" at column ' + (index + 1) + 
        '. Expected: "' + expectedHeader + '", Found: "' + actualHeader + '"'
      );
    }
  }
}