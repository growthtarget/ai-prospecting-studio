/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * RESEARCH STORAGE & GROUNDING PERSISTENCE ENGINE
 * ============================================================================
 * File: ResearchStorage.js
 *
 * Manages atomic persistence and deduplication of grounded research rows into:
 * - Company Database (10 columns, Column J = Campaign ID)
 * - Research Evidence (9 columns, Column B = Campaign ID)
 * ============================================================================
 */

/**
 * Builds a validated save plan and performs Campaign-scoped deduplication.
 *
 * @param {Array<Array<string>>} existingCompanyRows Existing Company Database rows
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
  const companyName = companyRow[1];
  const campaignId = companyRow[9];

  if (!campaignId || !companyName) {
    throw new Error('Campaign ID and Company Name are required before saving.');
  }

  // Validate each evidence row matches the 9-column Research Evidence schema
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
 * Allows targeting the same company across different campaigns, but prevents duplicates
 * (including division/subsidiary/parenthetical and domain stem variants) in the same campaign.
 *
 * @param {Array<Array<string>>} existingCompanyRows Existing 10-column rows
 * @param {Array<string>} newCompanyRow New 10-column row
 * @return {Object|null} Duplicate info or null
 */
function findResearchDuplicate_(existingCompanyRows, newCompanyRow) {
  const targetCampaignId = String(newCompanyRow[9] || '').trim(); // Col J
  const targetHost = extractResearchDomainHost_(newCompanyRow[2]); // Col C
  const targetStem = extractResearchDomainStem_(targetHost);
  const targetVariants = getResearchCompanyNameVariants_(newCompanyRow[1]); // Col B
  const targetPrimaryNorm = targetVariants.length ? targetVariants[0] : '';
  const targetFirstWord = targetPrimaryNorm ? targetPrimaryNorm.split(' ')[0] : '';

  for (let index = 0; index < existingCompanyRows.length; index++) {
    const row = existingCompanyRows[index];
    const rowCampaignId = String(row[9] || '').trim();

    // Isolated per Campaign ID
    if (rowCampaignId !== targetCampaignId) continue;

    const existingHost = extractResearchDomainHost_(row[2]);
    if (targetHost && existingHost && targetHost === existingHost) {
      return { companyId: row[0], reason: 'website domain' };
    }

    const existingVariants = getResearchCompanyNameVariants_(row[1]);
    for (let v = 0; v < targetVariants.length; v++) {
      if (existingVariants.indexOf(targetVariants[v]) >= 0) {
        return { companyId: row[0], reason: 'company name' };
      }
    }

    const existingPrimaryNorm = existingVariants.length ? existingVariants[0] : '';
    const existingStem = extractResearchDomainStem_(existingHost);

    // Same root domain stem across different TLDs when primary brand word aligns
    if (targetStem && existingStem && targetStem === existingStem && targetStem.length >= 6) {
      const existingFirstWord = existingPrimaryNorm ? existingPrimaryNorm.split(' ')[0] : '';
      if (targetFirstWord && existingFirstWord && targetFirstWord.length >= 4 && targetFirstWord === existingFirstWord) {
        return { companyId: row[0], reason: 'root domain brand' };
      }
    }

    // Parent/subsidiary brand prefix match (e.g. "Enbridge Inc." + enbridge.com vs "Enbridge Gas Inc." + enbridgegas.com)
    if (targetPrimaryNorm && existingPrimaryNorm && targetStem && existingStem) {
      const shorterName = targetPrimaryNorm.length <= existingPrimaryNorm.length ? targetPrimaryNorm : existingPrimaryNorm;
      const longerName = targetPrimaryNorm.length > existingPrimaryNorm.length ? targetPrimaryNorm : existingPrimaryNorm;
      const shorterStem = targetStem.length <= existingStem.length ? targetStem : existingStem;
      const longerStem = targetStem.length > existingStem.length ? targetStem : existingStem;

      if (shorterName.length >= 5 &&
          !isGenericResearchBrandPrefix_(shorterName) &&
          (longerName === shorterName || longerName.indexOf(shorterName + ' ') === 0) &&
          shorterStem.length >= 5 &&
          longerStem.indexOf(shorterStem) === 0) {
        return { companyId: row[0], reason: 'parent/subsidiary brand' };
      }
    }
  }

  return null;
}

/**
 * Extracts a clean host domain (without scheme, www, path, or query string).
 *
 * @param {string} website Raw URL or domain
 * @return {string} Clean lowercase host domain
 */
function extractResearchDomainHost_(website) {
  return String(website || '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .split(/[/?#]/)[0]
    .trim();
}

function extractResearchDomainStem_(domain) {
  const clean = extractResearchDomainHost_(domain);
  if (!clean) return '';
  const parts = clean.split('.').filter(Boolean);
  if (parts.length === 1) return parts[0];

  const secondLevelPrefixes = { 'co': true, 'com': true, 'org': true, 'net': true, 'gov': true, 'edu': true };
  if (parts.length >= 3 &&
      secondLevelPrefixes[parts[parts.length - 2]] &&
      parts[parts.length - 1].length === 2) {
    return parts[parts.length - 3];
  }
  return parts[parts.length - 2];
}

function getResearchCompanyNameVariants_(name) {
  const raw = String(name || '').trim();
  if (!raw) return [];

  const variants = [];
  const fullNorm = normalizeResearchCompanyEntity_(raw);
  if (fullNorm) variants.push(fullNorm);

  if (raw.indexOf('/') >= 0) {
    raw.split('/').forEach(function(part) {
      const normPart = normalizeResearchCompanyEntity_(part);
      if (normPart && normPart.length >= 4 && variants.indexOf(normPart) < 0) {
        variants.push(normPart);
      }
    });
  }

  return variants;
}

function normalizeResearchCompanyEntity_(name) {
  let clean = String(name || '').trim().toLowerCase();
  if (!clean) return '';

  // Remove parenthetical qualifiers such as "(Affiliated Clean Infrastructure Division)"
  clean = clean
    .replace(/\([^)]*\)/g, ' ')
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');

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

function isGenericResearchBrandPrefix_(word) {
  const generic = {
    'american': true, 'canadian': true, 'north': true, 'south': true, 'east': true, 'west': true,
    'global': true, 'national': true, 'united': true, 'clean': true, 'green': true, 'energy': true,
    'power': true, 'solar': true, 'carbon': true, 'pacific': true, 'atlantic': true, 'western': true,
    'eastern': true, 'northern': true, 'southern': true, 'first': true, 'premier': true
  };
  return Boolean(generic[String(word || '').trim().toLowerCase()]);
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
 * Atomically persists persona grounding evidence rows to Research Evidence under ScriptLock.
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