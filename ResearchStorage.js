function buildResearchSavePlan_(existingCompanyRows, prepared) {
  if (!prepared || !Array.isArray(prepared.companyRow) || prepared.companyRow.length !== 15) {
    throw new Error('Prepared company data does not match the Company Database columns.');
  }
  if (!Array.isArray(prepared.evidenceRows)) {
    throw new Error('Prepared evidence data is missing.');
  }

  const companyRow = prepared.companyRow.slice();
  if (!companyRow[1] || !companyRow[2]) {
    throw new Error('Campaign ID and company name are required before saving.');
  }
  prepared.evidenceRows.forEach(function(row) {
    if (!Array.isArray(row) || row.length !== 10 || row[1] !== companyRow[0]) {
      throw new Error('Prepared evidence does not match the Research Evidence columns.');
    }
  });

  const duplicate = findResearchDuplicate_(existingCompanyRows, companyRow);
  if (duplicate) {
    return {
      action: 'skip_duplicate',
      existingCompanyId: duplicate.companyId,
      matchReason: duplicate.reason
    };
  }

  if (prepared.evidenceRows.length === 0) {
    companyRow[11] = 'Unknown';
    companyRow[12] = 'Needs research evidence';
  }

  return {
    action: 'save',
    companyRow: companyRow,
    evidenceRows: prepared.evidenceRows
  };
}

function findResearchDuplicate_(existingCompanyRows, newCompanyRow) {
  const campaignId = String(newCompanyRow[1] || '').trim();
  const website = normalizeCompanyWebsite_(newCompanyRow[3]);
  const companyName = normalizeCompanyName_(newCompanyRow[2]);

  for (let index = 0; index < existingCompanyRows.length; index++) {
    const row = existingCompanyRows[index];
    if (String(row[1] || '').trim() !== campaignId) continue;

    const existingWebsite = normalizeCompanyWebsite_(row[3]);
    if (website && existingWebsite && website === existingWebsite) {
      return { companyId: row[0], reason: 'website' };
    }

    const existingName = normalizeCompanyName_(row[2]);
    if (companyName && existingName && companyName === existingName) {
      return { companyId: row[0], reason: 'company name' };
    }
  }
  return null;
}

function normalizeCompanyWebsite_(website) {
  return String(website || '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .split(/[?#]/)[0]
    .replace(/\/+$/, '');
}

function normalizeCompanyName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

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
      throw new Error('Required research tabs are missing.');
    }
    assertResearchHeaders_(companySheet, companyDefinition);
    assertResearchHeaders_(evidenceSheet, evidenceDefinition);

    const lastCompanyRow = companySheet.getLastRow();
    const existingRows = lastCompanyRow > 1
      ? companySheet.getRange(2, 1, lastCompanyRow - 1, 4).getDisplayValues()
      : [];
    const plan = buildResearchSavePlan_(existingRows, prepared);
    if (plan.action === 'skip_duplicate') {
      return {
        status: 'duplicate_skipped',
        existingCompanyId: plan.existingCompanyId,
        matchReason: plan.matchReason
      };
    }

    companyRowNumber = companySheet.getLastRow() + 1;
    companySheet.getRange(companyRowNumber, 1, 1, plan.companyRow.length).setValues([plan.companyRow]);
    SpreadsheetApp.flush();
    companyWritten = true;

    if (plan.evidenceRows.length > 0) {
      const firstEvidenceRow = evidenceSheet.getLastRow() + 1;
      evidenceSheet
        .getRange(firstEvidenceRow, 1, plan.evidenceRows.length, 10)
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
        companySheet.getRange(companyRowNumber, 13).setValue('Evidence save incomplete; manual review required');
        SpreadsheetApp.flush();
      } catch (statusError) {
        // Preserve the original save failure without logging row or source data.
      }
    }
    throw new Error('Research rows were not fully saved. Check both research tabs before retrying.');
  } finally {
    lock.releaseLock();
  }
}

function assertResearchHeaders_(sheet, definition) {
  if (!definition) throw new Error('Research sheet definition is missing.');
  const actual = sheet.getRange(1, 1, 1, definition.headers.length).getDisplayValues()[0];
  for (let index = 0; index < definition.headers.length; index++) {
    if (actual[index] !== definition.headers[index]) {
      throw new Error('Research sheet headers do not match the approved schema.');
    }
  }
}