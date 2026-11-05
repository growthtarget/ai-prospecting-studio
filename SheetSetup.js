function inspectSpreadsheetSetup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Open the master spreadsheet before checking setup.');
  }

  const messages = APP_CONFIG.sheetDefinitions.map(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.name);
    if (!sheet) {
      return '❌ Missing tab: ' + definition.name;
    }

    // Safety Fix 1: Ensure sheet has enough columns to hold all required headers
    const currentMaxCols = sheet.getMaxColumns();
    const neededCols = definition.headers.length;
    if (currentMaxCols < neededCols) {
      sheet.insertColumnsAfter(currentMaxCols, neededCols - currentMaxCols);
    }

    // Read the actual headers safely
    const actualHeaders = sheet
      .getRange(1, 1, 1, neededCols)
      .getDisplayValues()[0];

    // Safety Fix 2: Trim whitespace and compare cleanly
    const mismatches = [];
    definition.headers.forEach(function(expected, index) {
      const actual = String(actualHeaders[index] || '').trim();
      const expectedClean = String(expected || '').trim();
      if (actual !== expectedClean) {
        mismatches.push('Col ' + (index + 1) + ' (Expected: "' + expectedClean + '", Found: "' + (actual || 'EMPTY') + '")');
      }
    });

    if (mismatches.length > 0) {
      return '⚠️ Check headers in ' + definition.name + ':\n   - ' + mismatches.join('\n   - ');
    }

    return '✅ Ready: ' + definition.name;
  });

  return {
    spreadsheetName: spreadsheet.getName(),
    messages: messages
  };
}

function showSetupStatus() {
  const report = inspectSpreadsheetSetup();
  const message = [
    'Master Spreadsheet: ' + report.spreadsheetName,
    '----------------------------------------',
    report.messages.join('\n\n')
  ].join('\n');

  SpreadsheetApp.getUi().alert(
    APP_CONFIG.appName + ' Setup Status',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}