function inspectSpreadsheetSetup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Open the master spreadsheet before checking setup.');
  }

  const messages = APP_CONFIG.sheetDefinitions.map(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.name);
    if (!sheet) {
      return 'Missing tab: ' + definition.name;
    }

    const actualHeaders = sheet
      .getRange(1, 1, 1, definition.headers.length)
      .getDisplayValues()[0];
    const mismatches = definition.headers.filter(function(expected, index) {
      return actualHeaders[index] !== expected;
    });

    if (mismatches.length > 0) {
      return 'Check headers in: ' + definition.name;
    }
    return 'Ready: ' + definition.name;
  });

  return {
    spreadsheetName: spreadsheet.getName(),
    messages: messages
  };
}

function showSetupStatus() {
  const report = inspectSpreadsheetSetup();
  const message = [report.spreadsheetName, '', report.messages.join('\n')].join('\n');
  SpreadsheetApp.getUi().alert(
    APP_CONFIG.appName + ' setup status',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}