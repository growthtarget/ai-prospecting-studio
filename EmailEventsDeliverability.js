/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * EMAIL PERFORMANCE & DELIVERABILITY MODULE
 * ============================================================================
 * File: EmailEventsDeliverability.gs / EmailEventsDeliverability.js
 *
 * CAPABILITIES:
 * 1. Creates and formats the centralized "EmailEvents" data layer (11 columns)
 *    with strict validations and idempotent event deduplication.
 * 2. Creates and formats the "Email Deliverability" dashboard tab with Client
 *    and Campaign filter dropdowns, KPI tiles, rate formulas, comparison/error/
 *    daily-volume tables, and 3 native Google Sheets charts.
 * 3. Ingests events from Local Gmail, External Client Gmail Relays (v2.2+),
 *    and Microsoft 365 / Exchange Online (Graph API v1.0) using strict
 *    CLIENT_{NAME}_* isolation with zero fallback.
 * 4. Exposes queryEmailDeliverabilityMetrics(clientFilter, campaignFilter)
 *    and syncEmailDeliverabilityFromSidebar() so ControlPanel.html queries and
 *    syncs cleanly without blocking modal alerts or changing client config.
 * ============================================================================
 */

const EMAIL_EVENTS_SHEET_NAME = 'EmailEvents';
const DELIVERABILITY_DASHBOARD_SHEET_NAME = 'Email Deliverability';
const DELIVERABILITY_DELIVERY_GRACE_HOURS = 4;
const DELIVERABILITY_LOOKBACK_DAYS = 7;

const EMAIL_EVENTS_HEADERS = [
  'event_id',
  'client_id',
  'campaign_id',
  'provider',
  'provider_message_id',
  'event_type',
  'event_date',
  'recipient',
  'smtp_code',
  'bounce_type',
  'reason'
];

/**
 * Setup entry point: Creates/refreshes "EmailEvents" and "Email Deliverability"
 * tabs, formulas, dropdown validations, and embedded charts.
 */
function setupEmailDeliverabilityModule(ssOverride) {
  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;

  setupEmailEventsTab_(ss);
  setupEmailDeliverabilityDashboardTab_(ss);
  SpreadsheetApp.flush();

  try {
    ss.toast(
      'EmailEvents and Email Deliverability dashboard tabs are ready!',
      '📊 Deliverability Module Ready',
      8
    );
  } catch (e) {}
}

/**
 * Navigates directly to the "Email Deliverability" dashboard tab (creating it if needed).
 */
function openEmailDeliverabilitySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return { success: false, message: 'Spreadsheet not accessible.' };

  let dash = ss.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  if (!dash) {
    setupEmailDeliverabilityModule(ss);
    dash = ss.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  } else {
    setupEmailEventsTab_(ss);
    refreshDeliverabilityFilterDropdowns_(ss);
  }

  if (dash) {
    ss.setActiveSheet(dash);
    return { success: true, message: 'Switched to Email Deliverability.' };
  }
  return { success: false, message: 'Could not open Email Deliverability tab.' };
}

/**
 * Creates and formats the centralized "EmailEvents" sheet (11 canonical columns).
 * Preserves existing event rows cleanly.
 */
function setupEmailEventsTab_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return null;

  let sheet = spreadsheet.getSheetByName(EMAIL_EVENTS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_EVENTS_SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, EMAIL_EVENTS_HEADERS.length).setValues([EMAIL_EVENTS_HEADERS]);

  sheet.getRange(1, 1, 1, EMAIL_EVENTS_HEADERS.length)
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);

  applyDeliverabilityDropdown_(sheet.getRange('D2:D5000'), ['Gmail', 'Microsoft 365'], false);
  applyDeliverabilityDropdown_(sheet.getRange('F2:F5000'), ['Sent', 'Delivered', 'Bounced', 'Replied'], false);
  applyDeliverabilityDropdown_(sheet.getRange('J2:J5000'), ['Hard', 'Soft'], false);

  sheet.getRange('G2:G5000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('I2:I5000').setNumberFormat('@');

  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 220);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 340);

  return sheet;
}

/**
 * Creates/refreshes the "Email Deliverability" Dashboard tab, all formulas,
 * dynamic Client & Campaign filter dropdowns, and 3 embedded charts.
 */
function setupEmailDeliverabilityDashboardTab_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return null;

  let sheet = spreadsheet.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  }

  // Preserve user's active filter selections if already set
  const prevClient = String(sheet.getRange('B3').getValue() || '').trim();
  const prevCampaign = String(sheet.getRange('D3').getValue() || '').trim();

  // Remove old charts before rebuilding cleanly
  const existingCharts = sheet.getCharts();
  for (let i = 0; i < existingCharts.length; i++) {
    sheet.removeChart(existingCharts[i]);
  }

  // Break any old merges in rows 1-85 before re-laying out
  sheet.getRange('A1:J85').breakApart();
  sheet.getRange('A1:J85').clearContent().clearFormat();

  // 1. Title Banner (Row 1)
  sheet.getRange('A1:G1').merge();
  sheet.getRange('A1').setValue('📊 EMAIL PERFORMANCE & DELIVERABILITY DASHBOARD');
  sheet.getRange('H1').setValue('Last Synced:');
  sheet.getRange('I1:J1').merge();
  sheet.getRange('I1').setFormula('=IFERROR(TEXT(MAX(EmailEvents!G2:G),"yyyy-mm-dd hh:mm"),"No Events Yet")');
  sheet.getRange('A1:J1')
    .setBackground('#174ea6')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.getRange('A1').setFontSize(13);
  sheet.getRange('H1:J1').setHorizontalAlignment('right').setFontSize(10);
  sheet.setRowHeight(1, 34);

  // 2. Interactive Client & Campaign Filter Bar (Row 3)
  sheet.getRange('A3').setValue('Client Filter').setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sheet.getRange('B3').setValue(prevClient || 'All Clients').setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange('C3').setValue('Campaign Filter').setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sheet.getRange('D3').setValue(prevCampaign || 'All Campaigns').setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange('E3').setValue('Active Scope').setFontWeight('bold').setFontColor('#5f6368').setBackground('#f1f3f4');
  sheet.getRange('F3:J3').merge();
  sheet.getRange('F3').setFormula('=$B$3 & "  |  " & $D$3').setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange('A3:J3')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 28);

  // 3. Headline Tiles (Rows 5-7)
  sheet.getRange('A5:J5').merge();
  sheet.getRange('A5')
    .setValue('HEADLINE METRICS')
    .setFontWeight('bold')
    .setFontColor('#202124')
    .setBackground('#dadce0');

  const headlineCards = [
    {
      labelRange: 'A6:B6', valRange: 'A7:B7', label: 'SENT', bg: '#e8f0fe', fg: '#174ea6',
      formula: '=COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Sent")'
    },
    {
      labelRange: 'C6:D6', valRange: 'C7:D7', label: 'DELIVERED', bg: '#e6f4ea', fg: '#137333',
      formula: '=COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Delivered")'
    },
    {
      labelRange: 'E6:F6', valRange: 'E7:F7', label: 'HARD BOUNCES', bg: '#fce8e6', fg: '#c5221f',
      formula: '=COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Bounced",EmailEvents!$J$2:$J,"Hard")'
    },
    {
      labelRange: 'G6:H6', valRange: 'G7:H7', label: 'SOFT BOUNCES', bg: '#fef7e0', fg: '#b06000',
      formula: '=COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Bounced",EmailEvents!$J$2:$J,"Soft")'
    },
    {
      labelRange: 'I6:J6', valRange: 'I7:J7', label: 'REPLIES', bg: '#f3e8fd', fg: '#681da8',
      formula: '=COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Replied")'
    }
  ];

  headlineCards.forEach(function(card) {
    sheet.getRange(card.labelRange).merge()
      .setValue(card.label)
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground(card.bg)
      .setFontColor(card.fg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    const topLeftVal = card.valRange.split(':')[0];
    sheet.getRange(card.valRange).merge()
      .setBackground('#ffffff')
      .setFontWeight('bold')
      .setFontSize(20)
      .setFontColor(card.fg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0');
    sheet.getRange(topLeftVal).setFormula(card.formula);
  });
  sheet.getRange('A6:J7').setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 38);

  // 4. Performance Breakdown Tiles (Rows 9-11)
  sheet.getRange('A9:J9').merge();
  sheet.getRange('A9')
    .setValue('PERFORMANCE BREAKDOWN')
    .setFontWeight('bold')
    .setFontColor('#202124')
    .setBackground('#dadce0');

  const rateCards = [
    {
      labelRange: 'A10:C10', valRange: 'A11:C11', label: 'BOUNCE RATE ((Hard + Soft) / Sent)',
      bg: '#fce8e6', fg: '#c5221f',
      formula: '=IFERROR((E7+G7)/A7,0)'
    },
    {
      labelRange: 'D10:F10', valRange: 'D11:F11', label: 'DELIVERY RATE (Delivered / Sent)',
      bg: '#e6f4ea', fg: '#137333',
      formula: '=IFERROR(C7/A7,0)'
    },
    {
      labelRange: 'G10:J10', valRange: 'G11:J11', label: 'REPLY RATE (Replied / Delivered)',
      bg: '#e8f0fe', fg: '#174ea6',
      formula: '=IF(C7>0,IFERROR(I7/C7,0),IFERROR(I7/A7,0))'
    }
  ];

  rateCards.forEach(function(card) {
    sheet.getRange(card.labelRange).merge()
      .setValue(card.label)
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground(card.bg)
      .setFontColor(card.fg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    const topLeftVal = card.valRange.split(':')[0];
    sheet.getRange(card.valRange).merge()
      .setBackground('#ffffff')
      .setFontWeight('bold')
      .setFontSize(18)
      .setFontColor(card.fg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0.0%');
    sheet.getRange(topLeftVal).setFormula(card.formula);
  });
  sheet.getRange('A10:J11').setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(10, 22);
  sheet.setRowHeight(11, 34);

  // 5. Section 1: Campaign Comparison (Rows 13-29)
  sheet.getRange('A13:J13').merge();
  sheet.getRange('A13')
    .setValue('CAMPAIGN COMPARISON')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#202124');

  const campHeaders = [
    'Campaign ID', 'Sent', 'Delivered', 'Hard Bounces', 'Soft Bounces',
    'Replied', 'Client', 'Delivery Rate', 'Bounce Rate', 'Reply Rate'
  ];
  sheet.getRange(14, 1, 1, campHeaders.length)
    .setValues([campHeaders])
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A15').setFormula(
    '=IFERROR(ARRAY_CONSTRAIN(SORT(UNIQUE(FILTER(EmailEvents!C2:C,ARRAYFORMULA((EmailEvents!C2:C<>"")*(($B$3="All Clients")+(EmailEvents!B2:B=$B$3))*(($D$3="All Campaigns")+(EmailEvents!C2:C=$D$3))))),1,TRUE),15,1),' +
    'IFERROR(ARRAY_CONSTRAIN(SORT(UNIQUE(FILTER(Campaigns!A2:A,ARRAYFORMULA((Campaigns!A2:A<>"")*(($B$3="All Clients")+(Campaigns!C2:C=$B$3))*(($D$3="All Campaigns")+(Campaigns!A2:A=$D$3))))),1,TRUE),15,1),""))'
  );

  const campRowFormulas = [];
  for (let r = 15; r <= 29; r++) {
    campRowFormulas.push([
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,$A' + r + ',EmailEvents!$F$2:$F,"Sent"))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,$A' + r + ',EmailEvents!$F$2:$F,"Delivered"))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,$A' + r + ',EmailEvents!$F$2:$F,"Bounced",EmailEvents!$J$2:$J,"Hard"))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,$A' + r + ',EmailEvents!$F$2:$F,"Bounced",EmailEvents!$J$2:$J,"Soft"))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,$A' + r + ',EmailEvents!$F$2:$F,"Replied"))',
      '=IF($A' + r + '="","",IFERROR(INDEX(FILTER(EmailEvents!$B$2:$B,EmailEvents!$C$2:$C=$A' + r + ',EmailEvents!$B$2:$B<>""),1),IFERROR(INDEX(Campaigns!$C$2:$C,MATCH($A' + r + ',Campaigns!$A$2:$A,0)),"")))',
      '=IF($A' + r + '="","",IFERROR($C' + r + '/$B' + r + ',0))',
      '=IF($A' + r + '="","",IFERROR(($D' + r + '+$E' + r + ')/$B' + r + ',0))',
      '=IF($A' + r + '="","",IF($C' + r + '>0,IFERROR($F' + r + '/$C' + r + ',0),IFERROR($F' + r + '/$B' + r + ',0)))'
    ]);
  }
  sheet.getRange('B15:J29').setFormulas(campRowFormulas);
  sheet.getRange('B15:F29').setNumberFormat('0');
  sheet.getRange('H15:J29').setNumberFormat('0.0%');
  sheet.getRange('A15:J29').setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // 6. Section 2: SMTP Error Distribution (Rows 31-47)
  sheet.getRange('A31:E31').merge();
  sheet.getRange('A31')
    .setValue('SMTP ERROR DISTRIBUTION (BOUNCES)')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#202124');

  const smtpHeaders = ['SMTP Code', 'Occurrences', 'Bounce Type', '% of Bounces', 'Sample Reason'];
  sheet.getRange(32, 1, 1, smtpHeaders.length)
    .setValues([smtpHeaders])
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A33').setFormula(
    '=IFERROR(QUERY(FILTER({ARRAYFORMULA(IF(EmailEvents!I2:I="","UNKNOWN",EmailEvents!I2:I)),ARRAYFORMULA(IF(EmailEvents!J2:J="","Soft",EmailEvents!J2:J))},' +
    'ARRAYFORMULA((EmailEvents!F2:F="Bounced")*(($B$3="All Clients")+(EmailEvents!B2:B=$B$3))*(($D$3="All Campaigns")+(EmailEvents!C2:C=$D$3)))),' +
    '"SELECT Col1, COUNT(Col1), MAX(Col2) GROUP BY Col1 ORDER BY COUNT(Col1) DESC LIMIT 15 LABEL Col1 \'\', COUNT(Col1) \'\', MAX(Col2) \'\'",0),"")'
  );

  const smtpRowFormulas = [];
  for (let r = 33; r <= 47; r++) {
    smtpRowFormulas.push([
      '=IF($A' + r + '="","",IFERROR($B' + r + '/SUM($B$33:$B$47),0))',
      '=IF($A' + r + '="","",IFERROR(INDEX(FILTER(EmailEvents!$K$2:$K,ARRAYFORMULA((EmailEvents!$F$2:$F="Bounced")*((EmailEvents!$I$2:$I=$A' + r + ')+(($A' + r + '="UNKNOWN")*(EmailEvents!$I$2:$I="")))*(($B$3="All Clients")+(EmailEvents!$B$2:$B=$B$3))*(($D$3="All Campaigns")+(EmailEvents!$C$2:$C=$D$3)))),1),""))'
    ]);
  }
  sheet.getRange('D33:E47').setFormulas(smtpRowFormulas);
  sheet.getRange('A33:A47').setNumberFormat('@');
  sheet.getRange('B33:B47').setNumberFormat('0');
  sheet.getRange('D33:D47').setNumberFormat('0.0%');
  sheet.getRange('A33:E47').setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // 7. Section 3: Daily Sending Volume (Rows 49-80)
  sheet.getRange('A49:E49').merge();
  sheet.getRange('A49')
    .setValue('DAILY SENDING VOLUME (LAST 30 ACTIVE DAYS)')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#202124');

  const dailyHeaders = ['Date', 'Sent', 'Delivered', 'Bounced', 'Replied'];
  sheet.getRange(50, 1, 1, dailyHeaders.length)
    .setValues([dailyHeaders])
    .setFontWeight('bold')
    .setFontColor('#174ea6')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A51').setFormula(
    '=IFERROR(SORT(ARRAY_CONSTRAIN(SORT(UNIQUE(FILTER(ARRAYFORMULA(INT(EmailEvents!G2:G)),' +
    'ARRAYFORMULA(ISNUMBER(EmailEvents!G2:G)*(($B$3="All Clients")+(EmailEvents!B2:B=$B$3))*(($D$3="All Campaigns")+(EmailEvents!C2:C=$D$3))))),1,FALSE),30,1),1,TRUE),"")'
  );

  const dailyRowFormulas = [];
  for (let r = 51; r <= 80; r++) {
    dailyRowFormulas.push([
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Sent",EmailEvents!$G$2:$G,">="&$A' + r + ',EmailEvents!$G$2:$G,"<"&($A' + r + '+1)))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Delivered",EmailEvents!$G$2:$G,">="&$A' + r + ',EmailEvents!$G$2:$G,"<"&($A' + r + '+1)))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Bounced",EmailEvents!$G$2:$G,">="&$A' + r + ',EmailEvents!$G$2:$G,"<"&($A' + r + '+1)))',
      '=IF($A' + r + '="","",COUNTIFS(EmailEvents!$B$2:$B,IF($B$3="All Clients","<>",$B$3),EmailEvents!$C$2:$C,IF($D$3="All Campaigns","<>",$D$3),EmailEvents!$F$2:$F,"Replied",EmailEvents!$G$2:$G,">="&$A' + r + ',EmailEvents!$G$2:$G,"<"&($A' + r + '+1)))'
    ]);
  }
  sheet.getRange('B51:E80').setFormulas(dailyRowFormulas);
  sheet.getRange('A51:A80').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B51:E80').setNumberFormat('0');
  sheet.getRange('A51:E80').setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // Column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 145);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 155);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 24);
  sheet.setFrozenRows(3);

  // Refresh dynamic dropdowns on B3 & D3
  refreshDeliverabilityFilterDropdowns_(spreadsheet);

  // 8. Build Embedded Charts in Column L (Cols 12+)
  const campChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange('A14:F29'))
    .setNumHeaders(1)
    .setOption('title', 'Campaign Comparison (Sent, Delivered, Bounces & Replies)')
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#1a73e8', '#1e8e3e', '#d93025', '#f9ab00', '#9334e6'])
    .setOption('width', 560)
    .setOption('height', 310)
    .setPosition(13, 12, 0, 0)
    .build();
  sheet.insertChart(campChart);

  const smtpChart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange('A32:B47'))
    .setNumHeaders(1)
    .setOption('title', 'SMTP Bounce Error Distribution')
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'right' })
    .setOption('width', 560)
    .setOption('height', 310)
    .setPosition(31, 12, 0, 0)
    .build();
  sheet.insertChart(smtpChart);

  const dailyChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange('A50:E80'))
    .setNumHeaders(1)
    .setOption('title', 'Daily Sending Volume (Last 30 Active Days)')
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#1a73e8', '#1e8e3e', '#d93025', '#9334e6'])
    .setOption('width', 560)
    .setOption('height', 330)
    .setPosition(49, 12, 0, 0)
    .build();
  sheet.insertChart(dailyChart);

  return sheet;
}

/**
 * Refreshes the Client Filter (B3) and Campaign Filter (D3) dropdowns on
 * "Email Deliverability" from existing Lists & Settings, Campaigns, and EmailEvents.
 */
function refreshDeliverabilityFilterDropdowns_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const dash = spreadsheet.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  if (!dash) return;

  const clientSet = {};
  const clients = ['All Clients'];
  function addClient(c) {
    const clean = String(c || '').trim();
    if (!clean || clean === 'All Clients') return;
    const low = clean.toLowerCase();
    if (!clientSet[low]) {
      clientSet[low] = true;
      clients.push(clean);
    }
  }

  if (typeof getRegisteredClientsList_ === 'function') {
    getRegisteredClientsList_(spreadsheet).forEach(addClient);
  } else if (typeof getRegisteredClientsFromSettings_ === 'function') {
    getRegisteredClientsFromSettings_(spreadsheet).forEach(addClient);
  }

  const campaignSet = {};
  const campaigns = ['All Campaigns'];
  function addCampaign(cid) {
    const clean = String(cid || '').trim();
    if (!clean || clean === 'All Campaigns') return;
    const low = clean.toLowerCase();
    if (!campaignSet[low]) {
      campaignSet[low] = true;
      campaigns.push(clean);
    }
  }

  const campSheet = spreadsheet.getSheetByName('Campaigns');
  if (campSheet && campSheet.getLastRow() > 1) {
    const cData = campSheet.getRange(2, 1, campSheet.getLastRow() - 1, 3).getValues();
    cData.forEach(function(row) {
      addCampaign(row[0]);
      addClient(row[2]);
    });
  }

  const eventsSheet = spreadsheet.getSheetByName(EMAIL_EVENTS_SHEET_NAME);
  if (eventsSheet && eventsSheet.getLastRow() > 1) {
    const eData = eventsSheet.getRange(2, 2, eventsSheet.getLastRow() - 1, 2).getValues();
    eData.forEach(function(row) {
      addClient(row[0]);
      addCampaign(row[1]);
    });
  }

  const currentClient = String(dash.getRange('B3').getValue() || '').trim();
  if (currentClient && currentClient !== 'All Clients') addClient(currentClient);

  const currentCamp = String(dash.getRange('D3').getValue() || '').trim();
  if (currentCamp && currentCamp !== 'All Campaigns') addCampaign(currentCamp);

  applyDeliverabilityDropdown_(dash.getRange('B3'), clients, false);
  applyDeliverabilityDropdown_(dash.getRange('D3'), campaigns, false);

  if (!currentClient) dash.getRange('B3').setValue('All Clients');
  if (!currentCamp) dash.getRange('D3').setValue('All Campaigns');
}

/**
 * Syncs "Email Deliverability" tab filter cells (B3 / D3) to a Client & Campaign
 * without overriding valid options.
 */
function syncDeliverabilityTabFilters_(ss, clientFilter, campaignFilter) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!spreadsheet) return;
  const dash = spreadsheet.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME);
  if (!dash) return;

  refreshDeliverabilityFilterDropdowns_(spreadsheet);

  if (clientFilter) {
    dash.getRange('B3').setValue(String(clientFilter).trim() || 'All Clients');
  }
  if (campaignFilter) {
    dash.getRange('D3').setValue(String(campaignFilter).trim() || 'All Campaigns');
  }
}

/**
 * Server-side query API used by ControlPanel.html (and internal callers) to query
 * EmailEvents directly using the existing Client / Owner and Campaign filters.
 * Never mutates client configuration or leaks secrets.
 */
function queryEmailDeliverabilityMetrics(clientFilter, campaignFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    return { success: false, error: 'Spreadsheet not accessible.' };
  }

  const cleanClient = String(clientFilter || '').trim();
  const cleanCamp = String(campaignFilter || '').trim();
  const matchAllClients = (!cleanClient || cleanClient === 'All Clients' || cleanClient === '__ALL__');
  const matchAllCampaigns = (!cleanCamp || cleanCamp === 'All Campaigns' || cleanCamp === '__NEW_CAMPAIGN__' || cleanCamp === '__ALL__');

  const result = {
    success: true,
    clientFilter: matchAllClients ? 'All Clients' : cleanClient,
    campaignFilter: matchAllCampaigns ? 'All Campaigns' : cleanCamp,
    sent: 0,
    delivered: 0,
    hardBounces: 0,
    softBounces: 0,
    replied: 0,
    bounceRate: 0,
    deliveryRate: 0,
    replyRate: 0,
    lastEventAt: '',
    topSmtpError: ''
  };

  const sheet = ss.getSheetByName(EMAIL_EVENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return result;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, EMAIL_EVENTS_HEADERS.length).getValues();
  const smtpCounts = {};
  let maxDateMs = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rClient = String(row[1] || '').trim();
    const rCamp = String(row[2] || '').trim();
    const rType = String(row[5] || '').trim();
    const rDate = row[6];
    const rSmtp = String(row[8] || '').trim();
    const rBounce = String(row[9] || '').trim();

    if (!rType) continue;
    if (!matchAllClients && rClient.toLowerCase() !== cleanClient.toLowerCase()) continue;
    if (!matchAllCampaigns && rCamp.toLowerCase() !== cleanCamp.toLowerCase()) continue;

    if (rDate instanceof Date && !isNaN(rDate.getTime())) {
      if (rDate.getTime() > maxDateMs) maxDateMs = rDate.getTime();
    }

    if (rType === 'Sent') result.sent++;
    else if (rType === 'Delivered') result.delivered++;
    else if (rType === 'Replied') result.replied++;
    else if (rType === 'Bounced') {
      if (rBounce.toLowerCase() === 'hard') result.hardBounces++;
      else result.softBounces++;
      const codeKey = rSmtp || 'UNKNOWN';
      smtpCounts[codeKey] = (smtpCounts[codeKey] || 0) + 1;
    }
  }

  const totalBounces = result.hardBounces + result.softBounces;
  if (result.sent > 0) {
    result.bounceRate = totalBounces / result.sent;
    result.deliveryRate = result.delivered / result.sent;
  }
  if (result.delivered > 0) {
    result.replyRate = result.replied / result.delivered;
  } else if (result.sent > 0) {
    result.replyRate = result.replied / result.sent;
  }

  if (maxDateMs > 0) {
    result.lastEventAt = Utilities.formatDate(new Date(maxDateMs), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }

  let bestCode = '';
  let bestCount = 0;
  Object.keys(smtpCounts).forEach(function(code) {
    if (smtpCounts[code] > bestCount) {
      bestCount = smtpCounts[code];
      bestCode = code;
    }
  });
  if (bestCode) {
    result.topSmtpError = bestCode + ' (' + bestCount + ')';
  }

  return result;
}

/**
 * Deterministic Event ID generator ensuring idempotency across retries & hourly polls.
 */
function buildEmailEventId_(eventType, providerMessageId, recipient, fallbackSeed) {
  const prefixMap = {
    'Sent': 'EVT-SNT-',
    'Delivered': 'EVT-DLV-',
    'Bounced': 'EVT-BNC-',
    'Replied': 'EVT-RPL-'
  };
  const prefix = prefixMap[eventType] || 'EVT-';
  const rawKey = [
    String(eventType || '').trim(),
    String(providerMessageId || '').trim(),
    String(recipient || '').trim().toLowerCase(),
    String(fallbackSeed || '').trim()
  ].join('|');

  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawKey);
  const hex = bytes.map(function(b) {
    const unsigned = (b < 0) ? b + 256 : b;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('').slice(0, 18).toUpperCase();

  return prefix + hex;
}

/**
 * Normalizes provider name strictly to 'Gmail' or 'Microsoft 365' for EmailEvents Col D validation.
 */
function normalizeDeliverabilityProvider_(rawProvider) {
  const p = String(rawProvider || '').toLowerCase();
  if (p.indexOf('microsoft') >= 0 || p.indexOf('m365') >= 0 || p.indexOf('exchange') >= 0 || p.indexOf('outlook') >= 0) {
    return 'Microsoft 365';
  }
  return 'Gmail';
}

/**
 * Normalizes event_type strictly to one of: Sent, Delivered, Bounced, Replied.
 */
function normalizeDeliverabilityEventType_(rawType) {
  const t = String(rawType || '').trim().toLowerCase();
  if (t === 'sent') return 'Sent';
  if (t === 'delivered') return 'Delivered';
  if (t === 'bounced' || t === 'bounce') return 'Bounced';
  if (t === 'replied' || t === 'reply') return 'Replied';
  return 'Sent';
}

/**
 * Writes a single EmailEvent atomically (used inline by EmailSender.js on Sent).
 * Idempotent: skips if event_id already exists in EmailEvents.
 */
function recordEmailEvent_(eventObj, spreadsheet) {
  if (!eventObj) return false;
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return false;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('recordEmailEvent_ skipped: could not acquire lock.');
    return false;
  }

  try {
    const sheet = setupEmailEventsTab_(ss);
    if (!sheet) return false;

    const eventType = normalizeDeliverabilityEventType_(eventObj.event_type);
    const provider = normalizeDeliverabilityProvider_(eventObj.provider);
    const recipient = String(eventObj.recipient || '').trim().toLowerCase();
    const msgId = String(eventObj.provider_message_id || '').trim();
    const clientId = String(eventObj.client_id || 'Growth Target').trim();
    const campaignId = String(eventObj.campaign_id || '').trim();
    const eventDate = (eventObj.event_date instanceof Date && !isNaN(eventObj.event_date.getTime()))
      ? eventObj.event_date
      : new Date();
    const smtpCode = String(eventObj.smtp_code || (eventType === 'Sent' || eventType === 'Delivered' ? '250 2.0.0' : '')).trim();
    const bounceType = eventType === 'Bounced'
      ? (String(eventObj.bounce_type || '').toLowerCase() === 'hard' ? 'Hard' : 'Soft')
      : '';
    const reason = String(eventObj.reason || '').trim().slice(0, 500);

    const eventId = String(eventObj.event_id || '').trim() ||
      buildEmailEventId_(eventType, msgId || (campaignId + ':' + recipient), recipient, '');

    if (sheet.getLastRow() > 1) {
      const existingIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existingIds.length; i++) {
        if (String(existingIds[i][0] || '').trim() === eventId) {
          return false;
        }
      }
    }

    const row = [
      eventId,
      clientId,
      campaignId,
      provider,
      msgId,
      eventType,
      eventDate,
      recipient,
      smtpCode,
      bounceType,
      reason
    ];

    const targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 7).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(targetRow, 9).setNumberFormat('@');
    sheet.getRange(targetRow, 1, 1, EMAIL_EVENTS_HEADERS.length).setValues([row]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Installs an hourly time-driven trigger for syncEmailDeliverabilityEvents()
 * without creating duplicate triggers.
 */
function installDeliverabilitySyncTrigger() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  const triggers = ScriptApp.getProjectTriggers();
  let exists = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncEmailDeliverabilityEvents') {
      exists = true;
      break;
    }
  }

  if (!exists) {
    ScriptApp.newTrigger('syncEmailDeliverabilityEvents')
      .timeBased()
      .everyHours(1)
      .create();
  }

  const msg = exists
    ? 'Hourly Email Deliverability sync trigger is already active.'
    : 'Hourly Email Deliverability sync trigger installed! Events will sync automatically every hour.';

  if (ui) ui.alert('Deliverability Sync Trigger', msg, ui.ButtonSet.OK);
  else Logger.log(msg);

  return { success: true, alreadyExists: exists, message: msg };
}

/**
 * Menu & Time-Trigger Entry Point (shows summary dialog when run interactively from menu).
 */
function syncEmailDeliverabilityEvents() {
  return runEmailDeliverabilitySync_(false);
}

/**
 * Silent Sidebar Entry Point (never blocks ControlPanel.html with a modal ui.alert).
 */
function syncEmailDeliverabilityFromSidebar() {
  return runEmailDeliverabilitySync_(true);
}

/**
 * MASTER INGESTION RUNNER
 * 1. Promotes manually sent drafts from Gmail & Microsoft 365 into Sent events.
 * 2. Polls Gmail (local & ClientGmailRelay.gs) and Microsoft 365 (Graph API) per client
 *    with strict zero-fallback credential isolation.
 * 3. Parses SMTP DSN bounces (Hard/Soft) and Replies.
 * 4. Confirms Delivered events (via Exchange Message Trace or 4h no-bounce grace window),
 *    deduplicates against already-delivered message IDs, and reconciles late bounces cleanly.
 */
function runEmailDeliverabilitySync_(silentMode) {
  let ui = null;
  if (!silentMode) {
    try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return { success: false, message: 'No active spreadsheet found.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    const busyMsg = 'Another deliverability sync is currently running. Please wait a moment.';
    if (ui) ui.alert(busyMsg);
    return { success: false, message: busyMsg };
  }

  try {
    const eventsSheet = setupEmailEventsTab_(ss);
    if (!ss.getSheetByName(DELIVERABILITY_DASHBOARD_SHEET_NAME)) {
      setupEmailDeliverabilityDashboardTab_(ss);
    }

    // Load existing EmailEvents into fast lookup indexes
    const existingEventIds = {};
    const sentEvents = [];
    const bouncedByMsgId = {};
    const bouncedByRecipient = {};
    const deliveredByMsgId = {};
    const deliveredByRecipient = {};
    const assumedDeliveredRowByMsgKey = {};

    const lastEventRow = eventsSheet.getLastRow();
    if (lastEventRow > 1) {
      const existingData = eventsSheet.getRange(2, 1, lastEventRow - 1, EMAIL_EVENTS_HEADERS.length).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const row = existingData[i];
        const eId = String(row[0] || '').trim();
        const clientId = String(row[1] || '').trim();
        const campId = String(row[2] || '').trim();
        const prov = String(row[3] || '').trim();
        const msgId = String(row[4] || '').trim();
        const eType = String(row[5] || '').trim();
        const eDate = row[6];
        const recip = String(row[7] || '').trim().toLowerCase();
        const reason = String(row[10] || '').trim();

        if (eId) existingEventIds[eId] = true;

        const msgIdLow = msgId.toLowerCase();
        const recipKey = (clientId + '::' + recip).toLowerCase();

        if (eType === 'Sent') {
          sentEvents.push({
            client_id: clientId,
            campaign_id: campId,
            provider: prov,
            provider_message_id: msgId,
            event_date: (eDate instanceof Date) ? eDate : new Date(eDate),
            recipient: recip
          });
        } else if (eType === 'Bounced') {
          if (msgIdLow) bouncedByMsgId[msgIdLow] = true;
          if (recip) bouncedByRecipient[recipKey] = true;
        } else if (eType === 'Delivered') {
          if (msgIdLow) deliveredByMsgId[msgIdLow] = true;
          if (recip) deliveredByRecipient[recipKey] = true;
          if (reason.indexOf('No DSN bounce received') >= 0) {
            const msgKey = (clientId + '::' + (msgId || recip)).toLowerCase();
            assumedDeliveredRowByMsgKey[msgKey] = i + 2;
          }
        }
      }
    }

    // Collect tracked Prospect Review Queue rows (Drafts pending manual send + Sent rows)
    const queueInfo = readQueueTrackingContext_(ss);

    // Group active mailboxes by (clientName, providerType, mailbox)
    const mailboxTargets = collectActiveMailboxTargets_(ss, sentEvents, queueInfo);
    const pendingNewRows = [];
    const assumedRowsToDelete = {};
    const clientNotes = [];
    const msTokenCache = {};

    function stageNewEvent(evt) {
      const eType = normalizeDeliverabilityEventType_(evt.event_type);
      const prov = normalizeDeliverabilityProvider_(evt.provider);
      const recip = String(evt.recipient || '').trim().toLowerCase();
      const msgId = String(evt.provider_message_id || '').trim();
      const clientId = String(evt.client_id || 'Growth Target').trim();
      const campId = String(evt.campaign_id || '').trim();
      const eDate = (evt.event_date instanceof Date && !isNaN(evt.event_date.getTime()))
        ? evt.event_date
        : new Date();
      const smtpCode = String(evt.smtp_code || (eType === 'Bounced' ? 'UNKNOWN' : (eType === 'Sent' || eType === 'Delivered' ? '250 2.0.0' : ''))).trim();
      const bounceType = eType === 'Bounced'
        ? (String(evt.bounce_type || '').toLowerCase() === 'hard' ? 'Hard' : 'Soft')
        : '';
      const reason = String(evt.reason || '').trim().slice(0, 500);

      const msgIdLow = msgId.toLowerCase();
      const recipKey = (clientId + '::' + recip).toLowerCase();

      // Prevent duplicate Delivered rows for the same message (or same recipient if no message ID)
      if (eType === 'Delivered') {
        if ((msgIdLow && deliveredByMsgId[msgIdLow]) || (!msgIdLow && recip && deliveredByRecipient[recipKey])) {
          return false;
        }
      }

      const eId = String(evt.event_id || '').trim() ||
        buildEmailEventId_(eType, msgId || (campId + ':' + recip), recip, evt.fallbackSeed || '');

      if (existingEventIds[eId]) return false;
      existingEventIds[eId] = true;

      pendingNewRows.push([
        eId,
        clientId,
        campId,
        prov,
        msgId,
        eType,
        eDate,
        recip,
        smtpCode,
        bounceType,
        reason
      ]);

      if (eType === 'Sent') {
        sentEvents.push({
          client_id: clientId,
          campaign_id: campId,
          provider: prov,
          provider_message_id: msgId,
          event_date: eDate,
          recipient: recip
        });
      } else if (eType === 'Delivered') {
        if (msgIdLow) deliveredByMsgId[msgIdLow] = true;
        if (recip) deliveredByRecipient[recipKey] = true;
      } else if (eType === 'Bounced') {
        if (msgIdLow) bouncedByMsgId[msgIdLow] = true;
        if (recip) bouncedByRecipient[recipKey] = true;
        const msgKey = (clientId + '::' + (msgId || recip)).toLowerCase();
        if (assumedDeliveredRowByMsgKey[msgKey]) {
          assumedRowsToDelete[assumedDeliveredRowByMsgKey[msgKey]] = true;
        }
      }
      return true;
    }

    // Poll each isolated client mailbox target
    mailboxTargets.forEach(function(target) {
      try {
        if (target.provider === 'Microsoft 365') {
          syncMicrosoft365MailboxEvents_(
            target,
            sentEvents,
            queueInfo,
            msTokenCache,
            stageNewEvent,
            ss
          );
        } else {
          syncGmailMailboxEvents_(
            target,
            sentEvents,
            queueInfo,
            stageNewEvent,
            ss
          );
        }
      } catch (clientErr) {
        Logger.log('[Deliverability Sync] Skipped ' + target.clientName + ' (' + target.provider + '): ' + clientErr.message);
        clientNotes.push(target.clientName + ' (' + target.provider + '): ' + clientErr.message);
      }
    });

    // Apply 4-hour no-bounce delivery confirmation ONLY for Sent messages without existing Delivered/Bounced
    const graceCutoffMs = Date.now() - (DELIVERABILITY_DELIVERY_GRACE_HOURS * 60 * 60 * 1000);
    for (let i = 0; i < sentEvents.length; i++) {
      const s = sentEvents[i];
      if (!s.event_date || s.event_date.getTime() > graceCutoffMs) continue;

      const msgIdLow = String(s.provider_message_id || '').toLowerCase();
      const recipKey = (s.client_id + '::' + s.recipient).toLowerCase();
      if (msgIdLow) {
        if (bouncedByMsgId[msgIdLow] || deliveredByMsgId[msgIdLow]) continue;
      } else {
        if (bouncedByRecipient[recipKey] || deliveredByRecipient[recipKey]) continue;
      }

      stageNewEvent({
        client_id: s.client_id,
        campaign_id: s.campaign_id,
        provider: s.provider,
        provider_message_id: s.provider_message_id,
        event_type: 'Delivered',
        event_date: new Date(s.event_date.getTime() + 60000),
        recipient: s.recipient,
        smtp_code: '250 2.0.0',
        reason: 'No DSN bounce received within ' + DELIVERABILITY_DELIVERY_GRACE_HOURS + 'h window'
      });
    }

    // Delete any superseded assumed Delivered rows invalidated by late bounces (bottom-up)
    const rowsToDelete = Object.keys(assumedRowsToDelete).map(Number).sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(rNum) {
      eventsSheet.deleteRow(rNum);
    });

    // Batch append all newly discovered events
    if (pendingNewRows.length > 0) {
      const insertAt = eventsSheet.getLastRow() + 1;
      eventsSheet.getRange(insertAt, 7, pendingNewRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      eventsSheet.getRange(insertAt, 9, pendingNewRows.length, 1).setNumberFormat('@');
      eventsSheet.getRange(insertAt, 1, pendingNewRows.length, EMAIL_EVENTS_HEADERS.length).setValues(pendingNewRows);
    }

    refreshDeliverabilityFilterDropdowns_(ss);
    SpreadsheetApp.flush();

    const summary = 'Synced ' + pendingNewRows.length + ' new event(s) across ' + mailboxTargets.length + ' mailbox target(s).' +
      (clientNotes.length > 0 ? ('\nSkipped: ' + clientNotes.join(' | ')) : '');

    if (ui) ui.alert('Email Deliverability Sync Complete', summary, ui.ButtonSet.OK);
    else Logger.log(summary);

    return {
      success: true,
      addedCount: pendingNewRows.length,
      targetsChecked: mailboxTargets.length,
      notes: clientNotes,
      message: summary
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reads Prospect Review Queue rows for draft promotion & campaign attribution.
 */
function readQueueTrackingContext_(ss) {
  const ctx = {
    sheet: null,
    cols: {},
    draftRows: [],
    byMsgId: {},
    byRecipient: {}
  };

  const qSheet = ss.getSheetByName('Prospect Review Queue');
  if (!qSheet || qSheet.getLastRow() < 2) return ctx;

  ctx.sheet = qSheet;
  const headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  function findCol(names) {
    for (let i = 0; i < names.length; i++) {
      const idx = headers.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  ctx.cols = {
    campId: findCol(['Campaign ID']),
    email: findCol(['Business Email']),
    sendStatus: findCol(['Send Status']),
    sentAt: findCol(['Sent At']),
    dispatchMode: findCol(['Dispatch Mode']),
    msgId: findCol(['Thread / Message ID', 'Message ID'])
  };

  if (ctx.cols.campId < 0 || ctx.cols.email < 0) return ctx;

  const data = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, qSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const campId = String(row[ctx.cols.campId] || '').trim();
    const email = String(row[ctx.cols.email] || '').trim().toLowerCase();
    const sendStatus = ctx.cols.sendStatus >= 0 ? String(row[ctx.cols.sendStatus] || '').trim() : '';
    const dispatchMode = ctx.cols.dispatchMode >= 0 ? String(row[ctx.cols.dispatchMode] || '').trim() : '';
    const msgId = ctx.cols.msgId >= 0 ? String(row[ctx.cols.msgId] || '').trim() : '';

    if (!campId || !email) continue;

    const item = {
      rowIndex: i + 2,
      campaignId: campId,
      email: email,
      sendStatus: sendStatus,
      dispatchMode: dispatchMode,
      msgId: msgId
    };

    const stLow = sendStatus.toLowerCase();
    if ((stLow === 'draft created' || stLow === 'draft created in gmail') && msgId) {
      ctx.draftRows.push(item);
    }
    if (msgId) {
      ctx.byMsgId[msgId.toLowerCase()] = item;
    }
    ctx.byRecipient[email] = item;
  }

  return ctx;
}

/**
 * Discovers unique active (clientName, provider, mailbox) targets from Campaigns,
 * existing Sent events, and tracked Prospect Review Queue rows.
 */
function collectActiveMailboxTargets_(ss, sentEvents, queueInfo) {
  const targetsMap = {};

  function addTarget(clientName, providerRaw, mailbox, campaignId) {
    const cleanClient = String(clientName || 'Growth Target').trim();
    const normProv = normalizeDeliverabilityProvider_(providerRaw);
    const cleanBox = String(mailbox || '').trim().toLowerCase();
    const key = (cleanClient + '||' + normProv + '||' + cleanBox).toLowerCase();

    if (!targetsMap[key]) {
      targetsMap[key] = {
        clientName: cleanClient,
        provider: normProv,
        mailbox: cleanBox,
        campaignIds: {}
      };
    }
    if (campaignId) {
      targetsMap[key].campaignIds[campaignId] = true;
    }
  }

  const campSheet = ss.getSheetByName('Campaigns');
  const campMeta = {};
  if (campSheet && campSheet.getLastRow() > 1) {
    const cData = campSheet.getDataRange().getValues();
    const headers = cData[0].map(function(h) { return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    const idIdx = headers.indexOf('CAMPAIGNID') >= 0 ? headers.indexOf('CAMPAIGNID') : 0;
    const deptIdx = headers.indexOf('DEPARTMENT') >= 0 ? headers.indexOf('DEPARTMENT') : 2;
    const statusIdx = headers.indexOf('STATUS') >= 0 ? headers.indexOf('STATUS') : 3;
    const provIdx = headers.indexOf('EMAILSENDERPROVIDER') >= 0 ? headers.indexOf('EMAILSENDERPROVIDER') : 21;
    const jsonIdx = headers.indexOf('CAMPAIGNCONFIGJSON') >= 0 ? headers.indexOf('CAMPAIGNCONFIGJSON') : 26;

    for (let r = 1; r < cData.length; r++) {
      const cid = String(cData[r][idIdx] || '').trim();
      if (!cid) continue;
      const clientName = String(cData[r][deptIdx] || 'Growth Target').trim();
      const status = String(cData[r][statusIdx] || '').trim().toLowerCase();
      const prov = String(cData[r][provIdx] || '').trim();
      let mailbox = '';
      if (jsonIdx >= 0 && cData[r][jsonIdx]) {
        try {
          const parsed = JSON.parse(cData[r][jsonIdx]);
          if (parsed && parsed.senderEmail) mailbox = String(parsed.senderEmail).trim();
        } catch (e) {}
      }

      campMeta[cid] = { clientName: clientName, provider: prov, mailbox: mailbox };

      if (status !== 'paused' && status !== 'archived') {
        const pLow = prov.toLowerCase();
        if (!prov || pLow.indexOf('gmail') >= 0 || pLow.indexOf('workspace') >= 0 || pLow.indexOf('draft') >= 0 || pLow.indexOf('direct') >= 0 || pLow.indexOf('microsoft') >= 0 || pLow.indexOf('m365') >= 0) {
          addTarget(clientName, prov || 'Gmail', mailbox, cid);
        }
      }
    }
  }

  sentEvents.forEach(function(s) {
    const meta = campMeta[s.campaign_id] || {};
    addTarget(s.client_id || meta.clientName || 'Growth Target', s.provider || meta.provider || 'Gmail', meta.mailbox || '', s.campaign_id);
  });

  (queueInfo.draftRows || []).forEach(function(d) {
    const meta = campMeta[d.campaignId] || {};
    addTarget(meta.clientName || 'Growth Target', d.dispatchMode || meta.provider || 'Gmail', meta.mailbox || '', d.campaignId);
  });

  return Object.keys(targetsMap).map(function(k) { return targetsMap[k]; });
}

/**
 * Polls Gmail events (Local Agency GmailApp OR External Client GmailRelay.gs v2.2+).
 */
function syncGmailMailboxEvents_(target, sentEvents, queueInfo, stageNewEvent, ss) {
  const clientName = target.clientName;
  const isLocal = typeof isLocalAgencyClient_ === 'function' ? isLocalAgencyClient_(clientName) : true;
  const clientSent = sentEvents.filter(function(s) {
    return s.client_id.toLowerCase() === clientName.toLowerCase() && s.provider === 'Gmail';
  });
  const trackedRecipients = {};
  const sentByMsgId = {};
  const sentByRecipient = {};

  clientSent.forEach(function(s) {
    if (s.recipient) {
      trackedRecipients[s.recipient] = true;
      if (!sentByRecipient[s.recipient] || s.event_date > sentByRecipient[s.recipient].event_date) {
        sentByRecipient[s.recipient] = s;
      }
    }
    if (s.provider_message_id) {
      sentByMsgId[s.provider_message_id.toLowerCase()] = s;
    }
  });

  const clientDrafts = (queueInfo.draftRows || []).filter(function(d) {
    return target.campaignIds[d.campaignId] && normalizeDeliverabilityProvider_(d.dispatchMode || 'Gmail') === 'Gmail';
  });

  if (isLocal) {
    // 1. Promote any local Gmail drafts that the user sent manually
    clientDrafts.forEach(function(draftItem) {
      try {
        const draftObj = GmailApp.getDraft(draftItem.msgId);
        if (!draftObj) {
          const sentThreads = GmailApp.search('in:sent to:' + draftItem.email + ' newer_than:' + DELIVERABILITY_LOOKBACK_DAYS + 'd', 0, 5);
          if (sentThreads && sentThreads.length > 0) {
            const msgs = sentThreads[0].getMessages();
            const lastMsg = msgs[msgs.length - 1];
            let rfcId = '';
            try { rfcId = lastMsg ? (lastMsg.getHeader('Message-ID') || '') : ''; } catch (e) {}
            const sentMsgId = rfcId || (lastMsg ? lastMsg.getId() : draftItem.msgId);
            const sentAt = lastMsg ? lastMsg.getDate() : new Date();

            promoteQueueDraftToSent_(queueInfo, draftItem, sentMsgId, sentAt);
            stageNewEvent({
              client_id: clientName,
              campaign_id: draftItem.campaignId,
              provider: 'Gmail',
              provider_message_id: sentMsgId,
              event_type: 'Sent',
              event_date: sentAt,
              recipient: draftItem.email,
              smtp_code: '250 2.0.0',
              reason: 'Promoted from Gmail Draft to Sent'
            });
            trackedRecipients[draftItem.email] = true;
            sentByMsgId[sentMsgId.toLowerCase()] = { campaign_id: draftItem.campaignId, recipient: draftItem.email, event_date: sentAt, provider_message_id: sentMsgId };
          }
        }
      } catch (draftErr) {
        Logger.log('Local draft check notice (' + draftItem.msgId + '): ' + draftErr.message);
      }
    });

    // 2. Poll Gmail DSN bounce messages
    const bounceThreads = GmailApp.search('from:(mailer-daemon OR postmaster) newer_than:' + DELIVERABILITY_LOOKBACK_DAYS + 'd', 0, 50);
    bounceThreads.forEach(function(thread) {
      const msgs = thread.getMessages();
      msgs.forEach(function(msg) {
        let raw = '';
        try { raw = msg.getRawContent(); } catch (e) { raw = msg.getPlainBody() || ''; }
        const parsedBounce = parseSmtpBounceDetails_(msg.getSubject(), msg.getPlainBody() || '', raw);
        const recip = parsedBounce.recipient;
        if (!recip) return;

        const matchSent = (parsedBounce.originalMessageId && sentByMsgId[parsedBounce.originalMessageId.toLowerCase()]) ||
          sentByRecipient[recip] ||
          (queueInfo.byRecipient[recip] && target.campaignIds[queueInfo.byRecipient[recip].campaignId] ? { campaign_id: queueInfo.byRecipient[recip].campaignId } : null);
        if (!matchSent) return;

        stageNewEvent({
          client_id: clientName,
          campaign_id: matchSent.campaign_id,
          provider: 'Gmail',
          provider_message_id: parsedBounce.originalMessageId || matchSent.provider_message_id || msg.getId(),
          event_type: 'Bounced',
          event_date: msg.getDate(),
          recipient: recip,
          smtp_code: parsedBounce.smtpCode,
          bounce_type: parsedBounce.bounceType,
          reason: parsedBounce.reason,
          fallbackSeed: msg.getId()
        });
      });
    });

    // 3. Poll prospect replies matching tracked recipients
    const recipList = Object.keys(trackedRecipients);
    if (recipList.length > 0) {
      const myEmail = (function() {
        try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
      })();

      for (let idx = 0; idx < recipList.length; idx += 15) {
        const chunk = recipList.slice(idx, idx + 15);
        const query = 'in:anywhere newer_than:' + DELIVERABILITY_LOOKBACK_DAYS + 'd -from:mailer-daemon -from:postmaster (' +
          chunk.map(function(e) { return 'from:' + e; }).join(' OR ') + ')';
        const replyThreads = GmailApp.search(query, 0, 50);
        replyThreads.forEach(function(thread) {
          thread.getMessages().forEach(function(msg) {
            const sender = extractEmailAddress_(msg.getFrom());
            if (!sender || sender === myEmail || !trackedRecipients[sender]) return;
            const matchSent = sentByRecipient[sender] || { campaign_id: (queueInfo.byRecipient[sender] ? queueInfo.byRecipient[sender].campaignId : '') };
            if (!matchSent || !matchSent.campaign_id) return;

            stageNewEvent({
              client_id: clientName,
              campaign_id: matchSent.campaign_id,
              provider: 'Gmail',
              provider_message_id: msg.getId(),
              event_type: 'Replied',
              event_date: msg.getDate(),
              recipient: sender,
              smtp_code: '',
              reason: 'Reply received: ' + String(msg.getSubject() || '').slice(0, 120),
              fallbackSeed: msg.getId()
            });
          });
        });
      }
    }
    return;
  }

  // EXTERNAL CLIENT: Strict zero-fallback ClientGmailRelay.gs pollEvents call
  if (typeof resolveClientScopedSecret_ !== 'function') {
    throw new Error('resolveClientScopedSecret_ helper not found.');
  }
  const urlRes = resolveClientScopedSecret_(['GMAIL_RELAY_URL'], { clientName: clientName }, { throwOnMissing: true, spreadsheet: ss });
  const secRes = resolveClientScopedSecret_(['GMAIL_RELAY_SECRET'], { clientName: clientName }, { throwOnMissing: true, spreadsheet: ss });

  if (!/^https:\/\//i.test(urlRes.key)) {
    throw new Error('Invalid GMAIL_RELAY_URL for "' + clientName + '".');
  }

  const pollPayload = {
    secret: secRes.key,
    action: 'pollEvents',
    lookbackDays: DELIVERABILITY_LOOKBACK_DAYS,
    trackedDrafts: clientDrafts.map(function(d) { return { draftId: d.msgId, email: d.email, campaignId: d.campaignId }; }),
    trackedRecipients: Object.keys(trackedRecipients),
    clientName: clientName
  };

  const res = UrlFetchApp.fetch(urlRes.key, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(pollPayload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const bodyText = res.getContentText() || '';
  if (code >= 300) {
    throw new Error('Relay HTTP ' + code + ': ' + bodyText.slice(0, 180));
  }

  let parsed = {};
  try { parsed = JSON.parse(bodyText); } catch (e) {
    throw new Error('Relay returned non-JSON response during pollEvents.');
  }
  if (!parsed || parsed.success !== true) {
    throw new Error(parsed && parsed.error ? parsed.error : 'Relay pollEvents failed (upgrade ClientGmailRelay.gs to v2.2+).');
  }

  // Process promoted drafts from client Workspace
  (parsed.promotedDrafts || []).forEach(function(pd) {
    const matchDraft = clientDrafts.find(function(d) { return d.msgId === pd.draftId || d.email === String(pd.recipient || '').toLowerCase(); });
    const campId = (pd.campaignId || (matchDraft ? matchDraft.campaignId : ''));
    const recip = String(pd.recipient || (matchDraft ? matchDraft.email : '')).toLowerCase();
    const sentAt = pd.eventDate ? new Date(pd.eventDate) : new Date();
    const sentMsgId = String(pd.messageId || pd.draftId || '');

    if (matchDraft) {
      promoteQueueDraftToSent_(queueInfo, matchDraft, sentMsgId, sentAt);
    }
    if (campId && recip) {
      stageNewEvent({
        client_id: clientName,
        campaign_id: campId,
        provider: 'Gmail',
        provider_message_id: sentMsgId,
        event_type: 'Sent',
        event_date: sentAt,
        recipient: recip,
        smtp_code: '250 2.0.0',
        reason: 'Promoted from Client Gmail Draft to Sent'
      });
      trackedRecipients[recip] = true;
      sentByRecipient[recip] = { campaign_id: campId, recipient: recip, provider_message_id: sentMsgId, event_date: sentAt };
    }
  });

  // Process bounces from client Workspace
  (parsed.bounces || []).forEach(function(b) {
    const recip = String(b.recipient || '').toLowerCase();
    const matchSent = (b.originalMessageId && sentByMsgId[String(b.originalMessageId).toLowerCase()]) || sentByRecipient[recip];
    if (!matchSent && (!recip || !queueInfo.byRecipient[recip])) return;
    const campId = matchSent ? matchSent.campaign_id : queueInfo.byRecipient[recip].campaignId;

    stageNewEvent({
      client_id: clientName,
      campaign_id: campId,
      provider: 'Gmail',
      provider_message_id: b.originalMessageId || (matchSent ? matchSent.provider_message_id : '') || b.messageId || '',
      event_type: 'Bounced',
      event_date: b.eventDate ? new Date(b.eventDate) : new Date(),
      recipient: recip,
      smtp_code: b.smtpCode || 'UNKNOWN',
      bounce_type: b.bounceType || 'Soft',
      reason: b.reason || 'Bounce reported by Client Gmail Relay',
      fallbackSeed: b.messageId || ''
    });
  });

  // Process replies from client Workspace
  (parsed.replies || []).forEach(function(r) {
    const recip = String(r.recipient || '').toLowerCase();
    const matchSent = sentByRecipient[recip];
    const campId = matchSent ? matchSent.campaign_id : (queueInfo.byRecipient[recip] ? queueInfo.byRecipient[recip].campaignId : '');
    if (!campId || !recip) return;

    stageNewEvent({
      client_id: clientName,
      campaign_id: campId,
      provider: 'Gmail',
      provider_message_id: r.messageId || '',
      event_type: 'Replied',
      event_date: r.eventDate ? new Date(r.eventDate) : new Date(),
      recipient: recip,
      smtp_code: '',
      reason: r.reason || ('Reply received: ' + String(r.subject || '').slice(0, 120)),
      fallbackSeed: r.messageId || ''
    });
  });
}

/**
 * Polls Microsoft 365 / Exchange Online events via Microsoft Graph API v1.0
 * using strict CLIENT_{NAME}_MS_* credentials (or local MS_* for Growth Target).
 */
function syncMicrosoft365MailboxEvents_(target, sentEvents, queueInfo, msTokenCache, stageNewEvent, ss) {
  const clientName = target.clientName;
  const mailbox = target.mailbox;
  if (!mailbox || mailbox.indexOf('@') < 1) {
    throw new Error('No Sender Email / Mailbox configured for Microsoft 365 campaign.');
  }

  const tenantId = getSenderApiKey_('MS_TENANT_ID', ss, { clientName: clientName });
  const clientId = getSenderApiKey_('MS_CLIENT_ID', ss, { clientName: clientName });
  const clientSecret = getSenderApiKey_('MS_CLIENT_SECRET', ss, { clientName: clientName });

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET for "' + clientName + '".');
  }

  const accessToken = getMicrosoftGraphAccessToken_(tenantId, clientId, clientSecret, clientName, msTokenCache);

  const clientSent = sentEvents.filter(function(s) {
    return s.client_id.toLowerCase() === clientName.toLowerCase() && s.provider === 'Microsoft 365';
  });
  const sentByMsgId = {};
  const sentByRecipient = {};
  const trackedRecipients = {};

  clientSent.forEach(function(s) {
    if (s.recipient) {
      trackedRecipients[s.recipient] = true;
      if (!sentByRecipient[s.recipient] || s.event_date > sentByRecipient[s.recipient].event_date) {
        sentByRecipient[s.recipient] = s;
      }
    }
    if (s.provider_message_id) {
      sentByMsgId[s.provider_message_id.toLowerCase()] = s;
    }
  });

  // 1. Check tracked Microsoft 365 Drafts for manual sends (isDraft: false)
  const clientDrafts = (queueInfo.draftRows || []).filter(function(d) {
    return target.campaignIds[d.campaignId] && normalizeDeliverabilityProvider_(d.dispatchMode || '') === 'Microsoft 365';
  });

  clientDrafts.forEach(function(draftItem) {
    try {
      const msgUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(mailbox) +
        '/messages/' + encodeURIComponent(draftItem.msgId) +
        '?$select=id,isDraft,sentDateTime,internetMessageId,conversationId';
      const dRes = UrlFetchApp.fetch(msgUrl, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + accessToken },
        muteHttpExceptions: true
      });
      if (dRes.getResponseCode() === 200) {
        const dJson = JSON.parse(dRes.getContentText() || '{}');
        if (dJson.isDraft === false) {
          const sentAt = dJson.sentDateTime ? new Date(dJson.sentDateTime) : new Date();
          const finalMsgId = dJson.internetMessageId || dJson.id || draftItem.msgId;
          promoteQueueDraftToSent_(queueInfo, draftItem, finalMsgId, sentAt);
          stageNewEvent({
            client_id: clientName,
            campaign_id: draftItem.campaignId,
            provider: 'Microsoft 365',
            provider_message_id: finalMsgId,
            event_type: 'Sent',
            event_date: sentAt,
            recipient: draftItem.email,
            smtp_code: '250 2.0.0',
            reason: 'Promoted from Microsoft 365 Draft to Sent'
          });
          trackedRecipients[draftItem.email] = true;
          sentByMsgId[finalMsgId.toLowerCase()] = { campaign_id: draftItem.campaignId, recipient: draftItem.email, event_date: sentAt, provider_message_id: finalMsgId };
          sentByRecipient[draftItem.email] = sentByMsgId[finalMsgId.toLowerCase()];
        }
      }
    } catch (e) {
      Logger.log('M365 draft check notice (' + draftItem.msgId + '): ' + e.message);
    }
  });

  // 2. Optional Native Exchange Message Trace (Graph v1.0) if client granted ExchangeMessageTrace.Read.All
  const nowDate = new Date();
  const startIso = new Date(nowDate.getTime() - (DELIVERABILITY_LOOKBACK_DAYS * 24 * 60 * 60 * 1000)).toISOString();
  const traceStartIso = startIso.replace(/\.\d{3}Z$/, 'Z');
  const traceEndIso = nowDate.toISOString().replace(/\.\d{3}Z$/, 'Z');

  try {
    const traceFilter = "receivedDateTime ge " + traceStartIso + " and receivedDateTime le " + traceEndIso +
      " and senderAddress eq '" + mailbox.replace(/'/g, "''") + "'";
    const traceUrl = 'https://graph.microsoft.com/v1.0/admin/exchange/tracing/messageTraces?$filter=' +
      encodeURIComponent(traceFilter) + '&$top=1000';
    const traceRes = UrlFetchApp.fetch(traceUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    if (traceRes.getResponseCode() === 200) {
      const traceJson = JSON.parse(traceRes.getContentText() || '{}');
      const traces = Array.isArray(traceJson.value) ? traceJson.value : [];
      traces.forEach(function(tr) {
        const recip = String(tr.recipientAddress || '').trim().toLowerCase();
        const internetId = String(tr.messageId || '').trim();
        const status = String(tr.status || '').trim().toLowerCase();
        const matchSent = (internetId && sentByMsgId[internetId.toLowerCase()]) || sentByRecipient[recip];
        if (!matchSent) return;

        const evDate = tr.receivedDateTime ? new Date(tr.receivedDateTime) : new Date();
        if (status === 'delivered') {
          stageNewEvent({
            client_id: clientName,
            campaign_id: matchSent.campaign_id,
            provider: 'Microsoft 365',
            provider_message_id: internetId || matchSent.provider_message_id,
            event_type: 'Delivered',
            event_date: evDate,
            recipient: recip,
            smtp_code: '250 2.0.0',
            reason: 'Confirmed Delivered via Exchange Message Trace'
          });
        } else if (status === 'failed' || status === 'quarantined' || status === 'filteredasspam') {
          stageNewEvent({
            client_id: clientName,
            campaign_id: matchSent.campaign_id,
            provider: 'Microsoft 365',
            provider_message_id: internetId || matchSent.provider_message_id,
            event_type: 'Bounced',
            event_date: evDate,
            recipient: recip,
            smtp_code: status === 'quarantined' ? '550 5.7.1' : '550 5.0.0',
            bounce_type: 'Hard',
            reason: 'Exchange Trace status: ' + tr.status,
            fallbackSeed: tr.id || internetId
          });
        }
      });
    }
  } catch (traceErr) {
    Logger.log('Exchange Message Trace skipped (' + clientName + '): ' + traceErr.message);
  }

  // 3. Poll Mailbox Messages via Mail.ReadWrite for Exchange NDR Bounces & Prospect Replies
  let nextUrl = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(mailbox) +
    '/messages?$filter=' + encodeURIComponent('receivedDateTime ge ' + startIso) +
    '&$top=100&$select=id,subject,from,receivedDateTime,conversationId,internetMessageId,bodyPreview,body';

  let pageCount = 0;
  while (nextUrl && pageCount < 3) {
    pageCount++;
    const msgRes = UrlFetchApp.fetch(nextUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    if (msgRes.getResponseCode() !== 200) {
      throw new Error(parseMicrosoftGraphError_(msgRes.getResponseCode(), msgRes.getContentText(), 'read mailbox messages', mailbox));
    }

    const msgJson = JSON.parse(msgRes.getContentText() || '{}');
    const messages = Array.isArray(msgJson.value) ? msgJson.value : [];

    messages.forEach(function(msg) {
      const fromAddr = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '').trim().toLowerCase();
      if (!fromAddr || fromAddr === mailbox) return;

      const subj = String(msg.subject || '');
      const subjLow = subj.toLowerCase();
      const bodyText = stripHtmlTags_((msg.body && msg.body.content) || msg.bodyPreview || '');
      const msgDate = msg.receivedDateTime ? new Date(msg.receivedDateTime) : new Date();

      const isNdr = (
        fromAddr.indexOf('postmaster') >= 0 ||
        fromAddr.indexOf('mailer-daemon') >= 0 ||
        fromAddr.indexOf('microsoftexchange') >= 0 ||
        subjLow.indexOf('undeliverable') === 0 ||
        subjLow.indexOf('delivery status notification') >= 0 ||
        subjLow.indexOf('mail delivery failed') >= 0 ||
        subjLow.indexOf('returned mail') >= 0
      );

      if (isNdr) {
        const parsedBounce = parseSmtpBounceDetails_(subj, bodyText, bodyText);
        const recip = parsedBounce.recipient;
        if (!recip) return;

        const matchSent = (parsedBounce.originalMessageId && sentByMsgId[parsedBounce.originalMessageId.toLowerCase()]) ||
          sentByRecipient[recip] ||
          (queueInfo.byRecipient[recip] && target.campaignIds[queueInfo.byRecipient[recip].campaignId] ? { campaign_id: queueInfo.byRecipient[recip].campaignId } : null);
        if (!matchSent) return;

        stageNewEvent({
          client_id: clientName,
          campaign_id: matchSent.campaign_id,
          provider: 'Microsoft 365',
          provider_message_id: parsedBounce.originalMessageId || matchSent.provider_message_id || msg.internetMessageId || msg.id,
          event_type: 'Bounced',
          event_date: msgDate,
          recipient: recip,
          smtp_code: parsedBounce.smtpCode,
          bounce_type: parsedBounce.bounceType,
          reason: parsedBounce.reason,
          fallbackSeed: msg.id
        });
      } else if (trackedRecipients[fromAddr]) {
        const matchSent = sentByRecipient[fromAddr] || { campaign_id: (queueInfo.byRecipient[fromAddr] ? queueInfo.byRecipient[fromAddr].campaignId : '') };
        if (!matchSent || !matchSent.campaign_id) return;

        stageNewEvent({
          client_id: clientName,
          campaign_id: matchSent.campaign_id,
          provider: 'Microsoft 365',
          provider_message_id: msg.internetMessageId || msg.id,
          event_type: 'Replied',
          event_date: msgDate,
          recipient: fromAddr,
          smtp_code: '',
          reason: 'Reply received: ' + subj.slice(0, 120),
          fallbackSeed: msg.id
        });
      }
    });

    nextUrl = msgJson['@odata.nextLink'] || '';
  }
}

/**
 * Promotes a manually-sent Draft row in Prospect Review Queue to 'Sent'.
 */
function promoteQueueDraftToSent_(queueInfo, draftItem, sentMsgId, sentAt) {
  if (!queueInfo || !queueInfo.sheet || !draftItem || !draftItem.rowIndex) return;
  try {
    const sh = queueInfo.sheet;
    if (queueInfo.cols.sendStatus >= 0) {
      sh.getRange(draftItem.rowIndex, queueInfo.cols.sendStatus + 1).setValue('Sent');
    }
    if (queueInfo.cols.sentAt >= 0) {
      sh.getRange(draftItem.rowIndex, queueInfo.cols.sentAt + 1).setValue((sentAt || new Date()).toLocaleString());
    }
    if (queueInfo.cols.msgId >= 0 && sentMsgId) {
      sh.getRange(draftItem.rowIndex, queueInfo.cols.msgId + 1).setValue(sentMsgId);
    }
  } catch (e) {
    Logger.log('Could not promote queue draft row ' + draftItem.rowIndex + ': ' + e.message);
  }
}

/**
 * Parses Gmail DSN / Exchange NDR messages to extract recipient, SMTP code,
 * Hard vs Soft bounce classification, original Message-ID, and clean reason.
 */
function parseSmtpBounceDetails_(subject, plainBody, rawContent) {
  const combined = String(plainBody || '') + '\n' + String(rawContent || '');

  // 1. Extract recipient address
  let recipient = '';
  const recipPatterns = [
    /Final-Recipient:\s*(?:rfc822;)?\s*([^\s<>;"'\r\n]+@[^\s<>;"'\r\n]+)/i,
    /Original-Recipient:\s*(?:rfc822;)?\s*([^\s<>;"'\r\n]+@[^\s<>;"'\r\n]+)/i,
    /X-Failed-Recipients:\s*([^\s<>;"',\r\n]+@[^\s<>;"',\r\n]+)/i,
    /(?:delivery to the following recipient failed permanently|wasn't delivered to|undeliverable to|could not be delivered to)[:\s]+([^\s<>;"'\r\n]+@[^\s<>;"'\r\n]+)/i,
    /To:\s*([^\s<>;"'\r\n]+@[^\s<>;"'\r\n]+)/i
  ];
  for (let i = 0; i < recipPatterns.length; i++) {
    const m = combined.match(recipPatterns[i]);
    if (m && m[1]) {
      const candidate = m[1].trim().toLowerCase().replace(/[.,;>]+$/, '');
      if (candidate.indexOf('mailer-daemon') < 0 && candidate.indexOf('postmaster') < 0) {
        recipient = candidate;
        break;
      }
    }
  }

  // 2. Extract original Message-ID
  let originalMessageId = '';
  const msgIdMatch = combined.match(/(?:In-Reply-To|Original-Message-ID|Message-ID):\s*(<[^>]+>)/i);
  if (msgIdMatch && msgIdMatch[1]) {
    originalMessageId = msgIdMatch[1].trim();
  }

  // 3. Extract SMTP Code (e.g. "550 5.1.1", "451 4.7.1", "554", etc.)
  let smtpCode = 'UNKNOWN';
  const smtpFullMatch = combined.match(/\b([45]\d{2})[\s-]+([45]\.\d{1,3}\.\d{1,3})\b/);
  if (smtpFullMatch) {
    smtpCode = smtpFullMatch[1] + ' ' + smtpFullMatch[2];
  } else {
    const statusMatch = combined.match(/Status:\s*([45]\.\d{1,3}\.\d{1,3})/i);
    const code3Match = combined.match(/(?:Diagnostic-Code:[^\r\n]*|Remote server returned[^\r\n]*|SMTP error[^\r\n]*)\b([45]\d{2})\b/i) ||
      combined.match(/\b([45][0-5]\d)\b/);
    if (code3Match && statusMatch) {
      smtpCode = code3Match[1] + ' ' + statusMatch[1];
    } else if (code3Match) {
      smtpCode = code3Match[1];
    } else if (statusMatch) {
      smtpCode = statusMatch[1];
    }
  }

  // 4. Extract human-readable diagnostic reason
  let reason = '';
  const diagMatch = combined.match(/Diagnostic-Code:\s*(?:smtp;)?\s*([^\r\n]+(?:\r?\n[ \t]+[^\r\n]+)?)/i) ||
    combined.match(/Remote server returned\s*'([^']+)'/i);
  if (diagMatch && diagMatch[1]) {
    reason = diagMatch[1].replace(/\s+/g, ' ').trim();
  } else {
    reason = String(subject || '').trim() || String(plainBody || '').replace(/\s+/g, ' ').trim().slice(0, 220);
  }

  // 5. Classify Hard vs Soft bounce
  const codeLead = smtpCode.charAt(0);
  const reasonLow = (reason + ' ' + String(subject || '')).toLowerCase();
  let bounceType = 'Soft';

  if (codeLead === '4' ||
      smtpCode.indexOf('552') === 0 ||
      smtpCode.indexOf('5.2.2') >= 0 ||
      smtpCode.indexOf('5.4.7') >= 0 ||
      reasonLow.indexOf('mailbox full') >= 0 ||
      reasonLow.indexOf('over quota') >= 0 ||
      reasonLow.indexOf('temporarily') >= 0 ||
      reasonLow.indexOf('try again') >= 0 ||
      reasonLow.indexOf('rate limit') >= 0 ||
      reasonLow.indexOf('throttl') >= 0) {
    bounceType = 'Soft';
  } else if (codeLead === '5' ||
      reasonLow.indexOf('user unknown') >= 0 ||
      reasonLow.indexOf('does not exist') >= 0 ||
      reasonLow.indexOf('recipientnotfound') >= 0 ||
      reasonLow.indexOf('bad-mailbox') >= 0 ||
      reasonLow.indexOf('no such user') >= 0 ||
      reasonLow.indexOf('domain not found') >= 0 ||
      reasonLow.indexOf('permanent') >= 0) {
    bounceType = 'Hard';
  }

  return {
    recipient: recipient,
    originalMessageId: originalMessageId,
    smtpCode: smtpCode,
    bounceType: bounceType,
    reason: reason.slice(0, 350)
  };
}

function extractEmailAddress_(headerStr) {
  const s = String(headerStr || '').trim();
  const angle = s.match(/<([^>]+@[^>]+)>/);
  if (angle && angle[1]) return angle[1].trim().toLowerCase();
  const plain = s.match(/([^\s<>;"']+@[^\s<>;"']+)/);
  return plain && plain[1] ? plain[1].trim().toLowerCase() : '';
}

function stripHtmlTags_(html) {
  return String(html || '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;|&#39;/gi, "'")
    .replace(/&amp;/gi, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

function applyDeliverabilityDropdown_(range, items, allowInvalid) {
  if (!range || !Array.isArray(items) || items.length === 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  range.setDataValidation(rule);
}