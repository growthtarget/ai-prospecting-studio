function showPersonaMapperControlPanel() {
  const html = HtmlService.createHtmlOutputFromFile('ControlPanel').setTitle('Persona Mapper');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPersonaMapperControlPanelData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaigns = getSheetRecords_(ss.getSheetByName('Campaigns'));
  const companies = getSheetRecords_(ss.getSheetByName('Company Database'));
  const campaignConfigs = getSheetRecords_(ss.getSheetByName('Campaign Configuration'));
  const enriched = getSheetRecords_(ss.getSheetByName('Company Intelligence'));
  const personas = getSheetRecords_(ss.getSheetByName('Persona Leads'));
  const queue = getSheetRecords_(ss.getSheetByName('Prospect Review Queue'));
  const settings = getSheetRecords_(ss.getSheetByName('Persona Mapper Settings'));
  const settingsById = {};
  settings.forEach(function(r){ settingsById[String(r['Campaign ID'] || '')] = r; });
  function uniqueCount(rows,key){ return new Set(rows.map(function(r){return String(r[key] || '').trim();}).filter(Boolean)).size; }
  const processedIds = new Set(enriched.map(function(r){return String(r['Campaign ID'] || '');}).filter(Boolean));
  const pending = queue.filter(function(r){return !['Approved for Outreach','Rejected'].includes(String(r['Review Status'] || ''));});
  const metrics = {
    campaignsProcessed: processedIds.size,
    companiesEnriched: enriched.length,
    personasIdentified: personas.length,
    qualifiedPersonas: personas.filter(function(r){return confidenceRank_(r['Confidence']) >= 1 && Number(r['Persona Fit Score'] || 0) >= 50;}).length,
    emailsDiscovered: personas.filter(function(r){return String(r['Business Email'] || '').trim();}).length,
    emailsVerified: queue.filter(function(r){return String(r['Email Verification Status']) === 'Verified – Deliverable';}).length,
    emailsRiskyUnknown: queue.filter(function(r){return ['Risky','Unknown','Inferred – Unverified','Not Found'].includes(String(r['Email Verification Status'] || ''));}).length,
    awaitingReview: pending.length,
    approved: queue.filter(function(r){return String(r['Human Approval Status']) === 'Approved' && String(r['Review Status']) === 'Approved for Outreach';}).length,
    rejected: queue.filter(function(r){return String(r['Review Status']) === 'Rejected';}).length
  };
  const campaignRows = campaigns.map(function(r){
    const id = String(r['Campaign ID'] || '').trim();
    const config = settingsById[id] || {};
    const campaignDetail = campaignConfigs.find(function(x){return String(x['Campaign ID'] || '').trim() === id;}) || {};
    const source = companies.find(function(c){return String(c['Campaign ID'] || '').trim() === id;}) || {};
    return {id:id,name:r['Campaign name'] || r['Campaign Name'] || '',status:r.Status || r['Run status'] || '',country:campaignDetail['Countries/regions'] || config.Country || config['Country / Region'] || source.Country || source.Headquarters || '',industry:campaignDetail['Target industries'] || config.Industry || source.Industry || '',companySize:campaignDetail['Company size'] || config['Company size'] || source['Company size'] || '',companyCount:companies.filter(function(c){return String(c['Campaign ID'] || '').trim() === id;}).length,settings:config};
  }).filter(function(c){return c.id;});
  return JSON.parse(JSON.stringify({campaigns:campaignRows,metrics:metrics,queue:queue,companies:enriched,sourceCompanies:companies,personas:personas,templates:getPersonaMapperTemplates_(),aiApiConfigured:Boolean(PropertiesService.getScriptProperties().getProperty(APP_CONFIG.openAiApiKeyProperty))}));
}

function savePersonaMapperControlPanelSettings(payload) {
  if (!payload || !payload.campaignId) throw new Error('Choose a campaign before saving.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensurePersonaMapperSheets_(ss);
  const campaign = getSheetRecords_(ss.getSheetByName('Campaigns')).find(function(r){return String(r['Campaign ID'] || '').trim() === String(payload.campaignId).trim();});
  if (!campaign) throw new Error('The selected campaign could not be found.');
  const sheet = ss.getSheetByName('Persona Mapper Settings');
  const panelHeaders = ['Department / Functional Responsibility','Persona Seniority','Personas Per Company','Minimum Email Verification Status','Include Inferred Email Addresses','Route Unknown and Risky to Manual Review','Include Campaign-Specific Relevance','Country','Industry','Company size'];
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v || '').trim();});
  const missing = panelHeaders.filter(function(h){return headers.indexOf(h) < 0;});
  if (missing.length) { const start=sheet.getLastColumn()+1; sheet.insertColumnsAfter(sheet.getLastColumn(),missing.length); sheet.getRange(1,start,1,missing.length).setValues([missing]); headers=headers.concat(missing); }
  const idIndex = headers.indexOf('Campaign ID');
  const ids = sheet.getLastRow() > 1 ? sheet.getRange(2,idIndex+1,sheet.getLastRow()-1,1).getDisplayValues().flat() : [];
  let row = ids.findIndex(function(id){return String(id).trim() === String(payload.campaignId).trim();});
  row = row >= 0 ? row + 2 : sheet.getLastRow() + 1;
  const existing = row <= sheet.getLastRow() ? sheet.getRange(row,1,1,headers.length).getValues()[0] : headers.map(function(){return '';});
  const values = existing.slice();
  const campaignName = campaign['Campaign name'] || campaign['Campaign Name'] || '';
  const data = {
    'Campaign ID':String(payload.campaignId).trim(),'Campaign Name':campaignName,'Persona Mapper Enabled':true,
    'Persona Titles':payload.titles || '', 'Department / Functional Responsibility':payload.department || '', 'Persona Seniority':payload.seniority || '',
    'Search Keywords':payload.keywords || '', 'Campaign Focus':payload.focus || '', 'Personas Per Company':Math.max(1,Math.min(10,Number(payload.personasPerCompany) || 3)),
    'Evidence Period Months':Math.max(1,Math.min(60,Number(payload.lookbackMonths) || 12)), 'Research Depth':payload.researchDepth || 'Standard',
    'Include ESG Reporting Obligations':!!payload.criteria.esgReporting, 'Include ESG Reporting Maturity':!!payload.criteria.reportingMaturity,
    'Include Supply Chain Complexity':!!payload.criteria.supplyChain, 'Include Scope 3 Exposure':!!payload.criteria.scope3,
    'Include Sustainability Data Challenges':!!payload.criteria.sustainabilityData, 'Include Digital Transformation':!!payload.criteria.digitalTransformation,
    'Include ESG Governance':!!payload.criteria.governance, 'Include Business Triggers':!!payload.criteria.businessTriggers,
    'Include eValuater Solution Fit':!!payload.criteria.solutionFit, 'Include Campaign-Specific Relevance':!!payload.criteria.campaignRelevance,
    'Minimum Email Verification Status':payload.minimumEmailStatus || 'Verified – Deliverable', 'Include Inferred Email Addresses':!!payload.includeInferred,
    'Route Unknown and Risky to Manual Review':true, 'Email Verification Provider':payload.emailProvider || '', 'Email Verification Endpoint':payload.emailEndpoint || '',
    'Max Companies Per Run':Math.max(1,Math.min(10,Number(payload.batchSize) || 3)), 'Company Selection':payload.companySelection || 'Qualified only', 'Minimum Existing Company Score':Math.max(0,Math.min(100,Number(payload.minimumCompanyScore) || 0)), 'Selected Company IDs':(payload.selectedCompanyIds || []).join(','), 'AI Model':payload.aiModel || '',
    'Minimum Persona Confidence':payload.minimumConfidence || 'Medium','Country':payload.country || '', 'Industry':payload.industry || '', 'Company size':payload.companySize || ''
  };
  headers.forEach(function(h,i){if(Object.prototype.hasOwnProperty.call(data,h)) values[i]=data[h];});
  sheet.getRange(row,1,1,headers.length).setValues([values]);
  return {ok:true,message:'Campaign settings saved.'};
}

function runPersonaMapperFromControlPanel(campaignId) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Campaigns');
  if(!sheet) throw new Error('Campaigns sheet was not found.');
  const records=getSheetRecords_(sheet), index=records.findIndex(function(r){return String(r['Campaign ID'] || '').trim()===String(campaignId || '').trim();});
  if(index<0) throw new Error('Choose an existing campaign.');
  ss.setActiveSheet(sheet); sheet.setActiveRange(sheet.getRange(index+2,1));
  runPersonaMapperForSelectedCampaign();
  return {ok:true,message:'Persona Mapper batch finished or paused. Check the spreadsheet alert and execution log for its status.'};
}

function openPersonaMapperReviewRow(campaignId,personaId) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Prospect Review Queue');
  if(!sheet) throw new Error('Prospect Review Queue was not found.');
  const rows=getSheetRecords_(sheet), index=rows.findIndex(function(r){return String(r['Campaign ID'] || '')===String(campaignId) && String(r['Persona ID'] || '')===String(personaId);});
  if(index<0) throw new Error('That prospect is no longer in the review queue.');
  ss.setActiveSheet(sheet); sheet.setActiveRange(sheet.getRange(index+2,1));
  return {ok:true};
}

function updatePersonaMapperReviewFromControlPanel(campaignId,personaId,action,notes) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Prospect Review Queue'), ui=SpreadsheetApp.getUi();
  if(!sheet) throw new Error('Prospect Review Queue was not found.');
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v || '').trim();});
  const rows=getSheetRecords_(sheet), index=rows.findIndex(function(r){return String(r['Campaign ID'] || '')===String(campaignId) && String(r['Persona ID'] || '')===String(personaId);});
  if(index<0) throw new Error('That prospect is no longer in the review queue.');
  const r=rows[index], row=index+2, now=new Date(), reviewer=currentPersonaMapperUser_();
  if(String(r['Review Status'])==='Approved for Outreach' && action!=='edit') throw new Error('Approved prospects must be reviewed manually in the sheet before changing their status.');
  const set=function(field,value){const col=headers.indexOf(field);if(col>=0)sheet.getRange(row,col+1).setValue(value);};
  if(action==='approve') {
    if(!r['Business Email'] || String(r['Email Verification Status'])!=='Verified – Deliverable') throw new Error('Approval requires a business email marked Verified – Deliverable by the configured provider.');
    if(!['High','Medium'].includes(String(r['Email Identity Confidence']))) throw new Error('Email identity confidence must be High or Medium.');
    if(!r['Draft Subject'] || !r['Draft Email']) throw new Error('A draft subject and message are required.');
    if(String(r['Compliance Status'])!=='Cleared' || String(r['Suppression Check'])!=='Clear') throw new Error('Compliance Status must be Cleared and Suppression Check must be Clear.');
    set('Review Status','Approved for Outreach');set('Human Approval Status','Approved');set('Approved By',reviewer);set('Approved At',now);
  } else if(action==='research') { set('Review Status','Needs Further Research');set('Queue Status','Needs Manual Verification'); }
  else if(action==='reject') { set('Review Status','Rejected');set('Human Approval Status','Rejected');set('Queue Status','Rejected'); }
  else if(action==='manual') { set('Queue Status','Needs Manual Verification');set('Email Verification Status','Manual Verification Required'); }
  else throw new Error('Unsupported review action.');
  set('Reviewer',reviewer);set('Reviewed At',now);set('Reviewer Notes',notes || r['Reviewer Notes'] || '');
  const audit=ss.getSheetByName('Execution Logs');
  if(audit) audit.appendRow(['PM-REVIEW',campaignId,now,now,'Review updated',1,'Prospect review: '+action,String(personaId)]);
  return {ok:true,message:action==='approve'?'Approved for outreach. No email was sent.':'Review decision saved.'};
}


function getPersonaMapperTemplates_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Persona Targeting Templates');
  return getSheetRecords_(sheet).map(function(r){let data={};try{data=JSON.parse(String(r['Template JSON']||'{}'));}catch(e){}return {name:String(r['Template Name']||''),data:data};}).filter(function(r){return r.name;});
}

function savePersonaMapperTemplate(payload) {
  const name=String(payload && payload.templateName || '').trim();
  if(!name) throw new Error('Enter a name for this targeting template.');
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sheet=ss.getSheetByName('Persona Targeting Templates');
  if(!sheet){sheet=ss.insertSheet('Persona Targeting Templates');sheet.getRange(1,1,1,4).setValues([['Template ID','Template Name','Template JSON','Updated At']]);}
  const data={titles:payload.titles||'',department:payload.department||'',seniority:payload.seniority||'',personasPerCompany:payload.personasPerCompany||3,lookbackMonths:payload.lookbackMonths||12,keywords:payload.keywords||'',focus:payload.focus||'',criteria:payload.criteria||{},minimumConfidence:payload.minimumConfidence||'Medium',researchDepth:payload.researchDepth||'Standard',companySelection:payload.companySelection||'Qualified only',minimumCompanyScore:payload.minimumCompanyScore||0};
  const records=getSheetRecords_(sheet),index=records.findIndex(function(r){return String(r['Template Name']||'').trim().toLowerCase()===name.toLowerCase();}),row=index>=0?index+2:sheet.getLastRow()+1;
  sheet.getRange(row,1,1,4).setValues([['PM-TEMPLATE-'+name.toLowerCase().replace(/[^a-z0-9]+/g,'-'),name,JSON.stringify(data),new Date()]]);
  return {ok:true,message:'Targeting template saved.'};
}
