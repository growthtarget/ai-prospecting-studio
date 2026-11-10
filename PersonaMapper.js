const PERSONA_MAPPER_HEADERS_ = {
  companies: ['Campaign ID','Campaign Name','Company ID','Company Name','Country','Industry','Company Size','ESG Frameworks','Reporting Obligations','ESG Reporting Maturity','Supply Chain Profile','Scope 3 Exposure','Sustainability Data Needs','Digital Transformation Signals','ESG Governance','Business Triggers','eValuater Solution Fit','Company Opportunity Score','Company Opportunity','Evidence URLs','Evidence Checked Date','Verification Status','Human Review Status','Last Enriched At','Company Score Rationale'],
  personas: ['Campaign ID','Campaign Name','Company ID','Company Name','Industry','Persona ID','Name','Title','Persona Type','LinkedIn URL','Signal','Signal Date','Evidence URL','Buying Signal','Verification Status','Confidence','Persona Fit Score','Personalized Hook','Review Status','Last Verified At','Business Email','Email Domain','Email Source URL','Email Discovery Method','Email Verification Status','Email Verification Date','Email Verification Provider','Email Identity Confidence','Email Verification Notes'],
  settings: ['Campaign ID','Campaign Name','Persona Mapper Enabled','AI Provider','AI Model','Persona Titles','Search Keywords','Campaign Focus','Evidence Period Months','Research Depth','Include ESG Reporting Obligations','Include ESG Reporting Maturity','Include Supply Chain Complexity','Include Scope 3 Exposure','Include Sustainability Data Challenges','Include Digital Transformation','Include ESG Governance','Include Business Triggers','Include eValuater Solution Fit','Weight ESG Reporting Relevance','Weight Supply Chain Complexity','Weight Sustainability Data Needs','Weight Digital Transformation Relevance','Weight Business Triggers','Weight eValuater Solution Fit','Minimum Company Opportunity Score','Minimum Persona Confidence','Minimum Email Priority Score','Human Approval Required','Auto Send Enabled','Email Priority Weight Company Score','Email Priority Weight Persona Fit','Email Priority Weight Buying Signal','Notes','Max Companies Per Run','Email Verification Provider','Email Verification Endpoint','Email Sender Provider','Email Sender Endpoint'],
  queue: ['Campaign ID','Campaign Name','Company ID','Company Name','Persona ID','Name','Title','Business Email','Company Opportunity Score','Persona Fit Score','Buying Signal','Priority Score','Priority Tier','Priority Score Rationale','Draft Subject','Draft Email','Evidence URLs','Review Status','Reviewer','Reviewer Notes','Reviewed At','Human Approval Status','Approved By','Approved At','Send Status','Sent At','Send Error','Country','Industry','Company Opportunity','Persona Type','LinkedIn URL','Email Domain','Email Source URL','Email Discovery Method','Email Verification Status','Email Verification Date','Email Verification Provider','Email Identity Confidence','Email Verification Notes','Queue Status','Compliance Status','Suppression Check','Date Added','Personalized Hook']
};
const PERSONA_MAPPER_SHEETS_ = { companies:'Company Intelligence', personas:'Persona Leads', settings:'Persona Mapper Settings', queue:'Prospect Review Queue' };

/**
 * Called directly from the Sidebar Control Panel (No blocking popups; returns results to sidebar)
 */
function runPersonaMapperFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    return { success: false, message: 'Persona Mapper is already running. Try again after it finishes.' };
  }
  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const settings = getPersonaMapperSettings_(spreadsheet, campaignId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) {
      return { success: false, message: 'Persona Mapper is disabled for this campaign. Enable it in settings.' };
    }

    const providerChoice = String(settings['AI Provider'] || (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultAIProvider) || 'GEMINI').toUpperCase();
    const defaultModel = providerChoice === 'GEMINI' ? ((typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultGeminiModel) || 'gemini-3.8-flash') : 'gpt-4o';
    let modelChoice = String(settings['AI Model'] || defaultModel).trim();
    if (modelChoice === 'gemini-2.5-flash' || modelChoice === 'gemini-1.5-flash') {
      modelChoice = 'gemini-3.8-flash';
    }

    const companies = getPersonaMapperCompanies_(spreadsheet, campaignId, settings);
    if (!companies.length) {
      return { success: false, message: 'No qualified companies found for campaign ' + campaignId + ' in Company Database. Run Step 1 first!' };
    }

    const batchLimit = Math.max(1, Math.min(10, Number(settings['Max Companies Per Run']) || 3));
    const properties = PropertiesService.getDocumentProperties();
    const cursorKey = 'PERSONA_MAPPER_CURSOR_' + campaignId;
    let start = Number(properties.getProperty(cursorKey) || 0);
    if (start >= companies.length) start = 0;
    const end = Math.min(companies.length, start + batchLimit);
    let processed = 0;

    for (let index = start; index < end; index++) {
      const company = companies[index];
      const aiResponse = callAI({
        provider: providerChoice,
        model: modelChoice,
        prompt: buildPersonaMapperPrompt_(company, settings),
        enableWebSearch: true,
        json: true
      });

      const result = aiResponse.json || {};
      const citations = aiResponse.citations || [];
      const companyScore = scoreCompanyOpportunity_(result.company || {}, settings);
      const companyRecord = makeCompanyIntelligenceRecord_(company, result.company || {}, companyScore, citations, settings);
      
      upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.companies), PERSONA_MAPPER_HEADERS_.companies, 'company', companyRecord);
      const minCompanyScore = Number(settings['Minimum Company Opportunity Score']) || 0;

      (Array.isArray(result.personas) ? result.personas : []).forEach(function(persona) {
        const confidence = titleCaseSignal_(persona.confidence);
        if (confidence === 'Low' || confidenceRank_(confidence) < confidenceRank_(settings['Minimum Persona Confidence'] || 'Medium')) return;
        if (companyScore.score < minCompanyScore) return;
        const fitScore = scorePersonaFit_(persona.fitDimensions || {});
        const emailInfo = prepareEmailIntelligence_(persona, company, settings);
        const personaRecord = makePersonaRecord_(company, persona, fitScore, emailInfo, settings);
        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.personas), PERSONA_MAPPER_HEADERS_.personas, 'persona', personaRecord);
        const queueRecord = makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings);
        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue), PERSONA_MAPPER_HEADERS_.queue, 'queue', queueRecord);
      });

      processed++;
      properties.setProperty(cursorKey, String(index + 1));
    }

    if (end >= companies.length) {
      properties.deleteProperty(cursorKey);
      return { success: true, count: processed, message: 'Persona Mapper completed ' + processed + ' company record(s) for campaign ' + campaignId + '! Leads added to Review Queue.' };
    } else {
      properties.setProperty(cursorKey, String(end));
      return { success: true, count: processed, message: 'Processed ' + processed + ' companies. Click Step 2 again to run next batch.' };
    }
  } catch (err) {
    return { success: false, message: 'Persona Mapper error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Called from the Top Menu (Uses standard dialogs)
 */
function runPersonaMapperForSelectedCampaign() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) { ui.alert('Persona Mapper is already running. Try again after it finishes.'); return; }
  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const campaignId = getPersonaMapperCampaignId_(spreadsheet, ui);
    if (!campaignId) return;
    const settings = getPersonaMapperSettings_(spreadsheet, campaignId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) { ui.alert('Persona Mapper is disabled for this campaign. Set Persona Mapper Enabled to TRUE in Persona Mapper Settings, then try again.'); return; }
    
    const providerChoice = String(settings['AI Provider'] || (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultAIProvider) || 'GEMINI').toUpperCase();
    const defaultModel = providerChoice === 'GEMINI' ? ((typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultGeminiModel) || 'gemini-3.8-flash') : 'gpt-4o';
    let modelChoice = String(settings['AI Model'] || defaultModel).trim();
    if (modelChoice === 'gemini-2.5-flash' || modelChoice === 'gemini-1.5-flash') {
      modelChoice = 'gemini-3.8-flash';
    }

    const companies = getPersonaMapperCompanies_(spreadsheet, campaignId, settings);
    if (!companies.length) { ui.alert('No companies are saved under campaign ' + campaignId + ' in Company Database yet.'); return; }
    const batchLimit = Math.max(1, Math.min(10, Number(settings['Max Companies Per Run']) || 3));
    const properties = PropertiesService.getDocumentProperties();
    const cursorKey = 'PERSONA_MAPPER_CURSOR_' + campaignId;
    let start = Number(properties.getProperty(cursorKey) || 0);
    if (start >= companies.length) start = 0;
    const end = Math.min(companies.length, start + batchLimit);
    let processed = 0;
    
    for (let index = start; index < end; index++) {
      const company = companies[index];
      try {
        const aiResponse = callAI({
          provider: providerChoice,
          model: modelChoice,
          prompt: buildPersonaMapperPrompt_(company, settings),
          enableWebSearch: true,
          json: true
        });

        const result = aiResponse.json || {};
        const citations = aiResponse.citations || [];
        const companyScore = scoreCompanyOpportunity_(result.company || {}, settings);
        const companyRecord = makeCompanyIntelligenceRecord_(company, result.company || {}, companyScore, citations, settings);
        
        upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.companies), PERSONA_MAPPER_HEADERS_.companies, 'company', companyRecord);
        const minCompanyScore = Number(settings['Minimum Company Opportunity Score']) || 0;
        
        (Array.isArray(result.personas) ? result.personas : []).forEach(function(persona) {
          const confidence = titleCaseSignal_(persona.confidence);
          if (confidence === 'Low' || confidenceRank_(confidence) < confidenceRank_(settings['Minimum Persona Confidence'] || 'Medium')) return;
          if (companyScore.score < minCompanyScore) return;
          const fitScore = scorePersonaFit_(persona.fitDimensions || {});
          const emailInfo = prepareEmailIntelligence_(persona, company, settings);
          const personaRecord = makePersonaRecord_(company, persona, fitScore, emailInfo, settings);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.personas), PERSONA_MAPPER_HEADERS_.personas, 'persona', personaRecord);
          const queueRecord = makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue), PERSONA_MAPPER_HEADERS_.queue, 'queue', queueRecord);
        });
        processed++;
        properties.setProperty(cursorKey, String(index + 1));
      } catch (error) {
        console.error('Persona Mapper stopped on company ' + (company['Company Name'] || company['Company name'] || 'unknown') + ': ' + error.message);
        ui.alert('Persona Mapper paused after ' + processed + ' company record(s). The current company was not marked complete. Fix the issue and run again to resume.');
        return;
      }
    }
    if (end >= companies.length) {
      properties.deleteProperty(cursorKey);
      ui.alert('Persona Mapper completed ' + processed + ' company record(s) for campaign ' + campaignId + '. Review Company Intelligence, Persona Leads, and Prospect Review Queue. No email was sent.');
    } else {
      properties.setProperty(cursorKey, String(end));
      ui.alert('Persona Mapper processed ' + processed + ' company record(s). Run it again to continue the next batch. No email was sent.');
    }
  } finally { lock.releaseLock(); }
}

function buildPersonaMapperPrompt_(company,settings) {
  const criteria = [
    ['Include ESG Reporting Obligations','ESG reporting obligations and frameworks'],['Include ESG Reporting Maturity','ESG reporting maturity and disclosures'],
    ['Include Supply Chain Complexity','supply-chain complexity and supplier sustainability'],['Include Scope 3 Exposure','Scope 3 and value-chain data exposure'],
    ['Include Sustainability Data Challenges','sustainability data collection and reporting challenges'],['Include Digital Transformation','AI, analytics, automation, and digital transformation signals'],
    ['Include ESG Governance','ESG governance, ownership, and team structure'],['Include Business Triggers','recent commercial and regulatory triggers'],
    ['Include eValuater Solution Fit','eValuater and EvA solution fit'],['Include Campaign-Specific Relevance','relevance to the existing campaign objective and required criteria']
  ].filter(function(item){return toBoolean_(settings[item[0]]);}).map(function(item){return item[1];});
  return [
    'Act as the eValuater B2B account research, persona mapping, and business email discovery agent.',
    'Research ONLY the single existing company supplied below. Do not discover, suggest, or add other companies.',
    'Use web search and credible public evidence. Prefer official company and regulatory sources. Include direct source URLs for claims.',
    'Separate verified facts from interpretation. If evidence is missing, use an empty string or score 0. Never invent a person, title, source, activity, regulation, email address, or buying intent.',
    'Do not treat commercial fit as confirmed purchase intent. Do not infer inefficient or manual processes without evidence.',
    'Research within the last ' + (Number(settings['Evidence Period Months']) || 12) + ' months for signals unless describing a framework or standing disclosure.',
    'Research depth: ' + (settings['Research Depth'] || 'Standard') + '.',
    'Investigate only these campaign-selected company criteria: ' + (criteria.join('; ') || 'none selected') + '.',
    'Campaign focus: ' + (settings['Campaign Focus'] || ''),
    'Additional research keywords: ' + (settings['Search Keywords'] || ''),
    'Target job titles: ' + (settings['Persona Titles'] || 'Chief Sustainability Officer; Sustainability or ESG Director; Head of ESG Reporting; sustainable procurement or supply-chain sustainability leader; ESG data or digital transformation leader.') + '. Department or functional responsibility: ' + (settings['Department / Functional Responsibility'] || 'not specified') + '. Persona seniority: ' + (settings['Persona Seniority'] || 'not specified') + '.',
    'Company source record: ' + JSON.stringify(company),
    'For each qualified person, search for a publicly available professional business email in this order: official company directory or page, public professional biography or publication, reputable business contact source, then a company pattern only if supported by at least two public employee examples.',
    'Never return personal Gmail, Outlook, Yahoo, or other private email addresses. Never claim deliverability. If an address is inferred from a pattern, label emailDiscoveryMethod as Inferred – Unverified and keep emailIdentityConfidence Low. If no supported candidate is found, leave businessEmail empty and use Not Found.',
    'Return ONLY a JSON object with this shape:',
    '{"company":{"country":"","esgFrameworks":"","reportingObligations":"","reportingMaturity":"","supplyChainProfile":"","scope3Exposure":"","sustainabilityDataNeeds":"","digitalTransformationSignals":"","esgGovernance":"","businessTriggers":"","solutionFit":"","dimensions":{"esgReportingRelevance":0,"supplyChainComplexity":0,"sustainabilityDataNeeds":0,"digitalTransformationRelevance":0,"businessTriggers":0,"solutionFit":0},"evidence":[{"url":"","claim":""}]},"personas":[{"name":"","title":"","personaType":"Decision-maker, Influencer, or Operational Owner","linkedinUrl":"","signal":"","signalDate":"","evidenceUrl":"","buyingSignal":"Active, Relevant, or Emerging","verificationStatus":"Verified or Partially Verified","confidence":"High, Medium, or Low","fitDimensions":{"decisionAuthority":0,"functionalRelevance":0,"opportunityConnection":0,"recentActivity":0,"evidenceQuality":0},"personalizedHook":"","draftSubject":"","draftEmail":"","businessEmail":"","emailDomain":"","emailSourceUrl":"","emailDiscoveryMethod":"Published, Public biography, Business database, Inferred – Unverified, or Not Found","emailIdentityConfidence":"High, Medium, Low, or None"}]}',
    'Each company and persona score dimension must be an integer from 0 to 100. Score only what evidence supports.',
    'Return at most ' + Math.max(1,Math.min(10,Number(settings['Personas Per Company']) || 3)) + ' people. Exclude low-confidence identities. Do not discard a qualified person just because no email is found; return that person with businessEmail empty.',
    'Verify current company and title. Prefer public LinkedIn profiles, company biographies, articles, interviews, and dated announcements.',
    'Active means recent direct activity; Relevant means strong functional responsibility without a direct recent signal; Emerging means relevant role with limited current evidence.',
    'Draft a distinct peer-to-peer demo invitation under 150 words. Mention a specific verified activity or initiative, connect it to the person’s role, introduce EvA as eValuater’s AI-powered sustainability and ESG expert, and invite them to a short demonstration. Do not claim a problem or purchase intent that sources do not establish.',
    'Do not send messages, perform test emails, or claim that an email address is deliverable. Deliverability is checked separately by a configured verification provider.'
  ].join('\n');
}

function parsePersonaMapperResponse_(response) {
  let text = response && typeof response.output_text === 'string' ? response.output_text : '';
  if (!text && response && Array.isArray(response.output)) response.output.forEach(function(item){
    if (item.type !== 'message' || !Array.isArray(item.content)) return;
    item.content.forEach(function(part){if ((part.type === 'output_text' || part.type === 'text') && part.text) text += part.text;});
  });
  if (!text) throw new Error('The AI response did not contain readable output.');
  text = text.trim();
  const fence = String.fromCharCode(96).repeat(3);
  if (text.indexOf(fence) === 0) { const newline = text.indexOf('\n'); text = newline >= 0 ? text.slice(newline + 1) : text.slice(fence.length); text = text.replace(new RegExp('\\n' + fence + '$'),''); }
  return JSON.parse(text);
}

function collectPersonaMapperUrls_(response) {
  const urls = [];
  function walk(value) {
    if (!value) return;
    if (Array.isArray(value)) { value.forEach(walk); return; }
    if (typeof value !== 'object') return;
    if (value.url && /^https?:\/\//i.test(String(value.url))) urls.push(String(value.url));
    Object.keys(value).forEach(function(key){walk(value[key]);});
  }
  walk(response);
  return Array.from(new Set(urls));
}

function scoreCompanyOpportunity_(result,settings) {
  const dimensions = result.dimensions || {};
  const map = [['esgReportingRelevance','Weight ESG Reporting Relevance',20],['supplyChainComplexity','Weight Supply Chain Complexity',20],['sustainabilityDataNeeds','Weight Sustainability Data Needs',20],['digitalTransformationRelevance','Weight Digital Transformation Relevance',15],['businessTriggers','Weight Business Triggers',15],['solutionFit','Weight eValuater Solution Fit',10]];
  let weighted = 0, totalWeight = 0;
  map.forEach(function(item){const weight = configuredNumber_(settings,item[1],item[2]); weighted += clampScore_(dimensions[item[0]]) * weight; totalWeight += weight;});
  const score = totalWeight ? Math.round(weighted / totalWeight) : 0;
  const rationale = map.map(function(item){return item[0] + ': ' + clampScore_(dimensions[item[0]]) + '/100 at ' + configuredNumber_(settings,item[1],item[2]) + '% weight';}).join('; ');
  return {score:score,label:score >= 75 ? 'High' : score >= 50 ? 'Medium' : 'Low',rationale:rationale};
}

function scorePersonaFit_(dimensions) {
  const weights = [['decisionAuthority',25],['functionalRelevance',25],['opportunityConnection',25],['recentActivity',15],['evidenceQuality',10]];
  return Math.round(weights.reduce(function(sum,item){return sum + clampScore_(dimensions[item[0]]) * item[1];},0) / 100);
}

function makeCompanyIntelligenceRecord_(source,enrichment,score,citations,settings) {
  const evidence = Array.isArray(enrichment.evidence) ? enrichment.evidence : [];
  const urls = citations.concat(evidence.map(function(item){return item && item.url ? item.url : ''; })).filter(Boolean);
  return {'Campaign ID':source['Campaign ID'],'Campaign Name':settings['Campaign Name'] || '', 'Company ID':source['Company ID'] || '',
    'Company Name':source['Company name'] || source['Company Name'] || '', 'Country':enrichment.country || source['Headquarters'] || '', 'Industry':source['Industry'] || '',
    'Company Size':source['Company size'] || '', 'ESG Frameworks':enrichment.esgFrameworks || '', 'Reporting Obligations':enrichment.reportingObligations || '',
    'ESG Reporting Maturity':enrichment.reportingMaturity || '', 'Supply Chain Profile':enrichment.supplyChainProfile || '', 'Scope 3 Exposure':enrichment.scope3Exposure || '',
    'Sustainability Data Needs':enrichment.sustainabilityDataNeeds || '', 'Digital Transformation Signals':enrichment.digitalTransformationSignals || '',
    'ESG Governance':enrichment.esgGovernance || '', 'Business Triggers':enrichment.businessTriggers || '', 'eValuater Solution Fit':enrichment.solutionFit || '',
    'Company Opportunity Score':score.score, 'Company Opportunity':score.label, 'Company Score Rationale':score.rationale, 'Evidence URLs':Array.from(new Set(urls)).join('\n'),
    'Evidence Checked Date':new Date(), 'Verification Status':urls.length ? 'Evidence linked' : 'Needs review', 'Human Review Status':'Pending Review', 'Last Enriched At':new Date()};
}

function makePersonaRecord_(company,persona,fitScore,emailInfo,settings) {
  return {'Campaign ID':company['Campaign ID'],'Campaign Name':settings['Campaign Name'] || '', 'Company ID':company['Company ID'] || '',
    'Company Name':company['Company name'] || company['Company Name'] || '', 'Industry':company['Industry'] || '',
    'Persona ID':persona.linkedinUrl || [persona.name,company['Company ID'] || company['Company name'],company['Campaign ID']].join('|'),
    'Name':persona.name || '', 'Title':persona.title || '', 'Persona Type':persona.personaType || '', 'LinkedIn URL':persona.linkedinUrl || '',
    'Signal':persona.signal || '', 'Signal Date':persona.signalDate || '', 'Evidence URL':persona.evidenceUrl || '',
    'Buying Signal':titleCaseSignal_(persona.buyingSignal), 'Verification Status':persona.verificationStatus || '',
    'Confidence':titleCaseSignal_(persona.confidence), 'Persona Fit Score':fitScore, 'Personalized Hook':persona.personalizedHook || '',
    'Review Status':'Pending Review', 'Last Verified At':new Date(), 'Business Email':emailInfo.businessEmail,
    'Email Domain':emailInfo.emailDomain, 'Email Source URL':emailInfo.emailSourceUrl, 'Email Discovery Method':emailInfo.emailDiscoveryMethod,
    'Email Verification Status':emailInfo.emailVerificationStatus, 'Email Verification Date':emailInfo.emailVerificationDate,
    'Email Verification Provider':emailInfo.emailVerificationProvider, 'Email Identity Confidence':emailInfo.emailIdentityConfidence,
    'Email Verification Notes':emailInfo.emailVerificationNotes};
}

function makeEmailQueueRecord_(company,persona,companyScore,fitScore,citations,emailInfo,settings) {
  const signal = titleCaseSignal_(persona.buyingSignal), signalScore = signal === 'Active' ? 100 : signal === 'Relevant' ? 70 : 40;
  const weights = [configuredNumber_(settings,'Email Priority Weight Company Score',50),configuredNumber_(settings,'Email Priority Weight Persona Fit',30),configuredNumber_(settings,'Email Priority Weight Buying Signal',20)];
  const total = weights.reduce(function(sum,value){return sum + value;},0) || 100;
  const priority = Math.round((companyScore.score * weights[0] + fitScore * weights[1] + signalScore * weights[2]) / total);
  const name = persona.name || '', companyName = company['Company name'] || company['Company Name'] || '';
  const urls = citations.concat([persona.evidenceUrl || '',emailInfo.emailSourceUrl || '']).filter(Boolean);
  return {'Campaign ID':company['Campaign ID'],'Campaign Name':settings['Campaign Name'] || '', 'Company ID':company['Company ID'] || '', 'Company Name':companyName,
    'Persona ID':persona.linkedinUrl || [name,company['Company ID'] || companyName,company['Campaign ID']].join('|'), 'Name':name, 'Title':persona.title || '',
    'Business Email':emailInfo.businessEmail, 'Company Opportunity Score':companyScore.score, 'Persona Fit Score':fitScore, 'Buying Signal':signal,
    'Priority Score':priority, 'Priority Tier':priority >= 75 ? 'High' : priority >= 50 ? 'Medium' : 'Low',
    'Priority Score Rationale':'Company fit ' + companyScore.score + ' weighted ' + weights[0] + '%; persona fit ' + fitScore + ' weighted ' + weights[1] + '%; buying signal ' + signal + ' (' + signalScore + ') weighted ' + weights[2] + '%.',
    'Draft Subject':persona.draftSubject || ('eValuater demo for ' + companyName + ' sustainability workflows'), 'Draft Email':persona.draftEmail || persona.personalizedHook || '',
    'Evidence URLs':Array.from(new Set(urls)).join('\n'), 'Review Status':'Pending Review', 'Reviewer':'', 'Reviewer Notes':'', 'Reviewed At':'',
    'Human Approval Status':'Pending', 'Approved By':'', 'Approved At':'', 'Send Status':'Not Used', 'Sent At':'', 'Send Error':'',
    'Country':emailInfo.country || company['Headquarters'] || '', 'Industry':company['Industry'] || '', 'Company Opportunity':companyScore.label,
    'Persona Type':persona.personaType || '', 'LinkedIn URL':persona.linkedinUrl || '', 'Email Domain':emailInfo.emailDomain,
    'Email Source URL':emailInfo.emailSourceUrl, 'Email Discovery Method':emailInfo.emailDiscoveryMethod,
    'Email Verification Status':emailInfo.emailVerificationStatus, 'Email Verification Date':emailInfo.emailVerificationDate,
    'Email Verification Provider':emailInfo.emailVerificationProvider, 'Email Identity Confidence':emailInfo.emailIdentityConfidence,
    'Email Verification Notes':emailInfo.emailVerificationNotes, 'Queue Status':emailInfo.queueStatus,
    'Compliance Status':'Review Required', 'Suppression Check':'Review Required', 'Date Added':new Date(), 'Personalized Hook':persona.personalizedHook || ''};
}

function prepareEmailIntelligence_(persona,company,settings) {
  const candidate = String(persona.businessEmail || '').trim().toLowerCase();
  const method = String(persona.emailDiscoveryMethod || (candidate ? 'Unspecified public source' : 'Not Found')).trim();
  const sourceUrl = String(persona.emailSourceUrl || '').trim();
  const empty = {businessEmail:'',emailDomain:'',emailSourceUrl:sourceUrl,emailDiscoveryMethod:'Not Found',emailVerificationStatus:'Not Found',emailVerificationDate:'',emailVerificationProvider:'',emailIdentityConfidence:'None',emailVerificationNotes:'No suitable professional business email was found.',queueStatus:'Email Research Required',country:persona.country || ''};
  if (!candidate) return empty;
  if (/inferred/i.test(method) && !toBoolean_(settings['Include Inferred Email Addresses'])) return Object.assign({},empty,{emailSourceUrl:sourceUrl,emailVerificationNotes:'Inferred email addresses are excluded by campaign settings; no address was stored.'});
  const at = candidate.lastIndexOf('@');
  const domain = at > 0 ? candidate.slice(at + 1) : '';
  const normalizedMethod = /inferred/i.test(method) ? 'Inferred – Unverified' : method;
  if (!domain || !/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(candidate)) {
    return Object.assign({},empty,{businessEmail:candidate,emailDomain:domain,emailSourceUrl:sourceUrl,emailDiscoveryMethod:normalizedMethod,emailVerificationStatus:'Undeliverable',emailVerificationNotes:'Email syntax is invalid. No deliverability claim was made.',queueStatus:'Email Verification Failed'});
  }
  if (isPersonalEmailDomain_(domain)) {
    return Object.assign({},empty,{businessEmail:'',emailDomain:domain,emailSourceUrl:sourceUrl,emailDiscoveryMethod:'Not Found',emailVerificationStatus:'Not Found',emailVerificationNotes:'Personal email domains are excluded; use a professional company address.',queueStatus:'Email Research Required'});
  }
  const companyDomain = domainFromUrl_(company['Website'] || '');
  const domainMatches = Boolean(companyDomain && (domain === companyDomain || domain.endsWith('.' + companyDomain)));
  const mx = checkEmailDomainMx_(domain);
  let confidence = titleCaseSignal_(persona.emailIdentityConfidence);
  if (!sourceUrl || !domainMatches || /inferred/i.test(normalizedMethod)) confidence = 'Low';
  if (!['High','Medium','Low'].includes(confidence)) confidence = 'Low';
  
  const verifier = callConfiguredEmailVerifier_(candidate, settings);
  let status = /inferred/i.test(normalizedMethod) ? 'Inferred – Unverified' : 'Unknown';
  let queueStatus = 'Email Verification Required';
  let notes = [];
  if (!domainMatches) notes.push('Email domain does not match the company website domain or could not be matched.');
  if (!sourceUrl) notes.push('No direct source URL was provided for the email address.');
  if (mx.checked && !mx.hasMx) notes.push('No MX record was found for the domain; this is a domain check, not a mailbox result.');
  if (!mx.checked) notes.push('Domain mail configuration could not be checked.');
  
  if (!verifier.configured) {
    notes.push('No email verification provider is configured. Syntax and domain checks do not establish mailbox deliverability.');
  } else {
    notes.push(verifier.notes);
    if (/inferred/i.test(normalizedMethod)) {
      status = 'Inferred – Unverified';
      queueStatus = 'Needs Manual Verification';
    } else if (verifier.status === 'undeliverable') {
      status = 'Undeliverable'; queueStatus = 'Email Verification Failed';
    } else if (verifier.status === 'risky' || verifier.catchAll || verifier.disposable) {
      status = 'Risky'; queueStatus = 'Needs Manual Verification';
    } else if (verifier.status === 'deliverable') {
      if (domainMatches && sourceUrl && confidence !== 'Low') { status = 'Verified – Deliverable'; queueStatus = 'Ready for Review'; }
      else { status = 'Unknown'; queueStatus = 'Needs Manual Verification'; notes.push('Provider reported deliverability, but person-to-address identity evidence is insufficient.'); }
    } else {
      status = 'Unknown'; queueStatus = 'Needs Manual Verification';
    }
  }
  return {businessEmail:candidate,emailDomain:domain,emailSourceUrl:sourceUrl,emailDiscoveryMethod:normalizedMethod,emailVerificationStatus:status,
    emailVerificationDate:verifier.configured ? new Date() : '',emailVerificationProvider:verifier.provider || '',emailIdentityConfidence:confidence,
    emailVerificationNotes:notes.filter(Boolean).join(' '),queueStatus:queueStatus,country:persona.country || ''};
}

function callConfiguredEmailVerifier_(email, settings) {
  const provider = String(settings['Email Verification Provider'] || 'ZeroBounce').trim();
  const apiKey = PropertiesService.getScriptProperties().getProperty('EMAIL_VERIFICATION_API_KEY') || '';
  let endpoint = String(settings['Email Verification Endpoint'] || '').trim();

  if (!apiKey && provider !== 'None') {
    return { configured: false, status: 'unknown', provider: provider, notes: 'EMAIL_VERIFICATION_API_KEY is missing in Script Properties.' };
  }

  const p = provider.toLowerCase();

  try {
    if (p.indexOf('zerobounce') >= 0) {
      endpoint = endpoint || 'https://api.zerobounce.net/v2/validate';
      const url = endpoint + '?api_key=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email) + '&ip_address=';
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: 'ZeroBounce', notes: 'ZeroBounce returned status ' + response.getResponseCode() };
      
      const body = JSON.parse(response.getContentText() || '{}');
      const st = String(body.status || '').toLowerCase();
      let status = 'unknown';
      if (st === 'valid') status = 'deliverable';
      if (st === 'invalid') status = 'undeliverable';
      if (st === 'catch-all' || st === 'spamtrap' || st === 'abuse' || st === 'do_not_mail') status = 'risky';

      return { configured: true, status: status, provider: 'ZeroBounce', catchAll: st === 'catch-all', disposable: Boolean(body.sub_status === 'disposable'), notes: 'ZeroBounce result: ' + st };
    }

    if (p.indexOf('millionverifier') >= 0) {
      endpoint = endpoint || 'https://api.millionverifier.com/api/v3/';
      const url = endpoint + '?api=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: 'MillionVerifier', notes: 'MillionVerifier returned status ' + response.getResponseCode() };

      const body = JSON.parse(response.getContentText() || '{}');
      const res = String(body.result || '').toLowerCase();
      let status = 'unknown';
      if (res === 'ok') status = 'deliverable';
      if (res === 'invalid') status = 'undeliverable';
      if (res === 'catch_all' || res === 'disposable') status = 'risky';

      return { configured: true, status: status, provider: 'MillionVerifier', catchAll: res === 'catch_all', disposable: res === 'disposable', notes: 'MillionVerifier result: ' + res };
    }

    if (p.indexOf('neverbounce') >= 0) {
      endpoint = endpoint || 'https://api.neverbounce.com/v4/single/check';
      const url = endpoint + '?key=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: 'NeverBounce', notes: 'NeverBounce returned status ' + response.getResponseCode() };

      const body = JSON.parse(response.getContentText() || '{}');
      const res = String(body.result || '').toLowerCase();
      let status = 'unknown';
      if (res === 'valid') status = 'deliverable';
      if (res === 'invalid') status = 'undeliverable';
      if (res === 'catchall' || res === 'disposable') status = 'risky';

      return { configured: true, status: status, provider: 'NeverBounce', catchAll: res === 'catchall', disposable: res === 'disposable', notes: 'NeverBounce result: ' + res };
    }

    if (endpoint && /^https:\/\//i.test(endpoint)) {
      const response = UrlFetchApp.fetch(endpoint, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey },
        payload: JSON.stringify({ email: email }),
        muteHttpExceptions: true
      });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: provider, notes: 'Custom verifier returned status ' + response.getResponseCode() };

      const body = JSON.parse(response.getContentText() || '{}');
      const rawStatus = String(body.status || body.result || 'unknown').toLowerCase();
      let status = 'unknown';
      if (['deliverable','valid','safe','ok'].indexOf(rawStatus) >= 0) status = 'deliverable';
      if (['risky','catch_all','catch-all','accept_all','disposable'].indexOf(rawStatus) >= 0) status = 'risky';
      if (['undeliverable','invalid'].indexOf(rawStatus) >= 0) status = 'undeliverable';

      return { configured: true, status: status, provider: provider, catchAll: Boolean(body.catch_all), disposable: Boolean(body.disposable), notes: 'Custom result: ' + rawStatus };
    }

    return { configured: false, status: 'unknown', provider: provider, notes: 'Verifier provider not configured.' };
  } catch (err) {
    return { configured: true, status: 'unknown', provider: provider, notes: 'Verification request failed: ' + err.message };
  }
}

function checkEmailDomainMx_(domain) {
  try {
    const url = 'https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=MX';
    const response = UrlFetchApp.fetch(url,{muteHttpExceptions:true});
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return {checked:false,hasMx:false};
    const body = JSON.parse(response.getContentText() || '{}');
    return {checked:true,hasMx:Array.isArray(body.Answer) && body.Answer.some(function(item){return Number(item.type) === 15;})};
  } catch (error) { return {checked:false,hasMx:false}; }
}

function domainFromUrl_(url) {
  const match = String(url || '').toLowerCase().match(/^https?:\/\/(?:www\.)?([^/:?#]+)/);
  return match ? match[1].replace(/^www\./,'') : '';
}

function isPersonalEmailDomain_(domain) {
  return ['gmail.com','googlemail.com','outlook.com','hotmail.com','live.com','yahoo.com','icloud.com','me.com','aol.com','proton.me','protonmail.com'].indexOf(String(domain).toLowerCase()) >= 0;
}

function approveSelectedProspect() {
  const item=getSelectedEmailQueueRow_(); if(!item) return;
  const r=item.record, ui=SpreadsheetApp.getUi();
  const allowed=['Ready for Review','Needs Manual Verification'];
  if(allowed.indexOf(String(r['Queue Status']))<0) { ui.alert('This prospect is not ready for review. Resolve the queue status first.'); return; }
  if(!r['Business Email'] || String(r['Email Verification Status'])!=='Verified – Deliverable') { ui.alert('Approval requires a business email marked Verified – Deliverable.'); return; }
  if(['High','Medium'].indexOf(String(r['Email Identity Confidence']))<0) { ui.alert('Email identity confidence must be High or Medium.'); return; }
  if(!r['Draft Subject'] || !r['Draft Email']) { ui.alert('A reviewed draft subject and message are required.'); return; }
  if(String(r['Compliance Status'])!=='Cleared' || String(r['Suppression Check'])!=='Clear') { ui.alert('Compliance Status must be Cleared and Suppression Check must be Clear.'); return; }
  const answer=ui.alert('Approve for outreach?','This records human approval only. It does not send an email.',ui.ButtonSet.YES_NO); if(answer!==ui.Button.YES) return;
  const now=new Date(), reviewer=currentPersonaMapperUser_();
  setQueueCell_(item,'Review Status','Approved for Outreach'); setQueueCell_(item,'Human Approval Status','Approved');
  setQueueCell_(item,'Reviewer',reviewer); setQueueCell_(item,'Reviewed At',now); setQueueCell_(item,'Approved By',reviewer); setQueueCell_(item,'Approved At',now);
}

function updateSelectedProspectReview_(status,queueStatus) {
  const item=getSelectedEmailQueueRow_(); if(!item) return;
  const ui=SpreadsheetApp.getUi(), r=item.record;
  if(String(r['Review Status'])==='Approved for Outreach') { ui.alert('An approved prospect must be reviewed manually before changing status.'); return; }
  setQueueCell_(item,'Review Status',status); setQueueCell_(item,'Queue Status',queueStatus);
  setQueueCell_(item,'Reviewer',currentPersonaMapperUser_()); setQueueCell_(item,'Reviewed At',new Date());
  if(status==='Rejected') setQueueCell_(item,'Human Approval Status','Rejected');
}

function markSelectedProspectNeedsEditing(){updateSelectedProspectReview_('Needs Editing','Ready for Review');}
function markSelectedProspectNeedsFurtherResearch(){updateSelectedProspectReview_('Needs Further Research','Needs Manual Verification');}
function rejectSelectedProspect(){updateSelectedProspectReview_('Rejected','Rejected');}

function getSelectedEmailQueueRow_() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),ui=SpreadsheetApp.getUi();
  if(sheet.getName()!==PERSONA_MAPPER_SHEETS_.queue){ui.alert('Open Prospect Review Queue and select one prospect row first.');return null;}
  const range=sheet.getActiveRange(); if(!range||range.getNumRows()!==1||range.getRow()<2){ui.alert('Select exactly one prospect row below the header.');return null;}
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const values=sheet.getRange(range.getRow(),1,1,headers.length).getValues()[0],record={};
  headers.forEach(function(h,i){record[h]=values[i];}); return {sheet:sheet,row:range.getRow(),headers:headers,record:record};
}

function setQueueCell_(item,header,value){const i=item.headers.indexOf(header);if(i<0)throw new Error('Prospect Review Queue is missing the '+header+' column.');item.sheet.getRange(item.row,i+1).setValue(value);item.record[header]=value;}

function preparePersonaMapperSheets(){ensurePersonaMapperSheets_(SpreadsheetApp.getActiveSpreadsheet());SpreadsheetApp.getUi().alert('Persona Mapper and Prospect Review Queue tabs are ready.');}

function getPersonaMapperCampaignId_(spreadsheet,ui) {
  const sheet = spreadsheet.getActiveSheet();
  if (['Campaigns','Company Database','Persona Mapper Settings','Campaign Configuration'].indexOf(sheet.getName()) >= 0) {
    const range = sheet.getActiveRange();
    if (range && range.getRow() > 1) { const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0]; const values = sheet.getRange(range.getRow(),1,1,headers.length).getValues()[0]; const record = {}; headers.forEach(function(header,index){record[header] = values[index];}); if (record['Campaign ID']) return String(record['Campaign ID']).trim(); }
  }
  const response = ui.prompt('Persona Mapper','Enter the existing AI Prospecting Studio Campaign ID to enrich:',ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return '';
  return String(response.getResponseText() || '').trim();
}

function getPersonaMapperSettings_(spreadsheet,campaignId) {
  const sheet = spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.settings), records = getSheetRecords_(sheet);
  const match = records.find(function(row){return String(row['Campaign ID'] || '').trim().toLowerCase() === String(campaignId).trim().toLowerCase();});
  if (!match) throw new Error('Add a settings row for campaign ' + campaignId + ' in Persona Mapper Settings.');
  return match;
}

function getPersonaMapperCompanies_(spreadsheet,campaignId,settings) {
  const sheet = spreadsheet.getSheetByName('Company Database'); if (!sheet) throw new Error('Company Database sheet was not found.');
  const minScore = Number(settings && settings['Minimum Existing Company Score']) || 0;
  const qualifiedOnly = !settings || !settings['Company Selection'] || String(settings['Company Selection']) === 'Qualified only';
  return getSheetRecords_(sheet).filter(function(row) {
    if (String(row['Campaign ID'] || '').trim() !== String(campaignId).trim() || !String(row['Company ID'] || row['Company name'] || row['Company Name'] || '').trim()) return false;
    const qualification = String(row['Qualification result'] || row['Qualification Result'] || row['Review status'] || row['Review Status'] || '').trim().toLowerCase();
    if (qualifiedOnly && qualification && !/qualified|approved|pass/.test(qualification)) return false;
    const selectedIds = String(settings && settings['Selected Company IDs'] || '').split(',').map(function(v){return v.trim();}).filter(Boolean);
    const companyKey = String(row['Company ID'] || row['Company name'] || row['Company Name'] || '').trim();
    if (selectedIds.length && selectedIds.indexOf(companyKey) < 0) return false;
    const score = Number(row['Company Opportunity Score'] || row.Score || row.score || 0);
    if (minScore && score < minScore) return false;
    return true;
  });
}

function getSheetRecords_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues(), headers = values[0].map(function(value){return String(value || '').trim();});
  return values.slice(1).filter(function(row){return row.some(function(value){return value !== '' && value !== null;});}).map(function(row){const record = {}; headers.forEach(function(header,index){if(header) record[header] = row[index];}); return record;});
}

function ensurePersonaMapperSheets_(spreadsheet) {
  let legacy=spreadsheet.getSheetByName('Email Review Queue');
  let queue=spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue);
  if(!queue&&legacy){legacy.setName(PERSONA_MAPPER_SHEETS_.queue);queue=legacy;}
  [[PERSONA_MAPPER_SHEETS_.companies,PERSONA_MAPPER_HEADERS_.companies],[PERSONA_MAPPER_SHEETS_.personas,PERSONA_MAPPER_HEADERS_.personas],[PERSONA_MAPPER_SHEETS_.settings,PERSONA_MAPPER_HEADERS_.settings],[PERSONA_MAPPER_SHEETS_.queue,PERSONA_MAPPER_HEADERS_.queue]].forEach(function(d){
    let sh=spreadsheet.getSheetByName(d[0]);if(!sh)sh=spreadsheet.insertSheet(d[0]);
    if(sh.getLastRow()===0||sh.getLastColumn()===0){if(sh.getMaxColumns()<d[1].length)sh.insertColumnsAfter(sh.getMaxColumns(),d[1].length-sh.getMaxColumns());sh.getRange(1,1,1,d[1].length).setValues([d[1]]);return;}
    let headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
    if(d[0]===PERSONA_MAPPER_SHEETS_.queue){headers.forEach(function(h,i){if(h==='Recipient Email')sh.getRange(1,i+1).setValue('Business Email');if(h==='Review Notes')sh.getRange(1,i+1).setValue('Reviewer Notes');});headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});}
    const missing=d[1].filter(function(h){return headers.indexOf(h)<0;});
    if(missing.length){const col=sh.getLastColumn();if(sh.getMaxColumns()<col+missing.length)sh.insertColumnsAfter(sh.getMaxColumns(),col+missing.length-sh.getMaxColumns());sh.getRange(1,col+1,1,missing.length).setValues([missing]);}
  });
}

function upsertPreservingNonEmpty_(sheet,headers,keyType,record) {
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2,1,lastRow - 1,headers.length).getValues() : [];
  function same(row,header,value) {
    const index = headers.indexOf(header);
    if (index < 0) return false;
    return String(row[index] || '').trim().toLowerCase() === String(value || '').trim().toLowerCase();
  }
  function matches(row) {
    if (!same(row,'Campaign ID',record['Campaign ID'])) return false;
    if (keyType === 'company') {
      return record['Company ID'] ? same(row,'Company ID',record['Company ID']) : same(row,'Company Name',record['Company Name']);
    }
    if (keyType === 'persona') {
      return record['LinkedIn URL'] ? same(row,'LinkedIn URL',record['LinkedIn URL']) : same(row,'Name',record['Name']) && same(row,'Company Name',record['Company Name']);
    }
    return record['Persona ID'] ? same(row,'Persona ID',record['Persona ID']) : same(row,'Name',record['Name']) && same(row,'Company Name',record['Company Name']);
  }
  let targetRow = 0;
  for (let i = 0; i < rows.length; i++) { if (matches(rows[i])) { targetRow = i + 2; break; } }
  if (!targetRow) {
    sheet.getRange(sheet.getLastRow() + 1,1,1,headers.length).setValues([headers.map(function(header){return record[header] === undefined ? '' : record[header];})]);
    return;
  }
  const existing = sheet.getRange(targetRow,1,1,headers.length).getValues()[0], updated = existing.slice();
  headers.forEach(function(header,index){const incoming = record[header]; if ((existing[index] === '' || existing[index] === null) && incoming !== undefined && incoming !== '' && incoming !== null) updated[index] = incoming;});
  sheet.getRange(targetRow,1,1,headers.length).setValues([updated]);
}

function configuredNumber_(settings,key,fallback) { const value = settings[key]; if (value === '' || value === null || value === undefined) return fallback; const number = Number(value); return isFinite(number) ? number : fallback; }
function currentPersonaMapperUser_() { return Session.getActiveUser().getEmail() || 'Spreadsheet reviewer'; }
function toBoolean_(value) { if (value === true) return true; return ['true','yes','1','enabled'].indexOf(String(value || '').trim().toLowerCase()) >= 0; }
function clampScore_(value) { const number = Number(value); if (!isFinite(number)) return 0; return Math.max(0,Math.min(100,Math.round(number))); }
function confidenceRank_(value) { const normalized = titleCaseSignal_(value); return normalized === 'High' ? 2 : normalized === 'Medium' ? 1 : 0; }
function titleCaseSignal_(value) { const normalized = String(value || '').trim().toLowerCase(); return normalized ? normalized.charAt(0).toUpperCase() + normalized.slice(1) : ''; }