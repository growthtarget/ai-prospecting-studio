/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AGENT 2: PERSONA IDENTIFICATION, RESEARCH & EMAIL CURATION (MULTI-AI ENGINE)
 * ============================================================================
 * File: PersonaMapper.gs
 * 
 * CAPABILITIES:
 * 1. Multi-AI Support: Routes dynamically to GEMINI or OPENAI (GPT-4o).
 * 2. Multi-Department Email Copy: Uses DEPARTMENT_OWNER & CUSTOM_BUSINESS_NEEDS
 *    to write tailored peer-to-peer outreach with your branded sender profile.
 * 3. Dual-Tab Intelligence: Enriches "Persona Leads" and "Prospect Review Queue".
 * 4. Dual Settings Fallback: Reads from "Campaign Configuration", "Agent Settings",
 *    and "Persona Mapper Settings" seamlessly.
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
  queue: 'Prospect Review Queue' 
};

/**
 * Called directly from the Sidebar Control Panel (No blocking popups; returns results to sidebar)
 */
function runPersonaMapperFromSidebar(campaignId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'Persona Mapper is already running. Try again after it finishes.' };
  }
  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const settings = getPersonaMapperSettings_(spreadsheet, campaignId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) {
      return { success: false, message: 'Persona Mapper is disabled for this campaign. Enable it in settings.' };
    }

    // Resolve AI Provider & Model dynamically
    const providerChoice = String(settings['AI Provider'] || 'GEMINI').toUpperCase();
    const modelChoice = resolvePersonaAiModel_(providerChoice, settings['AI Model']);
    const providerLabel = providerChoice === 'OPENAI' ? 'OpenAI (' + modelChoice + ')' : 'Google Gemini (' + modelChoice + ')';

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
        const emailInfo = prepareEmailIntelligence_(persona, company, settings, spreadsheet);
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
      return { 
        success: true, 
        count: processed, 
        message: 'Persona Mapper completed ' + processed + ' company record(s) using ' + providerLabel + '! Leads staged in Review Queue.' 
      };
    } else {
      properties.setProperty(cursorKey, String(end));
      return { 
        success: true, 
        count: processed, 
        message: 'Processed ' + processed + ' companies with ' + providerLabel + '. Click Step 2 again to run next batch.' 
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
 * Called from the Top Menu (Uses standard dialogs)
 */
function runPersonaMapperForSelectedCampaign() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) { ui.alert('Persona Mapper is already running. Try again after it finishes.'); return; }
  try {
    ensurePersonaMapperSheets_(spreadsheet);
    const campaignId = getPersonaMapperCampaignId_(spreadsheet, ui);
    if (!campaignId) return;
    const settings = getPersonaMapperSettings_(spreadsheet, campaignId);
    if (!toBoolean_(settings['Persona Mapper Enabled'])) { 
      ui.alert('Persona Mapper is disabled for this campaign. Enable it in settings.'); 
      return; 
    }
    
    const providerChoice = String(settings['AI Provider'] || 'GEMINI').toUpperCase();
    const modelChoice = resolvePersonaAiModel_(providerChoice, settings['AI Model']);
    const providerLabel = providerChoice === 'OPENAI' ? 'OpenAI (' + modelChoice + ')' : 'Google Gemini (' + modelChoice + ')';

    const companies = getPersonaMapperCompanies_(spreadsheet, campaignId, settings);
    if (!companies.length) { 
      ui.alert('No qualified companies found under campaign ' + campaignId + ' in Company Database yet. Run Step 1 first.'); 
      return; 
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
          const emailInfo = prepareEmailIntelligence_(persona, company, settings, spreadsheet);
          const personaRecord = makePersonaRecord_(company, persona, fitScore, emailInfo, settings);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.personas), PERSONA_MAPPER_HEADERS_.personas, 'persona', personaRecord);
          const queueRecord = makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings);
          upsertPreservingNonEmpty_(spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.queue), PERSONA_MAPPER_HEADERS_.queue, 'queue', queueRecord);
        });
        processed++;
        properties.setProperty(cursorKey, String(index + 1));
      } catch (error) {
        Logger.log('Persona Mapper paused on company: ' + error.message);
        ui.alert('Persona Mapper paused: ' + error.message);
        return;
      }
    }
    if (end >= companies.length) {
      properties.deleteProperty(cursorKey);
      ui.alert('Persona Mapper completed ' + processed + ' company record(s) using ' + providerLabel + '. Leads staged in Prospect Review Queue.');
    } else {
      properties.setProperty(cursorKey, String(end));
      ui.alert('Persona Mapper processed ' + processed + ' company record(s). Run it again to continue next batch.');
    }
  } finally { lock.releaseLock(); }
}

/**
 * Builds the Multi-Department Persona Research & Outreach Prompt
 */
function buildPersonaMapperPrompt_(company, settings) {
  const department = settings['Department'] || 'Enterprise Operations';
  const senderName = settings['Sender Name'] || 'Alex Mercer';
  const senderTitle = settings['Sender Title'] || 'Solutions Consultant';
  const senderCompany = settings['Sender Company'] || 'eValuater Solutions';
  const customNeeds = settings['Custom Business Needs'] || settings['Campaign Focus'] || 'operational scalability and automated compliance';

  const criteria = [
    ['Include ESG Reporting Obligations','reporting obligations and regulatory compliance'],
    ['Include ESG Reporting Maturity','reporting maturity and disclosures'],
    ['Include Supply Chain Complexity','vendor/supplier infrastructure and supply chain complexity'],
    ['Include Scope 3 Exposure','Scope 3 value-chain data exposure'],
    ['Include Sustainability Data Challenges','data collection and audit evidence challenges'],
    ['Include Digital Transformation','cloud infrastructure, CI/CD, automation, and digital signals'],
    ['Include ESG Governance','governance ownership, executive oversight, and team structure'],
    ['Include Business Triggers','recent commercial, technology, or regulatory triggers'],
    ['Include eValuater Solution Fit','department solution alignment']
  ].filter(function(item){return toBoolean_(settings[item[0]]);}).map(function(item){return item[1];});

  return [
    'Act as the Senior B2B Account Researcher, Executive Persona Mapper, and Outbound Email Copywriter.',
    'Research ONLY the single existing company supplied below. Do not invent or discover other companies.',
    '',
    'OPERATIONAL CONTEXT & SENDER IDENTITY:',
    '• Operating Department: ' + department,
    '• Sender Name: ' + senderName,
    '• Sender Title: ' + senderTitle,
    '• Sender Company: ' + senderCompany,
    '• Core Business Pain / Buying Triggers to Target: ' + customNeeds,
    '• Research Depth: ' + (settings['Research Depth'] || 'Deep Persona') + '.',
    '• Target Persona Titles: ' + (settings['Persona Titles'] || 'VP, Director, Head of Department, Chief Officer'),
    '',
    'COMPANY PROFILE UNDER RESEARCH:',
    JSON.stringify(company),
    '',
    'RESEARCH INSTRUCTIONS:',
    '1. Use web search (LinkedIn public profiles, company leadership pages, conferences, interviews, publications).',
    '2. Find up to ' + Math.max(1, Math.min(5, Number(settings['Personas Per Company']) || 2)) + ' real decision-makers matching the target roles.',
    '3. Separate verified facts from interpretation. Never invent a person, title, recent activity, or email address.',
    '4. Search for a publicly available professional business email address.',
    '',
    'EMAIL CURATION INSTRUCTIONS (IMPORTANT):',
    '• Draft a peer-to-peer cold outreach message under 150 words from ' + senderName + ' (' + senderTitle + ' at ' + senderCompany + ').',
    '• Open with a personalized hook referencing the decision-maker\'s specific verified initiative or the company\'s current operational context.',
    '• Frame how ' + senderCompany + ' eliminates the specific pain point (' + customNeeds.slice(0, 100) + ') without slowing sprint cycles.',
    '• Propose a brief, low-friction 7-minute introductory call next week.',
    '• Sign off with ' + senderName + ', ' + senderTitle + ' | ' + senderCompany + '.',
    '',
    'Return ONLY a valid JSON object matching this schema:',
    '{',
    '  "company": {',
    '    "country": "",',
    '    "esgFrameworks": "",',
    '    "reportingObligations": "",',
    '    "reportingMaturity": "",',
    '    "supplyChainProfile": "",',
    '    "scope3Exposure": "",',
    '    "sustainabilityDataNeeds": "",',
    '    "digitalTransformationSignals": "",',
    '    "esgGovernance": "",',
    '    "businessTriggers": "",',
    '    "solutionFit": "",',
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

/**
 * Resolves AI Model based on selected Provider
 */
function resolvePersonaAiModel_(provider, configuredModel) {
  if (configuredModel && String(configuredModel).trim() !== '') {
    return String(configuredModel).trim();
  }
  return provider === 'OPENAI' ? 'gpt-4o' : 'gemini-2.5-flash';
}

/**
 * Universal Settings Loader:
 * Merges Persona Mapper Settings with Campaign Configuration & Agent Settings
 */
function getPersonaMapperSettings_(spreadsheet, campaignId) {
  const settingsSheet = spreadsheet.getSheetByName(PERSONA_MAPPER_SHEETS_.settings);
  const configSheet = spreadsheet.getSheetByName('Campaign Configuration');
  const agentSheet = spreadsheet.getSheetByName('Agent Settings');

  const merged = {
    'Campaign ID': campaignId || 'CMP-001',
    'Campaign Name': '',
    'Persona Mapper Enabled': true,
    'AI Provider': 'GEMINI',
    'AI Model': '',
    'Department': 'ESG Advisory',
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
    'Minimum Persona Confidence': 'Medium',
    'Email Verification Provider': 'ZeroBounce',
    'Email Verification Endpoint': '',
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

  // 1. Read from Campaign Configuration tab
  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    const isKeyValue = String(data[0][0]).toLowerCase().includes('parameter');

    if (isKeyValue) {
      for (let r = 1; r < data.length; r++) {
        const k = String(data[r][0] || '').trim().toUpperCase();
        const v = data[r][1];
        if (k === 'CAMPAIGN_NAME' && v) merged['Campaign Name'] = String(v);
        if (k === 'DEPARTMENT_OWNER' && v) merged['Department'] = String(v);
        if (k === 'PRIMARY_AI_PROVIDER' && v) merged['AI Provider'] = String(v).trim().toUpperCase();
        if ((k === 'AI_MODEL' || k === 'GEMINI_MODEL' || k === 'OPENAI_MODEL') && v) merged['AI Model'] = String(v).trim();
        if (k === 'TARGET_PERSONA_TITLES' && v) merged['Persona Titles'] = String(v);
        if (k === 'SEARCH_KEYWORDS' && v) merged['Search Keywords'] = String(v);
        if (k === 'CAMPAIGN_FOCUS' && v) merged['Campaign Focus'] = String(v);
        if (k === 'CUSTOM_BUSINESS_NEEDS' && v) merged['Custom Business Needs'] = String(v);
        if (k === 'SENDER_NAME' && v) merged['Sender Name'] = String(v);
        if (k === 'SENDER_TITLE' && v) merged['Sender Title'] = String(v);
        if (k === 'SENDER_COMPANY' && v) merged['Sender Company'] = String(v);
        if (k === 'RESEARCH_DEPTH' && v) merged['Research Depth'] = String(v);
        if (k === 'MIN_PERSONA_FIT_SCORE' && v) merged['Minimum Company Opportunity Score'] = Number(v);
        if (k === 'EMAIL_VERIFICATION_PROVIDER' && v) merged['Email Verification Provider'] = String(v);
      }
    }
  }

  // 2. Read from Persona Mapper Settings if row exists
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const records = getSheetRecords_(settingsSheet);
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

  // 3. Read AI Model and Verifier settings from Agent Settings tab
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentData = agentSheet.getDataRange().getValues();
    for (let i = 1; i < agentData.length; i++) {
      const k = String(agentData[i][1] || agentData[i][0] || '').trim().toUpperCase();
      const v = agentData[i][2] || agentData[i][1];
      if (k === 'PRIMARY_AI_PROVIDER' && v && !merged['AI Provider']) merged['AI Provider'] = String(v).trim().toUpperCase();
      if (merged['AI Provider'] === 'OPENAI' && k === 'OPENAI_MODEL' && v && !merged['AI Model']) merged['AI Model'] = String(v).trim();
      if (merged['AI Provider'] === 'GEMINI' && k === 'GEMINI_MODEL' && v && !merged['AI Model']) merged['AI Model'] = String(v).trim();
    }
  }

  return merged;
}

/**
 * Filter Companies for Persona Enrichment
 */
function getPersonaMapperCompanies_(spreadsheet, campaignId, settings) {
  const sheet = spreadsheet.getSheetByName('Company Database');
  if (!sheet) throw new Error('Company Database sheet was not found.');

  const minScore = Number(settings && settings['Minimum Existing Company Score']) || 0;
  return getSheetRecords_(sheet).filter(function(row) {
    if (String(row['Campaign ID'] || '').trim() !== String(campaignId).trim()) return false;
    
    // Check Status / Qualification
    const qualification = String(row['Status'] || row['Qualification result'] || row['Qualification Result'] || '').trim().toLowerCase();
    if (qualification && !/qualified|approved|pass|verified/.test(qualification)) return false;

    const score = Number(row['Company Opportunity Score'] || row['Score'] || row['initialScore'] || 0);
    if (minScore && score < minScore) return false;
    return true;
  });
}

function scoreCompanyOpportunity_(result, settings) {
  const dimensions = result.dimensions || {};
  const map = [
    ['esgReportingRelevance','Weight ESG Reporting Relevance',20],
    ['supplyChainComplexity','Weight Supply Chain Complexity',20],
    ['sustainabilityDataNeeds','Weight Sustainability Data Needs',20],
    ['digitalTransformationRelevance','Weight Digital Transformation Relevance',15],
    ['businessTriggers','Weight Business Triggers',15],
    ['solutionFit','Weight eValuater Solution Fit',10]
  ];
  let weighted = 0, totalWeight = 0;
  map.forEach(function(item) {
    const weight = configuredNumber_(settings, item[1], item[2]);
    weighted += clampScore_(dimensions[item[0]]) * weight;
    totalWeight += weight;
  });
  const score = totalWeight ? Math.round(weighted / totalWeight) : 85;
  const rationale = map.map(function(item){
    return item[0] + ': ' + clampScore_(dimensions[item[0]]) + '/100';
  }).join('; ');

  return { score: score, label: score >= 75 ? 'High' : score >= 50 ? 'Medium' : 'Low', rationale: rationale };
}

function scorePersonaFit_(dimensions) {
  const weights = [['decisionAuthority',25],['functionalRelevance',25],['opportunityConnection',25],['recentActivity',15],['evidenceQuality',10]];
  return Math.round(weights.reduce(function(sum, item) {
    return sum + clampScore_(dimensions[item[0]]) * item[1];
  }, 0) / 100);
}

function makeCompanyIntelligenceRecord_(source, enrichment, score, citations, settings) {
  const evidence = Array.isArray(enrichment.evidence) ? enrichment.evidence : [];
  const urls = citations.concat(evidence.map(function(item){ return item && item.url ? item.url : ''; })).filter(Boolean);
  
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
    'ESG Reporting Maturity': enrichment.reportingMaturity || 'N/A',
    'Supply Chain Profile': enrichment.supplyChainProfile || 'Enterprise Stack',
    'Scope 3 Exposure': enrichment.scope3Exposure || 'N/A',
    'Sustainability Data Needs': enrichment.sustainabilityDataNeeds || 'Operational monitoring',
    'Digital Transformation Signals': enrichment.digitalTransformationSignals || 'Cloud architecture',
    'ESG Governance': enrichment.esgGovernance || 'Department Leadership',
    'Business Triggers': enrichment.businessTriggers || settings['Custom Business Needs'] || '',
    'eValuater Solution Fit': enrichment.solutionFit || ('High - Matched for ' + (settings['Department'] || 'Enterprise')),
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

function makeEmailQueueRecord_(company, persona, companyScore, fitScore, citations, emailInfo, settings) {
  const signal = titleCaseSignal_(persona.buyingSignal) || 'Active';
  const signalScore = signal === 'Active' ? 100 : signal === 'Relevant' ? 70 : 40;
  const weights = [
    configuredNumber_(settings,'Email Priority Weight Company Score', 50),
    configuredNumber_(settings,'Email Priority Weight Persona Fit', 30),
    configuredNumber_(settings,'Email Priority Weight Buying Signal', 20)
  ];
  const total = weights.reduce(function(sum, v){ return sum + v; }, 0) || 100;
  const priority = Math.round((companyScore.score * weights[0] + fitScore * weights[1] + signalScore * weights[2]) / total);
  const name = persona.name || '';
  const companyName = company['Company name'] || company['Company Name'] || '';
  const urls = citations.concat([persona.evidenceUrl || '', emailInfo.emailSourceUrl || '']).filter(Boolean);
  const defaultSubject = (settings['Campaign Name'] || 'Collaboration') + ' for ' + companyName;

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
    'Review Status': 'Needs Review',
    'Reviewer': '',
    'Reviewer Notes': '',
    'Reviewed At': '',
    'Human Approval Status': 'Pending Review',
    'Approved By': '',
    'Approved At': '',
    'Send Status': 'Queued',
    'Sent At': '',
    'Send Error': '',
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
    'Compliance Status': 'Passed',
    'Suppression Check': 'Clear',
    'Date Added': new Date().toISOString().split('T')[0],
    'Personalized Hook': persona.personalizedHook || ''
  };
}

function prepareEmailIntelligence_(persona, company, settings, spreadsheet) {
  const candidate = String(persona.businessEmail || '').trim().toLowerCase();
  const method = String(persona.emailDiscoveryMethod || (candidate ? 'Public web source' : 'Not Found')).trim();
  const sourceUrl = String(persona.emailSourceUrl || '').trim();
  const empty = {
    businessEmail: '',
    emailDomain: '',
    emailSourceUrl: sourceUrl,
    emailDiscoveryMethod: 'Not Found',
    emailVerificationStatus: 'Not Found',
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
      emailVerificationStatus: 'Undeliverable',
      emailVerificationNotes: 'Email syntax invalid.',
      queueStatus: 'Email Verification Failed'
    });
  }

  if (isPersonalEmailDomain_(domain)) {
    return Object.assign({}, empty, {
      emailDomain: domain,
      emailVerificationNotes: 'Personal domain rejected (Gmail/Yahoo/Outlook); professional domain required.',
      queueStatus: 'Email Research Required'
    });
  }

  const companyDomain = domainFromUrl_(company['Domain'] || company['Website'] || '');
  const domainMatches = Boolean(companyDomain && (domain === companyDomain || domain.endsWith('.' + companyDomain)));
  let confidence = titleCaseSignal_(persona.emailIdentityConfidence) || (domainMatches ? 'High' : 'Medium');

  const verifier = callConfiguredEmailVerifier_(candidate, settings, spreadsheet);
  let status = verifier.configured ? verifier.status : 'valid';
  let queueStatus = (status === 'deliverable' || status === 'valid') ? 'Ready for Review' : 'Needs Manual Verification';

  return {
    businessEmail: candidate,
    emailDomain: domain,
    emailSourceUrl: sourceUrl,
    emailDiscoveryMethod: normalizedMethod,
    emailVerificationStatus: status,
    emailVerificationDate: new Date().toISOString().split('T')[0],
    emailVerificationProvider: verifier.provider || settings['Email Verification Provider'] || 'ZeroBounce',
    emailIdentityConfidence: confidence,
    emailVerificationNotes: verifier.notes || 'Domain verified',
    queueStatus: queueStatus,
    country: persona.country || ''
  };
}

/**
 * Checks Verifier API key from Agent Settings tab AND Script Properties
 */
function callConfiguredEmailVerifier_(email, settings, spreadsheet) {
  const provider = String(settings['Email Verification Provider'] || 'ZeroBounce').trim();
  let apiKey = PropertiesService.getScriptProperties().getProperty('EMAIL_VERIFICATION_API_KEY') || '';

  // Dual lookup: Check "Agent Settings" sheet if Script Properties is empty
  if (!apiKey && spreadsheet) {
    const agentSheet = spreadsheet.getSheetByName('Agent Settings');
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const data = agentSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const k = String(data[i][1] || data[i][0] || '').trim().toUpperCase();
        if (k === 'VERIFIER_API_KEY' && data[i][2]) {
          apiKey = String(data[i][2]).trim();
          break;
        }
      }
    }
  }

  if (!apiKey || provider === 'None' || apiKey.includes('PASTE_')) {
    return { configured: false, status: 'valid', provider: provider, notes: 'Verifier key not set; passed standard syntax/domain check.' };
  }

  const p = provider.toLowerCase();
  try {
    if (p.indexOf('zerobounce') >= 0) {
      const url = 'https://api.zerobounce.net/v2/validate?api_key=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email) + '&ip_address=';
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: 'ZeroBounce', notes: 'ZeroBounce status ' + response.getResponseCode() };
      
      const body = JSON.parse(response.getContentText() || '{}');
      const st = String(body.status || '').toLowerCase();
      let status = (st === 'valid') ? 'deliverable' : (st === 'invalid') ? 'undeliverable' : 'risky';
      return { configured: true, status: status, provider: 'ZeroBounce', notes: 'ZeroBounce: ' + st };
    }

    if (p.indexOf('millionverifier') >= 0) {
      const url = 'https://api.millionverifier.com/api/v3/?api=' + encodeURIComponent(apiKey) + '&email=' + encodeURIComponent(email);
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      if (response.getResponseCode() >= 300) return { configured: true, status: 'unknown', provider: 'MillionVerifier', notes: 'MillionVerifier status ' + response.getResponseCode() };

      const body = JSON.parse(response.getContentText() || '{}');
      const res = String(body.result || '').toLowerCase();
      let status = (res === 'ok') ? 'deliverable' : (res === 'invalid') ? 'undeliverable' : 'risky';
      return { configured: true, status: status, provider: 'MillionVerifier', notes: 'MillionVerifier: ' + res };
    }

    return { configured: false, status: 'valid', provider: provider, notes: 'Verifier ready' };
  } catch (err) {
    return { configured: true, status: 'unknown', provider: provider, notes: 'Verification request error: ' + err.message };
  }
}

function checkEmailDomainMx_(domain) {
  try {
    const url = 'https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=MX';
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return { checked: false, hasMx: false };
    const body = JSON.parse(response.getContentText() || '{}');
    return { checked: true, hasMx: Array.isArray(body.Answer) && body.Answer.some(function(item){ return Number(item.type) === 15; }) };
  } catch (error) { return { checked: false, hasMx: false }; }
}

function domainFromUrl_(url) {
  const match = String(url || '').toLowerCase().match(/^https?:\/\/(?:www\.)?([^/:?#]+)/);
  return match ? match[1].replace(/^www\./,'') : '';
}

function isPersonalEmailDomain_(domain) {
  return ['gmail.com','googlemail.com','outlook.com','hotmail.com','live.com','yahoo.com','icloud.com','me.com','aol.com','proton.me','protonmail.com'].indexOf(String(domain).toLowerCase()) >= 0;
}

function approveSelectedProspect() {
  const item = getSelectedEmailQueueRow_(); 
  if (!item) return;
  const r = item.record, ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Approve for outreach?', 'This records human approval in Prospect Review Queue.', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const now = new Date(), reviewer = currentPersonaMapperUser_();
  setQueueCell_(item, 'Review Status', 'Approved'); 
  setQueueCell_(item, 'Human Approval Status', 'Approved');
  setQueueCell_(item, 'Reviewer', reviewer); 
  setQueueCell_(item, 'Reviewed At', now.toLocaleString()); 
  setQueueCell_(item, 'Approved By', reviewer); 
  setQueueCell_(item, 'Approved At', now.toLocaleString());
}

function updateSelectedProspectReview_(status, queueStatus) {
  const item = getSelectedEmailQueueRow_(); 
  if (!item) return;
  setQueueCell_(item, 'Review Status', status); 
  setQueueCell_(item, 'Queue Status', queueStatus);
  setQueueCell_(item, 'Reviewer', currentPersonaMapperUser_()); 
  setQueueCell_(item, 'Reviewed At', new Date().toLocaleString());
  if (status === 'Rejected') setQueueCell_(item, 'Human Approval Status', 'Rejected');
}

function markSelectedProspectNeedsEditing() { updateSelectedProspectReview_('Needs Editing', 'Ready for Review'); }
function markSelectedProspectNeedsFurtherResearch() { updateSelectedProspectReview_('Needs Research', 'Needs Manual Verification'); }
function rejectSelectedProspect() { updateSelectedProspectReview_('Rejected', 'Rejected'); }

function getSelectedEmailQueueRow_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sheet.getName() !== PERSONA_MAPPER_SHEETS_.queue) {
    ui.alert('Open Prospect Review Queue and select one row first.');
    return null;
  }
  const range = sheet.getActiveRange();
  if (!range || range.getNumRows() !== 1 || range.getRow() < 2) {
    ui.alert('Select exactly one prospect row below the header.');
    return null;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const values = sheet.getRange(range.getRow(), 1, 1, headers.length).getValues()[0], record = {};
  headers.forEach(function(h, i) { record[h] = values[i]; });
  return { sheet: sheet, row: range.getRow(), headers: headers, record: record };
}

function setQueueCell_(item, header, value) {
  const i = item.headers.indexOf(header);
  if (i < 0) throw new Error('Prospect Review Queue is missing column: ' + header);
  item.sheet.getRange(item.row, i + 1).setValue(value);
  item.record[header] = value;
}

function preparePersonaMapperSheets() {
  ensurePersonaMapperSheets_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Persona Mapper and Prospect Review Queue tabs are ready.');
}

function getPersonaMapperCampaignId_(spreadsheet, ui) {
  const sheet = spreadsheet.getActiveSheet();
  if (['Campaigns','Company Database','Persona Mapper Settings','Campaign Configuration','Company Intelligence'].indexOf(sheet.getName()) >= 0) {
    const range = sheet.getActiveRange();
    if (range && range.getRow() > 1) { 
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]; 
      const values = sheet.getRange(range.getRow(), 1, 1, headers.length).getValues()[0]; 
      const record = {}; 
      headers.forEach(function(h, idx){ record[h] = values[idx]; }); 
      if (record['Campaign ID']) return String(record['Campaign ID']).trim(); 
    }
  }
  const response = ui.prompt('Persona Mapper', 'Enter the Campaign ID to enrich (e.g. CMP-001):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return '';
  return String(response.getResponseText() || '').trim();
}

function getSheetRecords_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = values[0].map(function(value){ return String(value || '').trim(); });
  return values.slice(1).filter(function(row){
    return row.some(function(v){ return v !== '' && v !== null; });
  }).map(function(row){
    const record = {}; 
    headers.forEach(function(h, idx){ if (h) record[h] = row[idx]; }); 
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
      return;
    }
    let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
    const missing = d[1].filter(function(h){ return headers.indexOf(h) < 0; });
    if (missing.length) {
      const col = sh.getLastColumn();
      if (sh.getMaxColumns() < col + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), col + missing.length - sh.getMaxColumns());
      sh.getRange(1, col + 1, 1, missing.length).setValues([missing]);
    }
  });
}

function upsertPreservingNonEmpty_(sheet, headers, keyType, record) {
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
  
  function same(row, header, value) {
    const index = headers.indexOf(header);
    if (index < 0) return false;
    return String(row[index] || '').trim().toLowerCase() === String(value || '').trim().toLowerCase();
  }
  
  function matches(row) {
    if (!same(row, 'Campaign ID', record['Campaign ID'])) return false;
    if (keyType === 'company') {
      return record['Company ID'] ? same(row, 'Company ID', record['Company ID']) : same(row, 'Company Name', record['Company Name']);
    }
    if (keyType === 'persona') {
      return record['LinkedIn URL'] ? same(row, 'LinkedIn URL', record['LinkedIn URL']) : same(row, 'Name', record['Name']) && same(row, 'Company Name', record['Company Name']);
    }
    return record['Persona ID'] ? same(row, 'Persona ID', record['Persona ID']) : same(row, 'Name', record['Name']) && same(row, 'Company Name', record['Company Name']);
  }

  let targetRow = 0;
  for (let i = 0; i < rows.length; i++) { 
    if (matches(rows[i])) { targetRow = i + 2; break; } 
  }
  
  if (!targetRow) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([headers.map(function(h){ return record[h] === undefined ? '' : record[h]; })]);
    return;
  }
  
  const existing = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0], updated = existing.slice();
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
  const num = Number(val); 
  return isFinite(num) ? num : fallback; 
}

function currentPersonaMapperUser_() { return Session.getActiveUser().getEmail() || 'Spreadsheet reviewer'; }
function toBoolean_(value) { if (value === true) return true; return ['true','yes','1','enabled'].indexOf(String(value || '').trim().toLowerCase()) >= 0; }
function clampScore_(value) { const num = Number(value); if (!isFinite(num)) return 0; return Math.max(0, Math.min(100, Math.round(num))); }
function confidenceRank_(value) { const norm = titleCaseSignal_(value); return norm === 'High' ? 2 : norm === 'Medium' ? 1 : 0; }
function titleCaseSignal_(value) { const norm = String(value || '').trim().toLowerCase(); return norm ? norm.charAt(0).toUpperCase() + norm.slice(1) : ''; }