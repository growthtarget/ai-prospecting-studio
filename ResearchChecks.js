/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * OFFLINE UNIT TEST SUITE & GROUNDING VERIFICATION HARNESS
 * ============================================================================
 * File: ResearchChecks.js
 * 
 * Executes comprehensive offline unit tests across Gemini Grounding, OpenAI,
 * Tab 8 (Company Database - 10 cols, Col J = Campaign ID), and
 * Tab 9 (Research Evidence - 9 cols) without making live API calls or sheet writes.
 * ============================================================================
 */

function runOfflineSampleChecks() {
  let checks = 0;
  function assert(condition, message) {
    if (!condition) throw new Error('Offline check failed: ' + message);
    checks++;
  }

  const sourceUrl = 'https://example.com/sustainability';
  const ungroundedUrl = 'https://industry.example.org/profile';

  // 1. Mock Structured Payload (as defined in ResearchSchema.js)
  const sampleCompanyOutput = {
    company: {
      name: 'Example Manufacturing Corp',
      website: 'https://example.com',
      headquarters: 'Canada',
      industry: 'Manufacturing',
      company_size: '500-1000',
      business_summary: 'Documented sustainable precision manufacturing.'
    },
    evidence: [
      {
        claim: 'The company publishes annual ESG disclosures.',
        extracted_information: 'Annual sustainability disclosures available on portal.',
        source_url: sourceUrl,
        source_type: 'company_website',
        suggested_quality: 'High',
        quality_reason: 'Direct corporate source.'
      },
      {
        claim: 'The company is expanding North American operations.',
        extracted_information: 'Third party analysis mentions Canadian plant expansion.',
        source_url: ungroundedUrl,
        source_type: 'reliable_industry_source',
        suggested_quality: 'Medium',
        quality_reason: 'Secondary industry article.'
      }
    ]
  };

  // 2. Mock OpenAI Assistants / Responses API Payload
  const openAiSampleResponse = {
    status: 'completed',
    output: [{
      type: 'message',
      content: [{
        type: 'output_text',
        text: JSON.stringify(sampleCompanyOutput),
        annotations: [{
          type: 'url_citation',
          url_citation: { url: sourceUrl, title: 'Example Sustainability Page' }
        }]
      }]
    }]
  };

  // Test OpenAI Preparation
  const preparedOpenAi = prepareCompanyResearchRows_('CAMP-TEST-001', openAiSampleResponse);
  assert(preparedOpenAi.companyRow.length === 10, 'OpenAI company row matches the canonical 10-column schema');
  assert(preparedOpenAi.companyRow[1] === 'Example Manufacturing Corp', 'Company Name is in Column B (index 1)');
  assert(preparedOpenAi.companyRow[2] === 'example.com', 'Domain is extracted into Column C (index 2)');
  assert(preparedOpenAi.companyRow[9] === 'CAMP-TEST-001', 'Campaign ID is correctly mapped to Column J (index 9)');
  assert(preparedOpenAi.evidenceRows.length === 2, 'both evidence items are mapped');
  assert(preparedOpenAi.evidenceRows[0].length === 9, 'evidence row matches the canonical 9-column schema');
  assert(preparedOpenAi.evidenceRows[0][1] === 'CAMP-TEST-001', 'Campaign ID is correctly placed in Column B of evidence');
  assert(preparedOpenAi.evidenceRows[0][3] === 'Company', 'Entity Type is marked as Company in Column D (index 3)');
  assert(preparedOpenAi.evidenceRows[0][7] === 'High', 'matched citation retains suggested quality in Verification Score');
  assert(preparedOpenAi.evidenceRows[1][7] === 'Unknown', 'ungrounded source URL is downgraded to Unknown in Verification Score');

  // 3. Mock Google Gemini API Response with Search Grounding (gemini-3.8-flash)
  const geminiSampleResponse = {
    candidates: [{
      content: {
        parts: [{ text: JSON.stringify(sampleCompanyOutput) }]
      },
      groundingMetadata: {
        webSearchQueries: ['Example Manufacturing ESG disclosures Canada'],
        groundingChunks: [{
          web: {
            uri: sourceUrl,
            title: 'Example Corporate ESG Portal'
          }
        }]
      }
    }]
  };

  // Test Gemini Grounding Preparation
  const preparedGemini = prepareCompanyResearchRows_('CAMP-TEST-001', geminiSampleResponse);
  assert(preparedGemini.companyRow.length === 10, 'Gemini company row matches the canonical 10-column schema');
  assert(preparedGemini.evidenceRows.length === 2, 'Gemini evidence items mapped');
  assert(preparedGemini.evidenceRows[0].length === 9, 'Gemini evidence matches canonical 9-column schema');
  assert(preparedGemini.evidenceRows[0][4].indexOf('ESG disclosures Canada') >= 0, 'Gemini webSearchQuery mapped to Search Query Used');
  assert(preparedGemini.evidenceRows[0][7] === 'High', 'Gemini groundingChunk correctly verifies sourceUrl as High');
  assert(preparedGemini.evidenceRows[1][7] === 'Unknown', 'Gemini ungrounded source is downgraded to Unknown');

  // 4. Test Scenario B Multi-Campaign Deduplication (Tab 8: 10 columns)
  // [ID, CompanyName, Domain, Country, Industry, Size, Method, Date, Status, CampaignID]
  const newPlan = buildResearchSavePlan_([], preparedGemini);
  assert(newPlan.action === 'save', 'new company is eligible for saving');

  const domainDuplicate = buildResearchSavePlan_([
    ['EX-001', 'Different Name Inc', 'example.com', 'Canada', 'Manufacturing', '500', 'AI', '2026-11-12', 'Discovered', 'CAMP-TEST-001']
  ], preparedGemini);
  assert(domainDuplicate.action === 'skip_duplicate' && domainDuplicate.matchReason === 'website domain', 'domain duplicate in same campaign is skipped');

  const nameDuplicate = buildResearchSavePlan_([
    ['EX-002', 'example manufacturing corp', 'other-domain.ca', 'Canada', 'Manufacturing', '500', 'AI', '2026-11-12', 'Discovered', 'CAMP-TEST-001']
  ], preparedGemini);
  assert(nameDuplicate.action === 'skip_duplicate' && nameDuplicate.matchReason === 'company name', 'company name duplicate in same campaign is skipped');

  const otherCampaignTarget = buildResearchSavePlan_([
    ['EX-003', 'Example Manufacturing Corp', 'example.com', 'Canada', 'Manufacturing', '500', 'AI', '2026-11-12', 'Discovered', 'OTHER-CAMPAIGN-002']
  ], preparedGemini);
  assert(otherCampaignTarget.action === 'save', 'Scenario B: same company can be researched under separate campaigns');

  const noEvidencePrepared = {
    companyRow: preparedGemini.companyRow.slice(),
    evidenceRows: []
  };
  const noEvidencePlan = buildResearchSavePlan_([], noEvidencePrepared);
  assert(noEvidencePlan.companyRow[8] === 'Needs Research Evidence', 'record without evidence is flagged in Status column');

  // 5. Test Persona Grounding Evidence Mapping
  const samplePersonaOutput = {
    company_name: 'Example Manufacturing Corp',
    persona: {
      name: 'Jane Doe',
      title: 'VP Sustainability & Compliance',
      persona_type: 'Executive Sponsor',
      linkedin_url: 'https://linkedin.com/in/janedoe',
      buying_signal: 'Mandated to deploy ESG tracking software by Q1.',
      confidence: 'High'
    },
    evidence: [{
      claim: 'Jane Doe oversees sustainability and regulatory reporting.',
      extracted_information: 'Jane Doe was appointed VP Sustainability in 2025.',
      source_url: sourceUrl,
      source_type: 'company_website',
      suggested_quality: 'High',
      quality_reason: 'Corporate executive biography page.'
    }]
  };

  const personaGeminiResponse = {
    candidates: [{
      content: { parts: [{ text: JSON.stringify(samplePersonaOutput) }] },
      groundingMetadata: {
        webSearchQueries: ['Jane Doe VP Sustainability Example Manufacturing'],
        groundingChunks: [{ web: { uri: sourceUrl, title: 'Leadership Team' } }]
      }
    }]
  };

  const preparedPersona = preparePersonaResearchRows_('CAMP-TEST-001', personaGeminiResponse, 'Example Manufacturing Corp');
  assert(preparedPersona.evidenceRows.length === 1, 'persona evidence mapped');
  assert(preparedPersona.evidenceRows[0].length === 9, 'persona evidence matches canonical 9-column schema');
  assert(preparedPersona.evidenceRows[0][1] === 'CAMP-TEST-001', 'persona evidence campaign ID is in Column B');
  assert(preparedPersona.evidenceRows[0][2] === 'Jane Doe', 'persona name mapped to Entity Name in Column C');
  assert(preparedPersona.evidenceRows[0][3] === 'Persona', 'Entity Type is marked Persona in Column D');
  assert(preparedPersona.evidenceRows[0][7] === 'High', 'persona grounding citation verified as High');

  const result = 'Passed ' + checks + ' offline checks across Gemini Grounding, OpenAI Citations, Tab 8 (10-col), and Tab 9 (9-col).';
  console.log(result);
  return result;
}