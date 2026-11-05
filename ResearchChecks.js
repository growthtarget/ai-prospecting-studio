function runOfflineSampleChecks() {
  let checks = 0;
  function assert(condition, message) {
    if (!condition) throw new Error('Offline sample check failed: ' + message);
    checks++;
  }

  const sourceUrl = 'https://example.com/sustainability';
  const sampleOutput = {
    company: {
      name: 'Example Manufacturing',
      website: 'https://example.com',
      headquarters: 'Canada',
      industry: 'Manufacturing',
      company_size: 'Unknown',
      business_summary: 'Publicly documented manufacturing activities.'
    },
    evidence: [
      {
        claim: 'The company publishes sustainability information.',
        extracted_information: 'A sustainability page is available on the company website.',
        source_url: sourceUrl,
        source_type: 'company_website',
        suggested_quality: 'High',
        quality_reason: 'Direct company source.'
      },
      {
        claim: 'The company has a supply-chain program.',
        extracted_information: 'A secondary source describes a supplier program.',
        source_url: 'https://industry.example.org/profile',
        source_type: 'reliable_industry_source',
        suggested_quality: 'Medium',
        quality_reason: 'Secondary source; confirm with the company.'
      }
    ]
  };
  const sampleResponse = {
    status: 'completed',
    output: [{
      type: 'message',
      content: [{
        type: 'output_text',
        text: JSON.stringify(sampleOutput),
        annotations: [{
          type: 'url_citation',
          url_citation: { url: sourceUrl, title: 'Example Sustainability Page' }
        }]
      }]
    }]
  };

  const prepared = prepareCompanyResearchRows_('CAMP-TEST-001', sampleResponse);
  assert(prepared.companyRow.length === 15, 'company row matches the approved column count');
  assert(prepared.evidenceRows.length === 2, 'both sample claims are mapped');
  assert(prepared.evidenceRows[0][8] === 'High', 'a matched citation keeps its suggested quality');
  assert(prepared.evidenceRows[0][7] === 'No', 'research is never marked independently verified');
  assert(prepared.evidenceRows[0][9].indexOf('Example Sustainability Page') >= 0, 'citation title is retained in reviewer notes');
  assert(prepared.evidenceRows[1][8] === 'Unknown', 'a source without matching citation is downgraded to Unknown');
  assert(prepared.companyRow[8] === 'Pending' && prepared.companyRow[10] === '', 'research does not assign qualification or score');
  assert(prepared.companyRow[11] === 'Unknown' && prepared.companyRow[12] === 'Pending human review', 'company-level evidence remains unapproved');

  const newPlan = buildResearchSavePlan_([], prepared);
  assert(newPlan.action === 'save', 'a new company is eligible for saving');
  const websiteDuplicate = buildResearchSavePlan_([
    ['EXISTING-1', 'CAMP-TEST-001', 'Different Name', 'www.example.com/']
  ], prepared);
  assert(websiteDuplicate.action === 'skip_duplicate' && websiteDuplicate.matchReason === 'website', 'website duplicates are skipped');
  const nameDuplicate = buildResearchSavePlan_([
    ['EXISTING-2', 'CAMP-TEST-001', 'Example Manufacturing', 'https://different.example']
  ], prepared);
  assert(nameDuplicate.action === 'skip_duplicate' && nameDuplicate.matchReason === 'company name', 'same-campaign name duplicates are skipped');
  const otherCampaign = buildResearchSavePlan_([
    ['EXISTING-3', 'ANOTHER-CAMPAIGN', 'Different Name', 'https://example.com']
  ], prepared);
  assert(otherCampaign.action === 'save', 'the same company can be researched in another campaign');
  const noEvidencePrepared = {
    companyRow: prepared.companyRow.slice(),
    evidenceRows: []
  };
  const noEvidencePlan = buildResearchSavePlan_([], noEvidencePrepared);
  assert(noEvidencePlan.companyRow[11] === 'Unknown' && noEvidencePlan.companyRow[12] === 'Needs research evidence', 'a record without evidence is flagged for review');

  const result = 'Passed ' + checks + ' offline sample checks. No API calls or spreadsheet writes were made.';
  console.log(result);
  return result;
}