function prepareCompanyResearchRows_(campaignId, response) {
  if (!campaignId) {
    throw new Error('Campaign ID is required.');
  }
  if (!response || !Array.isArray(response.output)) {
    throw new Error('A completed Responses API result is required.');
  }
  if (response.status && response.status !== 'completed') {
    throw new Error('The AI response is not complete and cannot be prepared for review.');
  }

  const outputText = extractResponseOutputText_(response);
  let research;
  try {
    research = JSON.parse(outputText);
  } catch (error) {
    throw new Error('The research response could not be read as structured data.');
  }

  if (!research.company || !research.company.name) {
    throw new Error('The research response is missing a company name.');
  }

  const companyId = Utilities.getUuid();
  const collectedAt = new Date().toISOString();
  const citationsByUrl = indexResponseCitations_(response);
  const company = research.company;
  const evidenceRows = (research.evidence || []).map(function(item) {
    const sourceUrl = String(item.source_url || '').trim();
    const citation = sourceUrl ? citationsByUrl[normalizeSourceUrl_(sourceUrl)] : null;
    const allowedQualities = ['High', 'Medium', 'Low', 'Unknown'];
    const suggestedQuality = allowedQualities.indexOf(item.suggested_quality) >= 0
      ? item.suggested_quality
      : 'Unknown';
    const quality = citation ? suggestedQuality : 'Unknown';
    const notes = [];

    if (citation && citation.title) notes.push('Source title: ' + citation.title);
    notes.push('AI-suggested quality: ' + suggestedQuality);
    if (item.quality_reason) notes.push('Reason: ' + item.quality_reason);
    notes.push('Not independently verified.');
    if (!citation) {
      notes.push('No matching citation was found in the API response; check the source manually.');
    }

    const allowedSourceTypes = [
      'company_website',
      'annual_or_sustainability_report',
      'regulatory_or_public_registry',
      'reliable_industry_source',
      'other_public_source',
      'unknown'
    ];
    const sourceType = allowedSourceTypes.indexOf(item.source_type) >= 0
      ? item.source_type
      : 'unknown';

    return [
      Utilities.getUuid(),
      companyId,
      String(item.claim || ''),
      String(item.extracted_information || ''),
      sourceUrl,
      sourceType,
      collectedAt,
      'No',
      quality,
      notes.join(' ')
    ];
  });

  const companyRow = [
    companyId,
    campaignId,
    String(company.name || ''),
    String(company.website || ''),
    String(company.headquarters || ''),
    String(company.industry || ''),
    String(company.company_size || 'Unknown'),
    String(company.business_summary || ''),
    'Pending',
    'Research collected; qualification has not been assessed.',
    '',
    'Unknown',
    'Pending human review',
    collectedAt,
    collectedAt
  ];

  return {
    companyRow: companyRow,
    evidenceRows: evidenceRows,
    citationCount: Object.keys(citationsByUrl).length,
    requiresHumanReview: true
  };
}

function extractResponseOutputText_(response) {
  const textParts = [];
  response.output.forEach(function(outputItem) {
    if (outputItem.type !== 'message' || !Array.isArray(outputItem.content)) return;
    outputItem.content.forEach(function(contentItem) {
      if (contentItem.type === 'output_text' && typeof contentItem.text === 'string') {
        textParts.push(contentItem.text);
      }
    });
  });
  if (textParts.length === 0) {
    throw new Error('The AI response did not contain structured text.');
  }
  return textParts.join('\n');
}

function indexResponseCitations_(response) {
  const citationsByUrl = {};
  response.output.forEach(function(outputItem) {
    if (!Array.isArray(outputItem.content)) return;
    outputItem.content.forEach(function(contentItem) {
      (contentItem.annotations || []).forEach(function(annotation) {
        if (annotation.type !== 'url_citation') return;
        const citation = annotation.url_citation || annotation;
        if (!citation.url) return;
        citationsByUrl[normalizeSourceUrl_(citation.url)] = {
          url: citation.url,
          title: citation.title || ''
        };
      });
    });
  });
  return citationsByUrl;
}

function normalizeSourceUrl_(url) {
  return String(url).trim().replace(/\/$/, '');
}