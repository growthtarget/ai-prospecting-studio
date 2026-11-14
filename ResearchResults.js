/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * RESEARCH RESULTS & GROUNDING PARSER (DUAL AI ROUTING)
 * ============================================================================
 * File: ResearchResults.js
 * 
 * Transforms structured AI responses and live search grounding metadata from
 * Google Gemini (gemini-3.8-flash) and OpenAI (gpt-4o) into canonical rows
 * matching Tab 9 (Company Database - 10 cols) and Tab 10 (Research Evidence - 9 cols).
 * ============================================================================
 */

/**
 * Prepares discovered company firmographics and grounding evidence into canonical spreadsheet rows.
 * 
 * @param {string} campaignId Active Campaign ID
 * @param {Object|string} response Raw API response (or wrapped callAI response) from Gemini or OpenAI
 * @param {Object} [options] Optional configuration (defaultSearchQuery, discoveryMethod)
 * @return {Object} Object containing canonical companyRow (10 cols) and evidenceRows (9 cols each)
 */
function prepareCompanyResearchRows_(campaignId, response, options) {
  if (!campaignId) {
    throw new Error('Campaign ID is required.');
  }
  if (!response) {
    throw new Error('An AI response is required.');
  }

  const opts = options || {};
  const outputText = extractResponseOutputText_(response);
  const cleanedText = cleanJsonString_(outputText);

  let research;
  try {
    research = JSON.parse(cleanedText);
  } catch (error) {
    throw new Error('The research response could not be read as structured data: ' + error.message);
  }

  if (!research.company || !research.company.name) {
    throw new Error('The research response is missing a valid company name.');
  }

  const company = research.company;
  const companyId = Utilities.getUuid();
  const collectedAt = new Date().toISOString();

  // Extract live search grounding citations and queries (supports wrapped & raw Gemini/OpenAI payloads)
  const grounding = extractGroundingMetadata_(response);
  const citationsByUrl = grounding.citationsByUrl;
  const citationCount = Object.keys(citationsByUrl).length;
  const searchQueries = grounding.searchQueries;
  const defaultQuery = opts.defaultSearchQuery || (searchQueries.length > 0 ? searchQueries.join('; ') : 'AI Grounded Search');
  const discoveryMethod = opts.discoveryMethod || 'AI Grounded Discovery';

  // Build Tab 10: Research Evidence rows (Exact 9 canonical columns)
  const evidenceRows = (research.evidence || []).map(function(item) {
    const sourceUrl = String(item.source_url || '').trim();
    const citation = sourceUrl ? citationsByUrl[normalizeSourceUrl_(sourceUrl)] : null;
    const allowedQualities = ['High', 'Medium', 'Low', 'Unknown'];
    const suggestedQuality = allowedQualities.indexOf(item.suggested_quality) >= 0
      ? item.suggested_quality
      : 'Unknown';

    // Anti-Hallucination Guard: If search citations were provided and this URL wasn't in them, downgrade quality
    const quality = (citationCount === 0 || citation) ? suggestedQuality : 'Unknown';
    const snippet = String(item.extracted_information || item.claim || '').trim();

    return [
      Utilities.getUuid(),                 // Col A: Evidence ID
      campaignId,                          // Col B: Campaign ID
      String(company.name || '').trim(),   // Col C: Entity Name
      'Company',                           // Col D: Entity Type
      defaultQuery,                        // Col E: Search Query Used
      sourceUrl,                           // Col F: Grounding Evidence URL
      snippet,                             // Col G: Extracted Snippet / Signal
      quality,                             // Col H: Verification Score
      collectedAt                          // Col I: Date Captured
    ];
  });

  // Extract clean domain for Tab 9: Company Database
  const domain = extractDomainFromUrl_(company.website);

  // Build Tab 9: Company Database row (Exact 10 canonical columns, Col J = Campaign ID)
  const companyRow = [
    companyId,                                          // Col A: Company ID
    String(company.name || '').trim(),                  // Col B: Company Name
    domain,                                             // Col C: Domain
    String(company.headquarters || '').trim(),          // Col D: Country
    String(company.industry || '').trim(),              // Col E: Industry
    String(company.company_size || 'Unknown').trim(),   // Col F: Employee Size
    discoveryMethod,                                    // Col G: Discovery Method
    collectedAt,                                        // Col H: Discovery Date
    'Pending Review',                                   // Col I: Status
    campaignId                                          // Col J: Campaign ID
  ];

  return {
    companyRow: companyRow,
    evidenceRows: evidenceRows,
    citationCount: citationCount,
    searchQueries: searchQueries,
    requiresHumanReview: true
  };
}

/**
 * Prepares persona research and buying signals into Tab 10: Research Evidence rows.
 * 
 * @param {string} campaignId Active Campaign ID
 * @param {Object|string} response Raw or wrapped API response from Gemini or OpenAI
 * @param {string} targetCompanyName Target account name
 * @param {Object} [options] Optional configuration
 * @return {Object} Object containing persona evidenceRows (9 cols each)
 */
function preparePersonaResearchRows_(campaignId, response, targetCompanyName, options) {
  if (!campaignId) throw new Error('Campaign ID is required.');
  if (!response) throw new Error('An AI response is required.');

  const opts = options || {};
  const outputText = extractResponseOutputText_(response);
  const cleanedText = cleanJsonString_(outputText);

  let research;
  try {
    research = JSON.parse(cleanedText);
  } catch (error) {
    throw new Error('The persona research response could not be read as structured data: ' + error.message);
  }

  const collectedAt = new Date().toISOString();
  const grounding = extractGroundingMetadata_(response);
  const citationsByUrl = grounding.citationsByUrl;
  const citationCount = Object.keys(citationsByUrl).length;
  const defaultQuery = opts.defaultSearchQuery || (grounding.searchQueries.length > 0 ? grounding.searchQueries.join('; ') : 'Persona Discovery Search');
  const entityName = research.persona && research.persona.name
    ? String(research.persona.name).trim()
    : (targetCompanyName || 'Persona Lead');

  const evidenceRows = (research.evidence || []).map(function(item) {
    const sourceUrl = String(item.source_url || '').trim();
    const citation = sourceUrl ? citationsByUrl[normalizeSourceUrl_(sourceUrl)] : null;
    const allowedQualities = ['High', 'Medium', 'Low', 'Unknown'];
    const suggestedQuality = allowedQualities.indexOf(item.suggested_quality) >= 0
      ? item.suggested_quality
      : 'Unknown';

    const quality = (citationCount === 0 || citation) ? suggestedQuality : 'Unknown';
    const snippet = String(item.extracted_information || item.claim || '').trim();

    return [
      Utilities.getUuid(),                 // Col A: Evidence ID
      campaignId,                          // Col B: Campaign ID
      entityName,                          // Col C: Entity Name
      'Persona',                           // Col D: Entity Type
      defaultQuery,                        // Col E: Search Query Used
      sourceUrl,                           // Col F: Grounding Evidence URL
      snippet,                             // Col G: Extracted Snippet / Signal
      quality,                             // Col H: Verification Score
      collectedAt                          // Col I: Date Captured
    ];
  });

  return {
    persona: research.persona || null,
    evidenceRows: evidenceRows,
    citationCount: citationCount,
    searchQueries: grounding.searchQueries
  };
}

/**
 * Universal extractor for raw text payload across Gemini, OpenAI, wrapped callAI responses, and strings.
 * 
 * @param {Object|string} response Raw or wrapped API response
 * @return {string} Plain text JSON content
 */
function extractResponseOutputText_(response) {
  if (typeof response === 'string') {
    return response;
  }
  if (!response || typeof response !== 'object') {
    throw new Error('The AI response did not contain structured text or valid message parts.');
  }

  // Unwrap callAI / callGemini_ wrapper if present
  const payload = (response.raw && typeof response.raw === 'object') ? response.raw : response;

  // Google Gemini structure: candidates[0].content.parts[].text
  if (payload.candidates && Array.isArray(payload.candidates) && payload.candidates.length > 0) {
    const candidate = payload.candidates[0];
    if (candidate.content && Array.isArray(candidate.content.parts)) {
      const parts = candidate.content.parts
        .map(function(part) { return part.text || ''; })
        .filter(function(text) { return text.length > 0; });
      if (parts.length > 0) return parts.join('\n');
    }
  }

  // OpenAI Chat Completions structure: choices[0].message.content
  if (payload.choices && Array.isArray(payload.choices) && payload.choices.length > 0) {
    const choice = payload.choices[0];
    if (choice.message && typeof choice.message.content === 'string') {
      return choice.message.content;
    }
  }

  // OpenAI Responses API structure: output[].content[].text
  if (Array.isArray(payload.output)) {
    const textParts = [];
    payload.output.forEach(function(outputItem) {
      if (!outputItem || outputItem.type !== 'message' || !Array.isArray(outputItem.content)) return;
      outputItem.content.forEach(function(contentItem) {
        if (contentItem && (contentItem.type === 'output_text' || contentItem.type === 'text') && typeof contentItem.text === 'string') {
          textParts.push(contentItem.text);
        }
      });
    });
    if (textParts.length > 0) return textParts.join('\n');
  }

  // OpenAI SDK helper or wrapped text fallback
  if (typeof payload.output_text === 'string' && payload.output_text.trim()) {
    return payload.output_text;
  }
  if (typeof response.text === 'string' && response.text.trim()) {
    return response.text;
  }
  if (typeof payload.text === 'string' && payload.text.trim()) {
    return payload.text;
  }
  if (response.json && typeof response.json === 'object') {
    return JSON.stringify(response.json);
  }

  throw new Error('The AI response did not contain structured text or valid message parts.');
}

/**
 * Universal grounding metadata extractor across Gemini Search Grounding,
 * OpenAI Responses annotations, and wrapped callAI citation arrays.
 * 
 * @param {Object} response Raw or wrapped API response
 * @return {Object} Object with citationsByUrl and searchQueries
 */
function extractGroundingMetadata_(response) {
  const citationsByUrl = {};
  const searchQueries = [];

  if (!response || typeof response !== 'object') {
    return { citationsByUrl: citationsByUrl, searchQueries: searchQueries };
  }

  // Unwrap callAI / callGemini_ wrapper if present
  const payload = (response.raw && typeof response.raw === 'object') ? response.raw : response;

  // 1. Google Gemini Grounding Metadata (gemini-3.8-flash)
  if (payload.candidates && Array.isArray(payload.candidates) && payload.candidates.length > 0) {
    const grounding = payload.candidates[0].groundingMetadata;
    if (grounding) {
      if (Array.isArray(grounding.webSearchQueries)) {
        grounding.webSearchQueries.forEach(function(query) {
          if (query && typeof query === 'string') searchQueries.push(query.trim());
        });
      }
      if (Array.isArray(grounding.groundingChunks)) {
        grounding.groundingChunks.forEach(function(chunk) {
          if (chunk.web && chunk.web.uri) {
            const url = chunk.web.uri;
            citationsByUrl[normalizeSourceUrl_(url)] = {
              url: url,
              title: chunk.web.title || ''
            };
          }
        });
      }
    }
  }

  // 2. OpenAI Responses API web_search_call queries & url_citation annotations
  if (Array.isArray(payload.output)) {
    payload.output.forEach(function(outputItem) {
      if (!outputItem) return;
      if (outputItem.type === 'web_search_call' && outputItem.action && outputItem.action.query) {
        searchQueries.push(String(outputItem.action.query).trim());
      }
      if (!Array.isArray(outputItem.content)) return;
      outputItem.content.forEach(function(contentItem) {
        (contentItem.annotations || []).forEach(function(annotation) {
          if (!annotation || annotation.type !== 'url_citation') return;
          const citation = annotation.url_citation || annotation;
          if (!citation.url) return;
          citationsByUrl[normalizeSourceUrl_(citation.url)] = {
            url: citation.url,
            title: citation.title || ''
          };
        });
      });
    });
  }

  // 3. Fallback to wrapped citations array if raw metadata had no URLs
  const fallbackCitations = Array.isArray(response.citations)
    ? response.citations
    : (Array.isArray(payload.citations) ? payload.citations : []);

  if (Object.keys(citationsByUrl).length === 0 && fallbackCitations.length > 0) {
    fallbackCitations.forEach(function(item) {
      const url = typeof item === 'string' ? item : (item && (item.url || item.uri) ? (item.url || item.uri) : '');
      if (url && /^https?:\/\//i.test(url)) {
        citationsByUrl[normalizeSourceUrl_(url)] = {
          url: url,
          title: (item && item.title) ? item.title : ''
        };
      }
    });
  }

  return {
    citationsByUrl: citationsByUrl,
    searchQueries: Array.from(new Set(searchQueries.filter(Boolean)))
  };
}

/**
 * Strips markdown code fences (```json ... ```) from model text output.
 * 
 * @param {string} text Raw text
 * @return {string} Sanitized JSON string
 */
function cleanJsonString_(text) {
  if (!text || typeof text !== 'string') return '';
  let cleaned = text.trim();

  const fenceMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch && fenceMatch[1]) {
    cleaned = fenceMatch[1].trim();
  }

  return cleaned;
}

/**
 * Normalizes source URLs for accurate indexing and comparison.
 * 
 * @param {string} url Source URL
 * @return {string} Normalized URL
 */
function normalizeSourceUrl_(url) {
  return String(url || '')
    .trim()
    .replace(/\/+$/, '');
}

/**
 * Extracts a clean domain name from a website URL.
 * 
 * @param {string} url Web URL
 * @return {string} Clean domain (e.g. example.com)
 */
function extractDomainFromUrl_(url) {
  if (!url) return '';
  return String(url)
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .split(/[/?#]/)[0];
}