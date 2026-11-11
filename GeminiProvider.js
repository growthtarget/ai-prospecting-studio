/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * GOOGLE GEMINI AI PROVIDER (WITH GOOGLE SEARCH GROUNDING)
 * ============================================================================
 * File: GeminiProvider.gs / GeminiProvider.js
 * 
 * Connects directly to Google Generative Language API (v1beta)
 * with real-time Google Search Grounding and structured JSON validation.
 * ============================================================================
 */

function callGemini_(options) {
  // Multi-Tier API Key Resolution
  let apiKey = PropertiesService.getScriptProperties().getProperty(
    (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.geminiApiKeyProperty) || 'GEMINI_API_KEY'
  );

  if (!apiKey && typeof getStudioApiKey === 'function') {
    apiKey = getStudioApiKey('GEMINI');
  }

  if (!apiKey) {
    apiKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
  }

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing. Add it under Project Settings > Script Properties or Agent Settings.');
  }

  // Model Resolution
  const defaultModel = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultGeminiModel)
    ? APP_CONFIG.defaultGeminiModel
    : 'gemini-3.8-flash';
  
  let model = options.model || defaultModel;
  
  // Auto-upgrade legacy model names if passed from older campaign records
  if (model === 'gemini-2.5-flash' || model === 'gemini-1.5-flash' || model === 'gemini-2.0-flash') {
    model = 'gemini-3.8-flash';
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + apiKey;

  // 1. Build Payload
  const payload = {
    contents: [
      {
        parts: [{ text: options.prompt }]
      }
    ],
    generationConfig: {
      temperature: options.temperature !== undefined ? options.temperature : 0.2
    }
  };

  // 2. Enable Google Search Grounding for live web research
  if (options.enableWebSearch !== false) {
    payload.tools = [{ google_search: {} }];
  }

  // 3. Request Structured JSON if specified
  if (options.json) {
    payload.generationConfig.responseMimeType = 'application/json';
  }

  // 4. Send Server-Side Request via UrlFetchApp
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const rawText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Gemini API Error (' + statusCode + '): ' + rawText);
  }

  const body = JSON.parse(rawText);
  const candidate = body.candidates && body.candidates[0];
  
  if (!candidate || !candidate.content || !candidate.content.parts) {
    throw new Error('Gemini returned an empty response.');
  }

  let textOutput = candidate.content.parts.map(function(p) { return p.text || ''; }).join('').trim();
  
  // Clean markdown code blocks if model wrapped JSON in ```json ... ```
  if (textOutput.indexOf('```') >= 0) {
    textOutput = textOutput.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  }

  // 5. Extract Live Google Search Citations
  const citations = [];
  if (candidate.groundingMetadata) {
    const chunks = candidate.groundingMetadata.groundingChunks || [];
    chunks.forEach(function(chunk) {
      if (chunk.web && chunk.web.uri) {
        citations.push(chunk.web.uri);
      }
    });
    const webSearchQueries = candidate.groundingMetadata.webSearchQueries || [];
    if (webSearchQueries.length > 0) {
      Logger.log('Gemini Search Grounding Queries: ' + webSearchQueries.join(' | '));
    }
  }

  // Parse JSON safely if requested
  let parsedJson = null;
  if (options.json) {
    try {
      parsedJson = JSON.parse(textOutput);
    } catch (e) {
      // Regex recovery if text wrapped JSON
      const jsonMatch = textOutput.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          parsedJson = JSON.parse(jsonMatch[0]);
        } catch (nestedErr) {
          Logger.log('Could not parse Gemini JSON regex match: ' + nestedErr.message);
        }
      }
      if (!parsedJson) {
        console.warn('Gemini response was not valid JSON: ' + textOutput.slice(0, 200));
      }
    }
  }

  return {
    text: textOutput,
    json: parsedJson,
    citations: Array.from(new Set(citations)),
    raw: body
  };
}