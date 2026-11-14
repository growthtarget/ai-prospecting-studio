/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * GOOGLE GEMINI AI PROVIDER (WITH GOOGLE SEARCH GROUNDING & USAGE TRACKING)
 * ============================================================================
 * File: GeminiProvider.gs / GeminiProvider.js
 * 
 * Connects directly to Google Generative Language API (v1beta)
 * with real-time Google Search Grounding, structured JSON validation,
 * and automated API usage and cost tracking.
 * ============================================================================
 */

function callGemini_(options) {
  // 0. Pre-Flight Budget Hard Cap Check
  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(options.campaignId);
  }

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

  // 4. Send Server-Side Request via UrlFetchApp with Usage Logging & Error Catching
  let response, statusCode, rawText;
  try {
    response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    statusCode = response.getResponseCode();
    rawText = response.getContentText();
  } catch (netErr) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: options.campaignId,
        serviceName: 'GEMINI',
        operation: options.operation || 'CompanyDiscovery',
        httpStatus: 500,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: netErr.toString(),
        modelName: model
      });
    }
    throw new Error('Gemini Network Error: ' + netErr.toString());
  }

  if (statusCode < 200 || statusCode >= 300) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: options.campaignId,
        serviceName: 'GEMINI',
        operation: options.operation || 'CompanyDiscovery',
        httpStatus: statusCode,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: rawText.slice(0, 300),
        modelName: model
      });
    }
    throw new Error('Gemini API Error (' + statusCode + '): ' + rawText);
  }

  const body = JSON.parse(rawText);

  // Extract Token Usage Metadata
  const usage = body.usageMetadata || {};
  const promptTokens = usage.promptTokenCount || 0;
  const candidateTokens = usage.candidatesTokenCount || 0;

  // Log Successful API Call
  if (typeof logApiUsage_ === 'function') {
    logApiUsage_({
      campaignId: options.campaignId,
      serviceName: 'GEMINI',
      operation: options.operation || 'CompanyDiscovery',
      httpStatus: statusCode,
      inputUnits: promptTokens,
      outputUnits: candidateTokens,
      status: 'SUCCESS',
      details: 'Prompt tokens: ' + promptTokens + ', Candidate tokens: ' + candidateTokens,
      modelName: model
    });
  }

  const candidate = body.candidates && body.candidates[0];
  
  if (!candidate || !candidate.content || !candidate.content.parts) {
    throw new Error('Gemini returned an empty response.');
  }

  let textOutput = candidate.content.parts.map(function(p) { return p.text || ''; }).join('').trim();
  
  if (textOutput.indexOf('```') >= 0) {
    textOutput = textOutput.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  }

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

  let parsedJson = null;
  if (options.json) {
    try {
      parsedJson = JSON.parse(textOutput);
    } catch (e) {
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