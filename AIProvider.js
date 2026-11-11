/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AI PROVIDER UNIVERSAL ROUTER
 * ============================================================================
 * File: AIProvider.gs / AIProvider.js
 * 
 * Routes requests dynamically to Gemini (with Google Search Grounding)
 * or OpenAI (GPT-4o) with structured JSON extraction and citation parsing.
 * ============================================================================
 */

/**
 * Universal AI Caller
 * Dispatches requests to Gemini or OpenAI based on configuration or settings.
 */
function callAI(options) {
  if (!options || !options.prompt) {
    throw new Error('callAI requires an options object with a prompt.');
  }

  // Determine provider: passed in options -> or configured default -> fallback to GEMINI
  const defaultProvider = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultAIProvider)
    ? APP_CONFIG.defaultAIProvider
    : 'GEMINI';

  const provider = String(options.provider || defaultProvider).toUpperCase();
  let response;

  if (provider === 'GEMINI') {
    if (typeof callGemini_ !== 'function') {
      throw new Error('callGemini_() is not loaded. Ensure GeminiProvider.gs is present.');
    }
    response = callGemini_(options);
  } else if (provider === 'OPENAI') {
    response = executeOpenAiRouting_(options);
  } else {
    throw new Error('Unsupported AI provider: ' + provider);
  }

  return {
    provider: provider,
    text: response.text || '',
    json: response.json || null,
    citations: response.citations || [],
    raw: response.raw || null
  };
}

/**
 * Executes OpenAI requests with dual-mode compatibility (Responses API & Chat Completions)
 */
function executeOpenAiRouting_(options) {
  let rawOpenAi = null;
  const model = options.model || 'gpt-4o';

  // Mode 1: Attempt via OpenAIProvider.js if requestOpenAIResponse_ exists
  if (typeof requestOpenAIResponse_ === 'function') {
    try {
      const openAiPayload = {
        model: model,
        input: options.prompt,
        tools: options.enableWebSearch !== false ? [{ type: 'web_search' }] : []
      };
      if (options.json) {
        openAiPayload.text = { format: { type: 'json_object' } };
      }
      rawOpenAi = requestOpenAIResponse_(openAiPayload);
    } catch (e) {
      Logger.log('requestOpenAIResponse_ error, falling back to direct Chat Completions: ' + e.message);
    }
  }

  // Mode 2: Direct fallback to standard v1/chat/completions if Mode 1 failed or returned null
  if (!rawOpenAi) {
    let apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey && typeof getStudioApiKey === 'function') {
      apiKey = getStudioApiKey('OPENAI');
    }
    if (!apiKey && typeof APP_CONFIG !== 'undefined' && APP_CONFIG.openAiApiKeyProperty) {
      apiKey = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.openAiApiKeyProperty);
    }
    if (!apiKey) {
      throw new Error('OPENAI_API_KEY is not configured in Project Settings (Script Properties).');
    }

    const payload = {
      model: model,
      messages: [{ role: 'user', content: options.prompt }],
      temperature: options.temperature !== undefined ? options.temperature : 0.2
    };

    if (options.json) {
      payload.response_format = { type: 'json_object' };
    }

    const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const statusCode = res.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) {
      throw new Error('OpenAI HTTP ' + statusCode + ': ' + res.getContentText());
    }

    rawOpenAi = JSON.parse(res.getContentText() || '{}');
  }

  // Parse normalized text, JSON, and citations
  const parsedJson = parsePersonaMapperResponse_(rawOpenAi);
  const citations = collectPersonaMapperUrls_(rawOpenAi);
  const outputText = extractRawTextFromOpenAi_(rawOpenAi);

  return {
    text: outputText || (parsedJson ? JSON.stringify(parsedJson) : ''),
    json: parsedJson,
    citations: citations,
    raw: rawOpenAi
  };
}

/**
 * Extracts and parses JSON safely from OpenAI response structures
 */
function parsePersonaMapperResponse_(raw) {
  if (!raw) return null;

  let text = extractRawTextFromOpenAi_(raw);
  if (!text) return null;

  // Clean markdown code fences if wrapped in ```json ... ```
  if (text.indexOf('```') >= 0) {
    text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  }

  try {
    return JSON.parse(text);
  } catch (err) {
    // If wrapped in additional text, attempt regex capture of first JSON object
    const match = text.match(/\{[\s\S]*\}/);
    if (match) {
      try {
        return JSON.parse(match[0]);
      } catch (nestedErr) {
        Logger.log('Could not parse extracted JSON object: ' + nestedErr.message);
      }
    }
    Logger.log('OpenAI response was not valid JSON: ' + text.slice(0, 200));
    return null;
  }
}

/**
 * Collects citations and web search grounding URLs from OpenAI response payloads
 */
function collectPersonaMapperUrls_(raw) {
  if (!raw) return [];
  const urls = [];

  // 1. Check for standard citation arrays in Responses API
  if (Array.isArray(raw.citations)) {
    raw.citations.forEach(c => {
      if (typeof c === 'string' && /^https?:\/\//i.test(c)) urls.push(c);
      else if (c && c.url) urls.push(c.url);
      else if (c && c.uri) urls.push(c.uri);
    });
  }

  // 2. Check for tool call web search citations
  if (Array.isArray(raw.output)) {
    raw.output.forEach(out => {
      if (out && Array.isArray(out.citations)) {
        out.citations.forEach(c => {
          if (c && c.url) urls.push(c.url);
        });
      }
    });
  }

  // 3. Scan extracted text for URLs if none were structured
  if (urls.length === 0) {
    const text = extractRawTextFromOpenAi_(raw);
    const matches = text.match(/https?:\/\/[^\s"'<>]+/g) || [];
    matches.forEach(u => urls.push(u));
  }

  return Array.from(new Set(urls.filter(Boolean)));
}

/**
 * Helper to extract raw text content from either API format
 */
function extractRawTextFromOpenAi_(raw) {
  if (!raw) return '';

  // Chat Completions format
  if (raw.choices && raw.choices[0] && raw.choices[0].message && raw.choices[0].message.content) {
    return String(raw.choices[0].message.content).trim();
  }

  // Responses API format (output_text or output array)
  if (typeof raw.output_text === 'string' && raw.output_text.trim()) {
    return raw.output_text.trim();
  }

  if (Array.isArray(raw.output) && raw.output[0]) {
    const item = raw.output[0];
    if (typeof item.text === 'string') return item.text.trim();
    if (typeof item.content === 'string') return item.content.trim();
  }

  if (typeof raw === 'string') return raw.trim();

  return '';
}