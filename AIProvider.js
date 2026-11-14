/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * AI PROVIDER UNIVERSAL ROUTER
 * ============================================================================
 * File: AIProvider.gs / AIProvider.js
 * 
 * Routes requests dynamically to Gemini (with Google Search Grounding)
 * or OpenAI (GPT-4o) with structured JSON extraction, citation parsing,
 * and strict ApiTracker budget guardrail enforcement.
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
 * Executes OpenAI requests with dual-mode compatibility (Responses API & guarded Chat Completions)
 */
function executeOpenAiRouting_(options) {
  let rawOpenAi = null;
  const defaultOpenAiModel = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultOpenAiModel)
    ? APP_CONFIG.defaultOpenAiModel
    : 'gpt-4o';
  const model = options.model || defaultOpenAiModel;

  // Mode 1: Attempt via OpenAIProvider.js /v1/responses
  if (typeof requestOpenAIResponse_ === 'function') {
    try {
      const openAiPayload = {
        campaignId: options.campaignId,
        operation: options.operation || 'OpenAIResponse',
        model: model,
        input: options.prompt,
        tools: options.enableWebSearch !== false ? [{ type: 'web_search' }] : []
      };
      if (options.json) {
        openAiPayload.text = { format: { type: 'json_object' } };
      }
      rawOpenAi = requestOpenAIResponse_(openAiPayload);
    } catch (e) {
      const errMsg = String(e && e.message ? e.message : e);
      // Never swallow budget cap aborts or missing API key errors
      if (errMsg.indexOf('RUNAWAY LOOP ABORTED') >= 0 ||
          errMsg.indexOf('not configured') >= 0 ||
          errMsg.indexOf('HTTP 401') >= 0) {
        throw e;
      }
      Logger.log('requestOpenAIResponse_ error, falling back to guarded Chat Completions: ' + errMsg);
    }
  }

  // Mode 2: Guarded fallback to Chat Completions via OpenAIProvider.js
  if (!rawOpenAi) {
    if (typeof requestOpenAIChatCompletion_ === 'function') {
      rawOpenAi = requestOpenAIChatCompletion_({
        campaignId: options.campaignId,
        operation: options.operation || 'OpenAIChatCompletion',
        model: model,
        prompt: options.prompt,
        json: options.json,
        temperature: options.temperature
      });
    } else {
      // Defensive fallback if OpenAIProvider.js helper is unavailable
      if (typeof checkApiBudgetGuardrails_ === 'function') {
        checkApiBudgetGuardrails_(options.campaignId);
      }

      const apiKey = resolveAiProviderOpenAiKey_();
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
  }

  const outputText = extractRawTextFromOpenAi_(rawOpenAi);
  const parsedJson = parsePersonaMapperResponse_(rawOpenAi);
  const citations = collectPersonaMapperUrls_(rawOpenAi);

  return {
    text: outputText || (parsedJson ? JSON.stringify(parsedJson) : ''),
    json: parsedJson,
    citations: citations,
    raw: rawOpenAi
  };
}

/**
 * Resolves valid OPENAI_API_KEY without accepting sheet placeholder strings.
 */
function resolveAiProviderOpenAiKey_() {
  if (typeof resolveValidOpenAiApiKey_ === 'function') {
    return resolveValidOpenAiApiKey_();
  }
  const candidates = [
    PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
    PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY'),
    typeof getStudioApiKey === 'function' ? getStudioApiKey('OPENAI') : ''
  ];
  for (let i = 0; i < candidates.length; i++) {
    const val = String(candidates[i] || '').trim();
    const upper = val.toUpperCase();
    if (val && upper.indexOf('PASTE_') < 0 && upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') < 0) {
      return val;
    }
  }
  return '';
}

/**
 * Extracts and parses JSON safely from OpenAI response structures
 */
function parsePersonaMapperResponse_(raw) {
  if (!raw) return null;

  let text = extractRawTextFromOpenAi_(raw);
  if (!text) return null;

  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch && fenceMatch[1]) {
    text = fenceMatch[1].trim();
  }

  try {
    return JSON.parse(text);
  } catch (err) {
    const firstBrace = text.indexOf('{');
    const lastBrace = text.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      try {
        return JSON.parse(text.slice(firstBrace, lastBrace + 1));
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

  // 1. Standard top-level citation arrays
  if (Array.isArray(raw.citations)) {
    raw.citations.forEach(function(c) {
      if (typeof c === 'string' && /^https?:\/\//i.test(c)) urls.push(c);
      else if (c && c.url) urls.push(c.url);
      else if (c && c.uri) urls.push(c.uri);
    });
  }

  // 2. OpenAI Responses API output[].content[].annotations[] (url_citation)
  if (Array.isArray(raw.output)) {
    raw.output.forEach(function(out) {
      if (!out) return;
      if (Array.isArray(out.citations)) {
        out.citations.forEach(function(c) {
          if (c && c.url) urls.push(c.url);
        });
      }
      if (Array.isArray(out.content)) {
        out.content.forEach(function(contentItem) {
          if (!contentItem || !Array.isArray(contentItem.annotations)) return;
          contentItem.annotations.forEach(function(annotation) {
            if (!annotation || annotation.type !== 'url_citation') return;
            const citation = annotation.url_citation || annotation;
            if (citation && citation.url) urls.push(citation.url);
          });
        });
      }
    });
  }

  // 3. Fallback scan of extracted text for URLs
  if (urls.length === 0) {
    const text = extractRawTextFromOpenAi_(raw);
    const matches = text.match(/https?:\/\/[^\s"'<>)\]]+/g) || [];
    matches.forEach(function(u) {
      urls.push(u.replace(/[.,;]+$/, ''));
    });
  }

  return Array.from(new Set(urls.filter(Boolean)));
}

/**
 * Helper to extract raw text content from Chat Completions or Responses API payloads
 */
function extractRawTextFromOpenAi_(raw) {
  if (!raw) return '';

  // Chat Completions format
  if (raw.choices && raw.choices[0] && raw.choices[0].message && raw.choices[0].message.content) {
    return String(raw.choices[0].message.content).trim();
  }

  // Responses API SDK convenience field
  if (typeof raw.output_text === 'string' && raw.output_text.trim()) {
    return raw.output_text.trim();
  }

  // Raw HTTP Responses API output[] array (skips web_search_call items, reads message content[])
  if (Array.isArray(raw.output)) {
    const textParts = [];
    raw.output.forEach(function(item) {
      if (!item) return;
      if (typeof item.text === 'string' && item.text.trim()) {
        textParts.push(item.text.trim());
        return;
      }
      if (typeof item.content === 'string' && item.content.trim()) {
        textParts.push(item.content.trim());
        return;
      }
      if (item.type === 'message' && Array.isArray(item.content)) {
        item.content.forEach(function(contentItem) {
          if (contentItem && (contentItem.type === 'output_text' || contentItem.type === 'text') && typeof contentItem.text === 'string') {
            textParts.push(contentItem.text.trim());
          }
        });
      }
    });
    if (textParts.length > 0) return textParts.join('\n');
  }

  if (typeof raw === 'string') return raw.trim();

  return '';
}