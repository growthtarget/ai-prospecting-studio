/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * OPENAI AI PROVIDER (WITH USAGE TRACKING & BUDGET GUARDRAILS)
 * ============================================================================
 * File: OpenAIProvider.gs / OpenAIProvider.js
 * 
 * Handles LLM calls to OpenAI Responses API & Chat Completions with
 * pre-flight budget hard-cap checks, placeholder-safe API key resolution,
 * accurate token extraction, and API usage logging.
 * ============================================================================
 */

function requestOpenAIResponse_(request) {
  if (!request || !request.model) {
    throw new Error('Choose a model in the campaign or agent settings before sending an OpenAI request.');
  }
  if (request.input === undefined || request.input === null || request.input === '') {
    throw new Error('AI request input is required.');
  }

  // 0. Pre-Flight Budget Hard Cap Check
  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(request.campaignId);
  }

  const apiKey = resolveValidOpenAiApiKey_();
  if (!apiKey) {
    throw new Error('The OpenAI API key is not configured in Apps Script Script Properties.');
  }

  const responsesUrl = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.openAiResponsesUrl)
    ? APP_CONFIG.openAiResponsesUrl
    : 'https://api.openai.com/v1/responses';

  try {
    return executeOpenAIResponsesFetch_(responsesUrl, apiKey, request, true);
  } catch (err) {
    if (err && err.isWebSearchFallbackEligible && Array.isArray(request.tools) && request.tools.length > 0) {
      Logger.log('OpenAI web_search tool unavailable for current tier/model (' + err.message + '). Retrying without web_search...');
      if (typeof checkApiBudgetGuardrails_ === 'function') {
        checkApiBudgetGuardrails_(request.campaignId);
      }
      return executeOpenAIResponsesFetch_(responsesUrl, apiKey, request, false);
    }
    throw err;
  }
}

/**
 * Executes a single fetch to OpenAI /v1/responses and logs usage accurately.
 */
function executeOpenAIResponsesFetch_(responsesUrl, apiKey, request, includeTools) {
  const payload = {
    model: request.model,
    input: request.input
  };
  if (request.instructions) payload.instructions = request.instructions;
  if (includeTools && Array.isArray(request.tools) && request.tools.length > 0) {
    payload.tools = request.tools;
  }
  if (request.text) payload.text = request.text;

  let response, status, rawText;
  try {
    response = UrlFetchApp.fetch(responsesUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    status = response.getResponseCode();
    rawText = response.getContentText();
  } catch (netErr) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: request.campaignId,
        serviceName: 'OPENAI',
        operation: request.operation || 'OpenAIResponse',
        httpStatus: 500,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: netErr.toString(),
        modelName: request.model
      });
    }
    throw new Error('OpenAI Network Error: ' + netErr.toString());
  }

  if (status < 200 || status >= 300) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: request.campaignId,
        serviceName: 'OPENAI',
        operation: request.operation || 'OpenAIResponse',
        httpStatus: status,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: rawText.slice(0, 300),
        modelName: request.model
      });
    }

    const err = new Error(
      'OpenAI request failed with HTTP ' + status + ': ' + rawText.slice(0, 250)
    );
    const lowerErr = String(rawText || '').toLowerCase();
    if (includeTools && (status === 400 || status === 403) &&
        (lowerErr.indexOf('web_search') >= 0 || lowerErr.indexOf('tool') >= 0)) {
      err.isWebSearchFallbackEligible = true;
    }
    throw err;
  }

  const parsed = JSON.parse(rawText || '{}');

  // OpenAI /v1/responses returns input_tokens & output_tokens (fallback to prompt/completion_tokens)
  const usage = parsed.usage || {};
  const inputTokens = Number(
    usage.input_tokens !== undefined ? usage.input_tokens : (usage.prompt_tokens || 0)
  ) || 0;
  const outputTokens = Number(
    usage.output_tokens !== undefined ? usage.output_tokens : (usage.completion_tokens || 0)
  ) || 0;

  if (typeof logApiUsage_ === 'function') {
    logApiUsage_({
      campaignId: request.campaignId,
      serviceName: 'OPENAI',
      operation: request.operation || 'OpenAIResponse',
      httpStatus: status,
      inputUnits: inputTokens,
      outputUnits: outputTokens,
      status: 'SUCCESS',
      details: 'Input tokens: ' + inputTokens + ', Output tokens: ' + outputTokens,
      modelName: request.model
    });
  }

  return parsed;
}

/**
 * Guarded Chat Completions helper (/v1/chat/completions) with full ApiTracker budget & usage logging.
 */
function requestOpenAIChatCompletion_(options) {
  if (!options || !options.prompt) {
    throw new Error('OpenAI Chat Completion requires a prompt.');
  }

  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(options.campaignId);
  }

  const apiKey = resolveValidOpenAiApiKey_();
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY is not configured in Project Settings (Script Properties).');
  }

  const model = options.model || ((typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultOpenAiModel) ? APP_CONFIG.defaultOpenAiModel : 'gpt-4o');
  const chatUrl = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.openAiChatCompletionsUrl)
    ? APP_CONFIG.openAiChatCompletionsUrl
    : 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: model,
    messages: [{ role: 'user', content: options.prompt }],
    temperature: options.temperature !== undefined ? options.temperature : 0.2
  };
  if (options.json) {
    payload.response_format = { type: 'json_object' };
  }

  let res, statusCode, rawText;
  try {
    res = UrlFetchApp.fetch(chatUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    statusCode = res.getResponseCode();
    rawText = res.getContentText();
  } catch (netErr) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: options.campaignId,
        serviceName: 'OPENAI',
        operation: options.operation || 'OpenAIChatCompletion',
        httpStatus: 500,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: netErr.toString(),
        modelName: model
      });
    }
    throw new Error('OpenAI Network Error: ' + netErr.toString());
  }

  if (statusCode < 200 || statusCode >= 300) {
    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: options.campaignId,
        serviceName: 'OPENAI',
        operation: options.operation || 'OpenAIChatCompletion',
        httpStatus: statusCode,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: rawText.slice(0, 300),
        modelName: model
      });
    }
    throw new Error('OpenAI HTTP ' + statusCode + ': ' + rawText.slice(0, 250));
  }

  const parsed = JSON.parse(rawText || '{}');
  const usage = parsed.usage || {};
  const inputTokens = Number(usage.prompt_tokens !== undefined ? usage.prompt_tokens : (usage.input_tokens || 0)) || 0;
  const outputTokens = Number(usage.completion_tokens !== undefined ? usage.completion_tokens : (usage.output_tokens || 0)) || 0;

  if (typeof logApiUsage_ === 'function') {
    logApiUsage_({
      campaignId: options.campaignId,
      serviceName: 'OPENAI',
      operation: options.operation || 'OpenAIChatCompletion',
      httpStatus: statusCode,
      inputUnits: inputTokens,
      outputUnits: outputTokens,
      status: 'SUCCESS',
      details: 'Prompt tokens: ' + inputTokens + ', Completion tokens: ' + outputTokens,
      modelName: model
    });
  }

  return parsed;
}

/**
 * Resolves a valid OPENAI_API_KEY, ignoring sheet status labels and placeholder strings.
 */
function resolveValidOpenAiApiKey_() {
  const propName = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.openAiApiKeyProperty)
    ? APP_CONFIG.openAiApiKeyProperty
    : 'OPENAI_API_KEY';

  const scriptVal = PropertiesService.getScriptProperties().getProperty(propName);
  if (isValidOpenAiKeyValue_(scriptVal)) return String(scriptVal).trim();

  const userVal = PropertiesService.getUserProperties().getProperty(propName);
  if (isValidOpenAiKeyValue_(userVal)) return String(userVal).trim();

  if (typeof getStudioApiKey === 'function') {
    const studioVal = getStudioApiKey('OPENAI');
    if (isValidOpenAiKeyValue_(studioVal)) return String(studioVal).trim();
  }

  return '';
}

function isValidOpenAiKeyValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const upper = s.toUpperCase();
  if (upper.indexOf('PASTE_') >= 0) return false;
  if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
  if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
  return true;
}