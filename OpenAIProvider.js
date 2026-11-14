/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * OPENAI AI PROVIDER (WITH USAGE TRACKING & BUDGET GUARDRAILS)
 * ============================================================================
 * File: OpenAIProvider.gs / OpenAIProvider.js
 * 
 * Handles fallback LLM calls to OpenAI with automated budget checks,
 * token usage extraction, and API usage logging.
 * ============================================================================
 */

function requestOpenAIResponse_(request) {
  if (!request.model) {
    throw new Error('Choose a model in the campaign or agent settings before sending an OpenAI request.');
  }
  if (request.input === undefined || request.input === null || request.input === '') {
    throw new Error('AI request input is required.');
  }

  // 0. Pre-Flight Budget Hard Cap Check
  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(request.campaignId);
  }

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty(APP_CONFIG.openAiApiKeyProperty);

  if (!apiKey) {
    throw new Error('The OpenAI API key is not configured in Apps Script settings.');
  }

  const payload = {
    model: request.model,
    input: request.input
  };
  if (request.instructions) payload.instructions = request.instructions;
  if (request.tools) payload.tools = request.tools;
  if (request.text) payload.text = request.text;

  let response, status, rawText;
  try {
    response = UrlFetchApp.fetch(APP_CONFIG.openAiResponsesUrl, {
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
        operation: request.operation || 'OpenAIFallback',
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
        operation: request.operation || 'OpenAIFallback',
        httpStatus: status,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: rawText.slice(0, 300),
        modelName: request.model
      });
    }
    throw new Error(
      'OpenAI request failed with HTTP ' + status + '. Check billing, model access, and request settings.'
    );
  }

  const parsed = JSON.parse(rawText);

  // Extract Token Usage
  const usage = parsed.usage || {};
  const promptTokens = usage.prompt_tokens || 0;
  const completionTokens = usage.completion_tokens || 0;

  // Log Successful API Call
  if (typeof logApiUsage_ === 'function') {
    logApiUsage_({
      campaignId: request.campaignId,
      serviceName: 'OPENAI',
      operation: request.operation || 'OpenAIFallback',
      httpStatus: status,
      inputUnits: promptTokens,
      outputUnits: completionTokens,
      status: 'SUCCESS',
      details: 'Prompt tokens: ' + promptTokens + ', Completion tokens: ' + completionTokens,
      modelName: request.model
    });
  }

  return parsed;
}