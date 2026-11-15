/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * GOOGLE GEMINI AI PROVIDER (WITH GOOGLE SEARCH GROUNDING & CLIENT ISOLATION)
 * ============================================================================
 * File: GeminiProvider.gs / GeminiProvider.js
 * 
 * Connects directly to Google Generative Language API (v1beta)
 * with real-time Google Search Grounding, free-tier/unpaid fallback safety,
 * automatic model cascade on 503/429 overload (3.8-flash -> 3.5-flash-lite -> 2.5-flash),
 * fast thinking configuration, resilient JSON parsing, strict client-scoped
 * API key resolution, and ApiTracker logging.
 * ============================================================================
 */

function callGemini_(options) {
  if (!options || !options.prompt) {
    throw new Error('callGemini_ requires an options object with a prompt.');
  }

  const startTimeMs = Date.now();
  const MAX_ELAPSED_MS = 220000; // Stop gracefully before Apps Script's 360s hard limit

  // 0. Pre-Flight Budget Hard Cap Check (Scoped per Client / Owner)
  if (typeof checkApiBudgetGuardrails_ === 'function') {
    checkApiBudgetGuardrails_(options.campaignId, options.clientName);
  }

  // 1. Strict Client-Scoped API Key Resolution (Zero fallback for external clients)
  const apiKey = resolveValidGeminiApiKey_(options);
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing. Add a valid key under Project Settings > Script Properties.');
  }

  // 2. Respect Configured Model & Build Resilient Free/Paid Cascade List
  const defaultModel = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultGeminiModel)
    ? APP_CONFIG.defaultGeminiModel
    : 'gemini-3.8-flash';

  let primaryModel = String(options.model || defaultModel).trim();
  if (!primaryModel || primaryModel === 'gemini-1.5-flash') {
    primaryModel = defaultModel;
  }

  const candidateModels = [primaryModel, 'gemini-3.5-flash-lite', 'gemini-2.5-flash', 'gemini-3.8-flash']
    .filter(function(m, idx, arr) {
      return Boolean(m) && arr.indexOf(m) === idx;
    });

  // 3. Check GOOGLE_SEARCH_GROUNDING setting & recent grounding quota skip cache (keyed by API key suffix)
  const groundingAllowedByConfig = isGeminiSearchGroundingEnabled_(options);
  const groundingSkipKey = 'GEMINI_GROUNDING_SKIP_' + apiKey.slice(-6);
  const groundingRecentlyBlocked = isCacheFlagSet_(groundingSkipKey);
  const wantGrounding = (options.enableWebSearch !== false) && groundingAllowedByConfig && !groundingRecentlyBlocked;

  let responseData = null;
  let lastError = null;

  for (let i = 0; i < candidateModels.length; i++) {
    const activeModel = candidateModels[i];

    if (Date.now() - startTimeMs > MAX_ELAPSED_MS) {
      throw new Error('Gemini request aborted safely before Apps Script timeout. Last error: ' + (lastError ? lastError.message : 'Timeout risk'));
    }

    if (typeof checkApiBudgetGuardrails_ === 'function') {
      checkApiBudgetGuardrails_(options.campaignId, options.clientName);
    }

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(activeModel) + ':generateContent?key=' + encodeURIComponent(apiKey);

    const tryGrounded = wantGrounding && i === 0 && !isCacheFlagSet_(groundingSkipKey);

    if (tryGrounded) {
      try {
        responseData = executeGeminiFetch_(url, activeModel, options, true, true);
        break;
      } catch (groundingErr) {
        lastError = groundingErr;
        if (groundingErr && groundingErr.isGroundingFallbackEligible) {
          Logger.log('Gemini Search Grounding unavailable (' + groundingErr.message + '). Skipping grounding for 10 mins and retrying...');
          if (groundingErr.isQuotaOrTierBlock) {
            setCacheFlag_(groundingSkipKey, 600);
          }

          if (Date.now() - startTimeMs > MAX_ELAPSED_MS) {
            throw groundingErr;
          }

          try {
            responseData = executeGeminiFetch_(url, activeModel, options, false, true);
            break;
          } catch (ungroundedErr) {
            lastError = ungroundedErr;
            if (ungroundedErr && ungroundedErr.isModelFallbackEligible && i < candidateModels.length - 1) {
              Logger.log('Model ' + activeModel + ' overloaded/exhausted (' + ungroundedErr.message + '). Falling back to next Gemini model...');
              continue;
            }
            throw ungroundedErr;
          }
        } else if (groundingErr && groundingErr.isModelFallbackEligible && i < candidateModels.length - 1) {
          Logger.log('Model ' + activeModel + ' unavailable (' + groundingErr.message + '). Falling back to next Gemini model...');
          continue;
        } else {
          throw groundingErr;
        }
      }
    } else {
      try {
        responseData = executeGeminiFetch_(url, activeModel, options, false, true);
        break;
      } catch (err) {
        lastError = err;
        if (err && err.isModelFallbackEligible && i < candidateModels.length - 1) {
          Logger.log('Model ' + activeModel + ' unavailable (' + err.message + '). Falling back to next Gemini model...');
          continue;
        }
        throw err;
      }
    }
  }

  if (!responseData || !responseData.body) {
    throw lastError || new Error('Gemini returned no response.');
  }

  const body = responseData.body;
  const candidate = body.candidates && body.candidates[0];
  if (!candidate || !candidate.content || !candidate.content.parts) {
    throw new Error('Gemini returned an empty response.');
  }

  // Filter out internal thinking parts (part.thought === true) so JSON extraction stays clean
  const nonThoughtParts = candidate.content.parts.filter(function(p) {
    return p && !p.thought && typeof p.text === 'string';
  });
  const partsToUse = nonThoughtParts.length > 0 ? nonThoughtParts : candidate.content.parts;

  const textOutput = partsToUse.map(function(p) {
    return p.text || '';
  }).join('').trim();

  // 5. Extract Grounding Citations
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

  // 6. Resilient JSON Parsing (Handles objects, arrays, markdown fences, and prose)
  let parsedJson = null;
  if (options.json) {
    parsedJson = parseGeminiJsonOutput_(textOutput);
    if (!parsedJson) {
      Logger.log('Gemini JSON parse failed. Raw output preview: ' + textOutput.slice(0, 500));
    }
  }

  return {
    text: textOutput,
    json: parsedJson,
    citations: Array.from(new Set(citations)),
    usedSearchGrounding: Boolean(responseData.usedGrounding),
    raw: body
  };
}

/**
 * Performs a single Gemini HTTP request and logs tokens/errors to ApiTracker.
 * Applies fast thinkingConfig per model family and clarifies ungrounded free-tier mode.
 */
function executeGeminiFetch_(url, model, options, enableGroundingTool, allowThinkingConfig) {
  let promptText = String(options.prompt || '');

  if (enableGroundingTool && options.json) {
    promptText += '\n\nIMPORTANT: Output ONLY raw valid JSON matching the requested schema. Do not wrap in markdown code fences or add extra commentary.';
  } else if (!enableGroundingTool && options.enableWebSearch !== false) {
    promptText += '\n\n[UNGROUNDED MODE NOTICE: Live web search is unavailable on this API tier. Do NOT return an empty list or refuse because search is unavailable. Instead, use your general knowledge of real, established organizations and executives matching the criteria, populate every required JSON field cleanly, and return valid JSON matching the requested schema.]';
  }

  const generationConfig = {
    temperature: options.temperature !== undefined ? options.temperature : 0.2,
    maxOutputTokens: 8192
  };

  if (allowThinkingConfig) {
    const lowerModel = String(model || '').toLowerCase();
    if (lowerModel.indexOf('gemini-3.') === 0 && lowerModel.indexOf('lite') < 0) {
      generationConfig.thinkingConfig = { thinkingLevel: 'low' };
    } else if (lowerModel.indexOf('gemini-2.5-flash') === 0) {
      generationConfig.thinkingConfig = { thinkingBudget: 0 };
    }
  }

  if (!enableGroundingTool && options.json) {
    generationConfig.responseMimeType = 'application/json';
  }

  const payload = {
    contents: [
      {
        parts: [{ text: promptText }]
      }
    ],
    generationConfig: generationConfig
  };

  if (enableGroundingTool) {
    payload.tools = [{ google_search: {} }];
  }

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
        clientName: options.clientName,
        serviceName: 'GEMINI',
        operation: options.operation || 'CompanyDiscovery',
        httpStatus: 500,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: netErr.toString().slice(0, 600),
        modelName: model
      });
    }
    const netError = new Error('Gemini Network Error (' + model + '): ' + netErr.toString());
    netError.isModelFallbackEligible = true;
    throw netError;
  }

  if (statusCode < 200 || statusCode >= 300) {
    const lowerErr = String(rawText || '').toLowerCase();

    if (allowThinkingConfig && statusCode === 400 && lowerErr.indexOf('thinking') >= 0) {
      return executeGeminiFetch_(url, model, options, enableGroundingTool, false);
    }

    if (typeof logApiUsage_ === 'function') {
      logApiUsage_({
        campaignId: options.campaignId,
        clientName: options.clientName,
        serviceName: 'GEMINI',
        operation: options.operation || 'CompanyDiscovery',
        httpStatus: statusCode,
        inputUnits: 0,
        outputUnits: 0,
        status: 'ERROR',
        details: (enableGroundingTool ? '[Grounded Attempt] ' : '[Standard JSON] ') + rawText.slice(0, 600),
        modelName: model
      });
    }

    const err = new Error('Gemini API Error (' + statusCode + ' on ' + model + '): ' + rawText);

    if (enableGroundingTool && (
      statusCode === 429 ||
      statusCode === 503 ||
      (statusCode === 400 && (lowerErr.indexOf('google_search') >= 0 || lowerErr.indexOf('tool') >= 0 || lowerErr.indexOf('grounding') >= 0 || lowerErr.indexOf('mime') >= 0)) ||
      (statusCode === 403 && (lowerErr.indexOf('grounding') >= 0 || lowerErr.indexOf('search') >= 0))
    )) {
      err.isGroundingFallbackEligible = true;
      err.isQuotaOrTierBlock = (statusCode === 429 || statusCode === 403);
    }

    if (statusCode === 503 || statusCode === 429 || statusCode === 404 || statusCode >= 500) {
      err.isModelFallbackEligible = true;
    }

    throw err;
  }

  const body = JSON.parse(rawText || '{}');
  const usage = body.usageMetadata || {};
  const promptTokens = usage.promptTokenCount || 0;
  const candidateTokens = usage.candidatesTokenCount || 0;

  if (typeof logApiUsage_ === 'function') {
    logApiUsage_({
      campaignId: options.campaignId,
      clientName: options.clientName,
      serviceName: 'GEMINI',
      operation: options.operation || 'CompanyDiscovery',
      httpStatus: statusCode,
      inputUnits: promptTokens,
      outputUnits: candidateTokens,
      status: 'SUCCESS',
      details: (enableGroundingTool ? '[Search Grounding] ' : '[Standard JSON] ') +
        'Model: ' + model + ' | Prompt tokens: ' + promptTokens + ', Candidate tokens: ' + candidateTokens,
      modelName: model
    });
  }

  return { body: body, statusCode: statusCode, usedGrounding: enableGroundingTool };
}

/**
 * Safely parses JSON (object or array) from Gemini output even if wrapped in fences or prose.
 */
function parseGeminiJsonOutput_(textOutput) {
  if (!textOutput || typeof textOutput !== 'string') return null;

  let cleaned = textOutput.trim();

  const fenceMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch && fenceMatch[1]) {
    cleaned = fenceMatch[1].trim();
  }

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const firstBrace = cleaned.indexOf('{');
    const firstBracket = cleaned.indexOf('[');
    let startIdx = -1;
    let endIdx = -1;

    if (firstBracket >= 0 && (firstBrace < 0 || firstBracket < firstBrace)) {
      startIdx = firstBracket;
      endIdx = cleaned.lastIndexOf(']');
    } else if (firstBrace >= 0) {
      startIdx = firstBrace;
      endIdx = cleaned.lastIndexOf('}');
    }

    if (startIdx >= 0 && endIdx > startIdx) {
      try {
        return JSON.parse(cleaned.slice(startIdx, endIdx + 1));
      } catch (nestedErr) {
        Logger.log('Could not parse Gemini JSON slice: ' + nestedErr.message);
      }
    }
    return null;
  }
}

function isCacheFlagSet_(key) {
  try {
    return Boolean(CacheService.getScriptCache().get(key));
  } catch (e) {
    return false;
  }
}

function setCacheFlag_(key, seconds) {
  try {
    CacheService.getScriptCache().put(key, '1', Math.max(60, Math.min(21600, seconds || 600)));
  } catch (e) {}
}

/**
 * Resolves a valid Gemini API key strictly scoped to the campaign's Client / Owner.
 * External clients require CLIENT_{NAME}_GEMINI_API_KEY (zero local fallback).
 */
function resolveValidGeminiApiKey_(options) {
  const propName = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.geminiApiKeyProperty)
    ? APP_CONFIG.geminiApiKeyProperty
    : 'GEMINI_API_KEY';

  if (typeof resolveClientScopedSecret_ === 'function') {
    const resolved = resolveClientScopedSecret_(
      [propName, 'GEMINI_API_KEY'],
      {
        campaignId: options && options.campaignId ? options.campaignId : '',
        clientName: options && options.clientName ? options.clientName : ''
      },
      {
        serviceLabel: 'Gemini AI',
        throwOnMissing: true,
        allowAgentSettingsForLocal: true
      }
    );
    return resolved.key || '';
  }

  const scriptVal = PropertiesService.getScriptProperties().getProperty(propName);
  if (isValidGeminiKeyValue_(scriptVal)) return String(scriptVal).trim();

  const userVal = PropertiesService.getUserProperties().getProperty(propName);
  if (isValidGeminiKeyValue_(userVal)) return String(userVal).trim();

  if (typeof getStudioApiKey === 'function') {
    const studioVal = getStudioApiKey('GEMINI', options);
    if (isValidGeminiKeyValue_(studioVal)) return String(studioVal).trim();
  }

  return '';
}

function isValidGeminiKeyValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const upper = s.toUpperCase();
  if (upper.indexOf('PASTE_') >= 0) return false;
  if (upper.indexOf('CONFIGURED IN SCRIPT PROPERTIES') >= 0) return false;
  if (upper.indexOf('PROTECTED') >= 0 && s.length < 35) return false;
  return true;
}

/**
 * Checks GOOGLE_SEARCH_GROUNDING in client-specific override or Agent Settings (defaults to true if unset).
 */
function isGeminiSearchGroundingEnabled_(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
    if (options && typeof resolveCampaignClientName_ === 'function' && typeof isLocalAgencyClient_ === 'function' && typeof normalizeClientPropertyPrefix_ === 'function') {
      const clientName = resolveCampaignClientName_(options.campaignId, options.clientName, ss);
      if (!isLocalAgencyClient_(clientName)) {
        const overrideProp = normalizeClientPropertyPrefix_(clientName) + 'GOOGLE_SEARCH_GROUNDING';
        const overrideVal = String(PropertiesService.getScriptProperties().getProperty(overrideProp) || '').trim().toUpperCase();
        if (overrideVal === 'FALSE' || overrideVal === 'NO' || overrideVal === '0' || overrideVal === 'DISABLED') {
          return false;
        }
        if (overrideVal === 'TRUE' || overrideVal === 'YES' || overrideVal === '1' || overrideVal === 'ENABLED') {
          return true;
        }
      }
    }

    if (!ss) return true;
    const sheet = ss.getSheetByName('Agent Settings');
    if (!sheet || sheet.getLastRow() <= 1) return true;

    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const key = String(data[r][1] || data[r][0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (key === 'GOOGLESEARCHGROUNDING') {
        const val = String(data[r][2] !== undefined ? data[r][2] : data[r][1]).trim().toUpperCase();
        if (val === 'FALSE' || val === 'NO' || val === '0' || val === 'DISABLED') {
          return false;
        }
        return true;
      }
    }
  } catch (e) {
    Logger.log('isGeminiSearchGroundingEnabled_ check notice: ' + e.message);
  }
  return true;
}