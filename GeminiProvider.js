/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * GOOGLE GEMINI AI PROVIDER (WITH GOOGLE SEARCH GROUNDING & CLIENT ISOLATION)
 * ============================================================================
 * File: GeminiProvider.gs / GeminiProvider.js
 *
 * Connects directly to Google Generative Language API (v1beta)
 * with real-time Google Search Grounding, Gemini 4 + Gemini 3 compatibility,
 * free-tier/unpaid fallback safety, cost-aware model cascading
 * (4-hydrogen / 3.8-flash -> 3.6-flash -> 3.5-flash-lite),
 * thinkingConfig management, resilient JSON parsing, strict client-scoped
 * API key resolution, and ApiTracker token/thought logging.
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

  // 2. Respect Configured Model (Including Gemini 4 & Future Models) + Cost-Safe Cascade
  const defaultModel = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.defaultGeminiModel)
    ? String(APP_CONFIG.defaultGeminiModel).trim()
    : 'gemini-4-hydrogen';

  let primaryModel = String(options.model || defaultModel).trim();
  const retiredLegacyModels = {
    'gemini-1.5-flash': true,
    'gemini-1.5-pro': true,
    'gemini-1.0-pro': true,
    'gemini-3-pro-preview': true
  };
  if (!primaryModel || retiredLegacyModels[primaryModel.toLowerCase()]) {
    primaryModel = defaultModel;
  }

  const candidateModels = buildGeminiCandidateModels_(primaryModel, defaultModel);

  // 3. Check GOOGLE_SEARCH_GROUNDING setting & recent grounding quota skip cache (keyed by API key suffix)
  const groundingAllowedByConfig = isGeminiSearchGroundingEnabled_(options);
  const groundingSkipKey = 'GEMINI_GROUNDING_SKIP_' + apiKey.slice(-6);
  const wantGrounding = (options.enableWebSearch !== false) && groundingAllowedByConfig && !isCacheFlagSet_(groundingSkipKey);

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

    // Allow grounding on fallback models too unless a confirmed Search tier/quota block was cached
    const tryGrounded = wantGrounding && !isCacheFlagSet_(groundingSkipKey);

    if (tryGrounded) {
      try {
        responseData = executeGeminiFetch_(url, activeModel, options, true, true);
        break;
      } catch (groundingErr) {
        lastError = groundingErr;

        if (groundingErr && groundingErr.isApiKeyInvalid) {
          throw groundingErr;
        }

        // Case A: Confirmed Search Grounding tier/quota/tool rejection -> cache skip & retry ungrounded
        if (groundingErr && groundingErr.isGroundingFallbackEligible && groundingErr.isQuotaOrTierBlock) {
          const skipSeconds = groundingErr.skipCacheSeconds || 300;
          Logger.log('Gemini Search Grounding tier/quota unavailable (' + groundingErr.message + '). Skipping grounding for ' + skipSeconds + 's and retrying ungrounded...');
          setCacheFlag_(groundingSkipKey, skipSeconds);

          if (Date.now() - startTimeMs > MAX_ELAPSED_MS) {
            throw groundingErr;
          }

          try {
            responseData = executeGeminiFetch_(url, activeModel, options, false, true);
            break;
          } catch (ungroundedErr) {
            lastError = ungroundedErr;
            if (ungroundedErr && ungroundedErr.isModelFallbackEligible && i < candidateModels.length - 1) {
              Logger.log('Model ' + activeModel + ' overloaded/unavailable (' + ungroundedErr.message + '). Falling back to next Gemini model...');
              Utilities.sleep(1000);
              continue;
            }
            throw ungroundedErr;
          }
        }

        // Case B: Transient overload / RPM 429 / model unavailable -> keep grounding ON and try next model
        if (groundingErr && groundingErr.isModelFallbackEligible && i < candidateModels.length - 1) {
          Logger.log('Model ' + activeModel + ' transient error (' + groundingErr.message + '). Retrying next Gemini model WITH Search Grounding...');
          Utilities.sleep(1200);
          continue;
        }

        // Case C: Final model still hit a fallback-eligible error (e.g., ambiguous 429) -> one ungrounded rescue attempt without poisoning cache
        if (groundingErr && groundingErr.isGroundingFallbackEligible && i === candidateModels.length - 1) {
          Logger.log('Final grounded attempt failed (' + groundingErr.message + '). Trying ungrounded rescue without disabling future grounding...');
          responseData = executeGeminiFetch_(url, activeModel, options, false, true);
          break;
        }

        throw groundingErr;
      }
    } else {
      try {
        responseData = executeGeminiFetch_(url, activeModel, options, false, true);
        break;
      } catch (err) {
        lastError = err;
        if (err && err.isApiKeyInvalid) {
          throw err;
        }
        if (err && err.isModelFallbackEligible && i < candidateModels.length - 1) {
          Logger.log('Model ' + activeModel + ' unavailable (' + err.message + '). Falling back to next Gemini model...');
          Utilities.sleep(1000);
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
    const finishInfo = candidate && candidate.finishReason ? ' (finishReason: ' + candidate.finishReason + ')' : '';
    throw new Error('Gemini returned an empty response' + finishInfo + '.');
  }

  if (candidate.finishReason && candidate.finishReason !== 'STOP') {
    Logger.log('Gemini finishReason on ' + (responseData.model || primaryModel) + ': ' + candidate.finishReason);
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
      Logger.log('Gemini Search Grounding Queries (' + webSearchQueries.length + '): ' + webSearchQueries.join(' | '));
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
 * Builds a cost-aware Gemini model cascade list honoring the configured primary model.
 * Never auto-upgrades a budget -lite model into expensive paid-only Carbon/Neon/Argon models.
 */
function buildGeminiCandidateModels_(primaryModel, defaultModel) {
  const cleanPrimary = String(primaryModel || defaultModel || 'gemini-4-hydrogen').trim();
  const lowerPrimary = cleanPrimary.toLowerCase();

  let fallbacks;
  if (lowerPrimary.indexOf('lite') >= 0) {
    fallbacks = [cleanPrimary, 'gemini-3.6-flash', 'gemini-3.8-flash', 'gemini-4-hydrogen'];
  } else {
    fallbacks = [cleanPrimary, 'gemini-4-hydrogen', 'gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite'];
  }

  return fallbacks.filter(function(m, idx, arr) {
    return Boolean(m) && arr.indexOf(m) === idx;
  });
}

/**
 * Builds model-appropriate generationConfig for Gemini 4+, Gemini 3, and legacy models.
 * - Omits temperature on Gemini 4+ per Google migration guidelines.
 * - Raises maxOutputTokens to 32K/64K so thinking tokens do not truncate JSON.
 * - Configures low-latency thinkingLevel (never sends 'minimal' to Gemini 4 or 3.8-flash).
 */
function buildGeminiGenerationConfig_(model, options, enableGroundingTool, allowThinkingConfig) {
  const lowerModel = String(model || '').trim().toLowerCase();
  const isGemini4Plus = /^gemini-(?:[4-9]|\d{2,})(?:[-.]|$)/.test(lowerModel);
  const isGemini3 = /^gemini-3(?:[-.]|$)/.test(lowerModel);
  const isGemini25 = /^gemini-2\.5(?:[-.]|$)/.test(lowerModel);
  const isLite = lowerModel.indexOf('lite') >= 0;

  let defaultMaxTokens = 65536;
  if (lowerModel.indexOf('gemini-4-carbon') === 0 || lowerModel.indexOf('gemini-4-argon') === 0) {
    defaultMaxTokens = 32768;
  } else if (/^gemini-(?:1|2\.0)(?:[-.]|$)/.test(lowerModel)) {
    defaultMaxTokens = 8192;
  }

  const generationConfig = {
    maxOutputTokens: (options && options.maxOutputTokens) ? Number(options.maxOutputTokens) : defaultMaxTokens
  };

  if (!isGemini4Plus) {
    generationConfig.temperature = (options && options.temperature !== undefined) ? options.temperature : 0.2;
  }

  if (allowThinkingConfig) {
    const requestedLevel = (options && options.thinkingLevel)
      ? String(options.thinkingLevel).trim().toLowerCase()
      : '';

    if (isGemini4Plus) {
      generationConfig.thinkingConfig = {
        thinkingLevel: (requestedLevel && requestedLevel !== 'minimal') ? requestedLevel : 'low'
      };
    } else if (isGemini3 && !isLite) {
      generationConfig.thinkingConfig = {
        thinkingLevel: (requestedLevel && requestedLevel !== 'minimal') ? requestedLevel : 'low'
      };
    } else if (isGemini3 && isLite && requestedLevel) {
      generationConfig.thinkingConfig = {
        thinkingLevel: requestedLevel
      };
    } else if (isGemini25 && lowerModel.indexOf('pro') < 0) {
      generationConfig.thinkingConfig = { thinkingBudget: 0 };
    }
  }

  if (!enableGroundingTool && options && options.json) {
    generationConfig.responseMimeType = 'application/json';
  }

  return generationConfig;
}

/**
 * Performs a single Gemini HTTP request and logs tokens/thinking/errors to ApiTracker.
 */
function executeGeminiFetch_(url, model, options, enableGroundingTool, allowThinkingConfig) {
  let promptText = String(options.prompt || '');

  if (enableGroundingTool && options.json) {
    promptText += '\n\nIMPORTANT: Use Google Search grounding to verify real, current facts and official URLs before answering. Do NOT guess or fabricate profile URLs. Output ONLY raw valid JSON matching the requested schema. Do not wrap in markdown code fences or add extra commentary.';
  } else if (!enableGroundingTool && options.enableWebSearch !== false) {
    promptText += '\n\n[UNGROUNDED MODE NOTICE: Live web search is unavailable on this request. Do NOT return an empty list or refuse solely because live search is unavailable—use verified knowledge for established organizations where applicable. CRITICAL ACCURACY RULE: Never invent, guess, or construct URLs, LinkedIn profile slugs, search-result links, or unverified executive names. If an exact profile URL or person name is uncertain, leave that URL blank ("") or omit the unverified person rather than guessing.]';
  }

  const generationConfig = buildGeminiGenerationConfig_(model, options, enableGroundingTool, allowThinkingConfig);

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

    if (lowerErr.indexOf('api key not valid') >= 0 || lowerErr.indexOf('api_key_invalid') >= 0) {
      err.isApiKeyInvalid = true;
      throw err;
    }

    if (enableGroundingTool) {
      const explicitGroundingReject =
        (statusCode === 400 && (
          lowerErr.indexOf('google_search') >= 0 ||
          lowerErr.indexOf('search_retrieval') >= 0 ||
          lowerErr.indexOf('grounding') >= 0 ||
          lowerErr.indexOf('mime') >= 0 ||
          lowerErr.indexOf('tool') >= 0
        )) ||
        (statusCode === 403 && (
          lowerErr.indexOf('grounding') >= 0 ||
          lowerErr.indexOf('search') >= 0 ||
          lowerErr.indexOf('permission') >= 0
        )) ||
        (statusCode === 429 && (
          lowerErr.indexOf('search') >= 0 ||
          lowerErr.indexOf('grounding') >= 0 ||
          lowerErr.indexOf('free_tier') >= 0 ||
          lowerErr.indexOf('free tier') >= 0 ||
          lowerErr.indexOf('quota_limit_value: "0"') >= 0 ||
          lowerErr.indexOf('limit: 0') >= 0
        ));

      if (explicitGroundingReject) {
        err.isGroundingFallbackEligible = true;
        err.isQuotaOrTierBlock = true;
        err.skipCacheSeconds = (statusCode === 429) ? 180 : 600;
      } else if (statusCode === 429 || statusCode === 503) {
        err.isGroundingFallbackEligible = true;
        err.isQuotaOrTierBlock = false;
      }
    }

    const modelTierOrMissing400 = (statusCode === 400 && (
      lowerErr.indexOf('model') >= 0 ||
      lowerErr.indexOf('not found') >= 0 ||
      lowerErr.indexOf('deprecated') >= 0 ||
      lowerErr.indexOf('not supported') >= 0 ||
      lowerErr.indexOf('free tier') >= 0
    ));

    if (statusCode === 503 || statusCode === 429 || statusCode === 404 || statusCode === 403 || statusCode >= 500 || modelTierOrMissing400) {
      err.isModelFallbackEligible = true;
    }

    throw err;
  }

  const body = JSON.parse(rawText || '{}');
  const usage = body.usageMetadata || {};
  const promptTokens = (usage.promptTokenCount || 0) + (usage.toolUsePromptTokenCount || 0);
  const candidateTokens = (usage.candidatesTokenCount || 0) + (usage.thoughtsTokenCount || 0);
  const firstCandidate = body.candidates && body.candidates[0];
  const groundingMeta = firstCandidate && firstCandidate.groundingMetadata ? firstCandidate.groundingMetadata : null;
  const searchQueriesCount = groundingMeta && Array.isArray(groundingMeta.webSearchQueries)
    ? groundingMeta.webSearchQueries.filter(Boolean).length
    : 0;
  const groundingChunksCount = groundingMeta && Array.isArray(groundingMeta.groundingChunks)
    ? groundingMeta.groundingChunks.length
    : 0;

  if (typeof logApiUsage_ === 'function') {
    const modeTag = enableGroundingTool
      ? (searchQueriesCount > 0 || groundingChunksCount > 0
        ? '[Search Grounding] '
        : '[Search Grounding Enabled - 0 Queries] ')
      : '[Standard JSON] ';

    const extraUsage = (usage.thoughtsTokenCount ? (', Thought tokens: ' + usage.thoughtsTokenCount) : '') +
      (enableGroundingTool ? (', Queries: ' + searchQueriesCount + ', Sources: ' + groundingChunksCount) : '');

    logApiUsage_({
      campaignId: options.campaignId,
      clientName: options.clientName,
      serviceName: 'GEMINI',
      operation: options.operation || 'CompanyDiscovery',
      httpStatus: statusCode,
      inputUnits: promptTokens,
      outputUnits: candidateTokens,
      status: 'SUCCESS',
      details: modeTag + 'Model: ' + model + ' | Input tokens: ' + promptTokens + ', Output tokens: ' + candidateTokens + extraUsage,
      modelName: model
    });
  }

  return {
    body: body,
    statusCode: statusCode,
    usedGrounding: enableGroundingTool,
    model: model
  };
}

/**
 * Safely parses JSON (object or array) from Gemini output even if wrapped in fences or prose.
 */
function parseGeminiJsonOutput_(textOutput) {
  if (!textOutput || typeof textOutput !== 'string') return null;

  let cleaned = textOutput.trim();

  const fenceMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch && fenceMatch[1]) {
    const fenced = fenceMatch[1].trim();
    try {
      return JSON.parse(fenced);
    } catch (fenceErr) {
      cleaned = fenced;
    }
  }

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const firstBrace = cleaned.indexOf('{');
    const lastBrace = cleaned.lastIndexOf('}');
    const firstBracket = cleaned.indexOf('[');
    const lastBracket = cleaned.lastIndexOf(']');

    const candidates = [];
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      candidates.push({ index: firstBrace, slice: cleaned.slice(firstBrace, lastBrace + 1) });
    }
    if (firstBracket >= 0 && lastBracket > firstBracket) {
      candidates.push({ index: firstBracket, slice: cleaned.slice(firstBracket, lastBracket + 1) });
    }

    candidates.sort(function(a, b) { return a.index - b.index; });

    for (let i = 0; i < candidates.length; i++) {
      try {
        return JSON.parse(candidates[i].slice);
      } catch (nestedErr) {}
    }

    Logger.log('Could not parse Gemini JSON slice from preview: ' + cleaned.slice(0, 300));
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
    CacheService.getScriptCache().put(key, '1', Math.max(60, Math.min(21600, seconds || 300)));
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
    let clientName = '';
    let isLocal = true;

    if (options && typeof resolveCampaignClientName_ === 'function' && typeof isLocalAgencyClient_ === 'function') {
      clientName = resolveCampaignClientName_(options.campaignId, options.clientName, ss);
      isLocal = isLocalAgencyClient_(clientName);
    }

    const scriptProps = PropertiesService.getScriptProperties();

    if (!isLocal && typeof normalizeClientPropertyPrefix_ === 'function') {
      const overrideProp = normalizeClientPropertyPrefix_(clientName) + 'GOOGLE_SEARCH_GROUNDING';
      const overrideVal = String(scriptProps.getProperty(overrideProp) || '').trim().toUpperCase();
      if (overrideVal === 'FALSE' || overrideVal === 'NO' || overrideVal === '0' || overrideVal === 'DISABLED') {
        return false;
      }
      if (overrideVal === 'TRUE' || overrideVal === 'YES' || overrideVal === '1' || overrideVal === 'ENABLED') {
        return true;
      }
    } else if (isLocal) {
      const localPropVal = String(scriptProps.getProperty('GOOGLE_SEARCH_GROUNDING') || '').trim().toUpperCase();
      if (localPropVal === 'FALSE' || localPropVal === 'NO' || localPropVal === '0' || localPropVal === 'DISABLED') {
        return false;
      }
      if (localPropVal === 'TRUE' || localPropVal === 'YES' || localPropVal === '1' || localPropVal === 'ENABLED') {
        return true;
      }
    }

    if (!ss) return true;
    const sheet = ss.getSheetByName('Agent Settings');
    if (!sheet || sheet.getLastRow() <= 1) return true;

    const targetNormKey = (!isLocal && typeof normalizeClientPropertyPrefix_ === 'function')
      ? (normalizeClientPropertyPrefix_(clientName) + 'GOOGLE_SEARCH_GROUNDING').replace(/[^A-Z0-9]/g, '')
      : 'GOOGLESEARCHGROUNDING';

    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const key = String(data[r][1] || data[r][0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (key === targetNormKey) {
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