/**
 * Gemini AI Provider for AI Prospecting Studio
 * Connects directly to Google Generative Language API with Search Grounding
 */
function callGemini_(options) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(
    APP_CONFIG.geminiApiKeyProperty || 'GEMINI_API_KEY'
  );
  
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing. Add it under Project Settings > Script Properties.');
  }

  const model = options.model || APP_CONFIG.defaultGeminiModel || 'gemini-2.5-flash';
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

  // 2. Enable Google Search Grounding for live research
  if (options.enableWebSearch !== false) {
    payload.tools = [{ google_search: {} }];
  }

  // 3. Request Structured JSON if needed
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
  if (textOutput.indexOf('```') === 0) {
    textOutput = textOutput.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
  }

  // 5. Extract Live Google Search Citations automatically
  const citations = [];
  if (candidate.groundingMetadata) {
    const chunks = candidate.groundingMetadata.groundingChunks || [];
    chunks.forEach(function(chunk) {
      if (chunk.web && chunk.web.uri) {
        citations.push(chunk.web.uri);
      }
    });
    const webSearchQueries = candidate.groundingMetadata.webSearchQueries || [];
  }

  // Parse JSON if requested
  let parsedJson = null;
  if (options.json) {
    try {
      parsedJson = JSON.parse(textOutput);
    } catch (e) {
      console.warn('Gemini response was not valid JSON: ' + textOutput);
    }
  }

  return {
    text: textOutput,
    json: parsedJson,
    citations: Array.from(new Set(citations)),
    raw: body
  };
}