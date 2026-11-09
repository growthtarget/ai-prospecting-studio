/**
 * AI Provider Universal Router
 * Dispatches requests to Gemini or OpenAI based on configuration or settings.
 */
function callAI(options) {
  if (!options || !options.prompt) {
    throw new Error('callAI requires an options object with a prompt.');
  }

  // Determine provider: passed in options -> or configured default -> fallback to GEMINI
  const provider = String(
    options.provider || APP_CONFIG.defaultAIProvider || 'GEMINI'
  ).toUpperCase();

  let response;

  if (provider === 'GEMINI') {
    response = callGemini_(options);
  } else if (provider === 'OPENAI') {
    // Bridges to your existing OpenAI implementation
    const openAiPayload = {
      model: options.model || 'gpt-4o',
      input: options.prompt,
      tools: options.enableWebSearch !== false ? [{ type: 'web_search' }] : []
    };
    if (options.json) {
      openAiPayload.text = { format: { type: 'json_object' } };
    }
    const rawOpenAi = requestOpenAIResponse_(openAiPayload);
    const parsed = parsePersonaMapperResponse_(rawOpenAi);
    const citations = collectPersonaMapperUrls_(rawOpenAi);

    response = {
      text: typeof rawOpenAi.output_text === 'string' ? rawOpenAi.output_text : JSON.stringify(parsed),
      json: parsed,
      citations: citations,
      raw: rawOpenAi
    };
  } else {
    throw new Error('Unsupported AI provider: ' + provider);
  }

  return {
    provider: provider,
    text: response.text,
    json: response.json,
    citations: response.citations || [],
    raw: response.raw
  };
}