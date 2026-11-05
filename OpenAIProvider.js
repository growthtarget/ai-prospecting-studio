function requestOpenAIResponse_(request) {
  if (!request.model) {
    throw new Error('Choose a model in the campaign or agent settings before sending an OpenAI request.');
  }
  if (request.input === undefined || request.input === null || request.input === '') {
    throw new Error('AI request input is required.');
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

  const response = UrlFetchApp.fetch(APP_CONFIG.openAiResponsesUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(
      'OpenAI request failed with HTTP ' + status + '. Check billing, model access, and request settings.'
    );
  }

  return JSON.parse(response.getContentText());
}