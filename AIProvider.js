function requestAIResponse(request) {
  if (!request || typeof request !== 'object') {
    throw new Error('An AI request object is required.');
  }

  const provider = String(
    request.provider || APP_CONFIG.defaultAIProvider
  ).toLowerCase();

  if (provider === 'openai') {
    return requestOpenAIResponse_(request);
  }

  throw new Error('Unsupported AI provider: ' + provider);
}