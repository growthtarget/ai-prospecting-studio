const APP_CONFIG = Object.freeze({
  appName: 'AI Prospecting Studio',
  defaultAIProvider: 'GEMINI', // Default is now Google Gemini!
  defaultGeminiModel: 'gemini-2.5-flash',
  geminiApiKeyProperty: 'GEMINI_API_KEY',
  openAiApiKeyProperty: 'OPENAI_API_KEY',
  openAiResponsesUrl: 'https://api.openai.com/v1/responses',
  sheetDefinitions: [
    { name: 'Dashboard', headers: ['Campaign ID', 'Campaign name', 'Run status', 'Companies found', 'Qualified', 'Need review', 'Last run'] },
    { name: 'Campaigns', headers: ['Campaign ID', 'Campaign name', 'Objective', 'Status', 'Company limit', 'Research depth', 'Created date', 'Last updated'] },
    { name: 'Campaign Configuration', headers: ['Campaign ID', 'Target industries', 'Countries/regions', 'Company size', 'Relevant activities', 'Required criteria', 'Exclusions', 'Research instructions', 'Evidence expectations', 'Scoring criteria and weights', 'Required output fields', 'Selected model'] },
    { name: 'Company Database', headers: ['Company ID', 'Campaign ID', 'Company name', 'Website', 'Headquarters', 'Industry', 'Company size', 'Business summary', 'Qualification result', 'Qualification explanation', 'Score', 'Evidence quality', 'Review status', 'First researched date', 'Last researched date'] },
    { name: 'Research Evidence', headers: ['Evidence ID', 'Company ID', 'Claim', 'Extracted information', 'Source URL', 'Source type', 'Date collected', 'Independently verified', 'Evidence quality', 'Reviewer notes'] },
    { name: 'Agent Settings', headers: ['Setting name', 'Value', 'Description'] },
    { name: 'Execution Logs', headers: ['Run ID', 'Campaign ID', 'Start time', 'End time', 'Status', 'Companies processed', 'Current step', 'Error summary'] },
    { name: 'Lists & Settings', headers: ['List name', 'Value', 'Active status'] }
  ]
});