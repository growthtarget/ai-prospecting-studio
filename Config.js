/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * GLOBAL SYSTEM CONFIGURATION & WORKSPACE DEFINITIONS
 * ============================================================================
 * File: Config.gs / Config.js
 * 
 * Provides frozen master settings, API property keys, model designations,
 * and canonical header schemas across all 12 project tabs.
 * ============================================================================
 */

const APP_CONFIG = Object.freeze({
  appName: 'AI Prospecting Studio + Persona Mapper',
  version: '2.0',
  defaultAIProvider: 'GEMINI',
  defaultGeminiModel: 'gemini-3.8-flash',
  defaultOpenAiModel: 'gpt-4o',
  geminiApiKeyProperty: 'GEMINI_API_KEY',
  openAiApiKeyProperty: 'OPENAI_API_KEY',
  openAiResponsesUrl: 'https://api.openai.com/v1/responses',
  openAiChatCompletionsUrl: 'https://api.openai.com/v1/chat/completions',
  defaultEmailVerifier: 'ZeroBounce',
  defaultSenderMode: 'Workspace Gmail (Drafts Only)',
  
  // 12 Confirmed Tab Schemas
  sheetDefinitions: [
    {
      name: 'Dashboard',
      headers: ['Campaign ID', 'Campaign Name', 'Run Status', 'Companies Found', 'Qualified', 'Need Review', 'Last Run']
    },
    {
      name: 'Campaigns',
      headers: [
        'Campaign ID', 'Campaign Name', 'Department', 'Status', 'Target Countries',
        'Target Industries', 'Company Size', 'Discovery Limit', 'Research Depth',
        'Evidence Period Months', 'Min Company Score', 'Min Persona Score',
        'Min Email Priority Score', 'Target Titles', 'Campaign Focus', 'Search Keywords',
        'Custom Business Needs', 'AI Provider', 'AI Model', 'Email Verifier Provider',
        'Email Verifier Endpoint', 'Email Sender Provider', 'Email Sender Endpoint',
        'Sender Name', 'Sender Title', 'Sender Company', 'Campaign Config JSON', 'Last Modified'
      ]
    },
    {
      name: 'Company Intelligence',
      headers: [
        'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Country',
        'Industry', 'Company Size', 'ESG Frameworks', 'Reporting Obligations',
        'ESG Reporting Maturity', 'Supply Chain Profile', 'Scope 3 Exposure',
        'Sustainability Data Needs', 'Digital Transformation Signals', 'ESG Governance',
        'Business Triggers', 'eValuater Solution Fit', 'Company Opportunity Score',
        'Company Opportunity', 'Evidence URLs', 'Evidence Checked Date',
        'Verification Status', 'Human Review Status', 'Last Enriched At', 'Company Score Rationale'
      ]
    },
    {
      name: 'Persona Leads',
      headers: [
        'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Industry',
        'Persona ID', 'Name', 'Title', 'Persona Type', 'LinkedIn URL', 'Signal',
        'Signal Date', 'Evidence URL', 'Buying Signal', 'Verification Status',
        'Confidence', 'Persona Fit Score', 'Personalized Hook', 'Review Status',
        'Last Verified At', 'Business Email', 'Email Domain', 'Email Source URL',
        'Email Discovery Method', 'Email Verification Status', 'Email Verification Date',
        'Email Verification Provider', 'Email Identity Confidence', 'Email Verification Notes'
      ]
    },
    {
      name: 'Persona Mapper Settings',
      headers: [
        'Campaign ID', 'Campaign Name', 'Persona Mapper Enabled', 'AI Provider',
        'AI Model', 'Persona Titles', 'Search Keywords', 'Campaign Focus',
        'Evidence Period Months', 'Research Depth', 'Include ESG Reporting Obligations',
        'Include ESG Reporting Maturity', 'Include Supply Chain Complexity',
        'Include Scope 3 Exposure', 'Include Sustainability Data Challenges',
        'Include Digital Transformation', 'Include ESG Governance', 'Include Business Triggers',
        'Include eValuater Solution Fit', 'Weight ESG Reporting Relevance',
        'Weight Supply Chain Complexity', 'Weight Sustainability Data Needs',
        'Weight Digital Transformation Relevance', 'Weight Business Triggers',
        'Weight eValuater Solution Fit', 'Minimum Company Opportunity Score',
        'Minimum Persona Confidence', 'Minimum Email Priority Score',
        'Human Approval Required', 'Auto Send Enabled', 'Email Priority Weight Company Score',
        'Email Priority Weight Persona Fit', 'Email Priority Weight Buying Signal',
        'Notes', 'Max Companies Per Run', 'Email Verification Provider',
        'Email Verification Endpoint', 'Email Sender Provider', 'Email Sender Endpoint'
      ]
    },
    {
      name: 'Prospect Review Queue',
      headers: [
        'Campaign ID', 'Campaign Name', 'Company ID', 'Company Name', 'Persona ID',
        'Name', 'Title', 'Business Email', 'Company Opportunity Score', 'Persona Fit Score',
        'Buying Signal', 'Priority Score', 'Priority Tier', 'Priority Score Rationale',
        'Draft Subject', 'Draft Email', 'Evidence URLs', 'Review Status', 'Reviewer',
        'Reviewer Notes', 'Reviewed At', 'Human Approval Status', 'Approved By',
        'Approved At', 'Send Status', 'Sent At', 'Send Error', 'Country', 'Industry',
        'Company Opportunity', 'Persona Type', 'LinkedIn URL', 'Email Domain',
        'Email Source URL', 'Email Discovery Method', 'Email Verification Status',
        'Email Verification Date', 'Email Verification Provider', 'Email Identity Confidence',
        'Email Verification Notes', 'Queue Status', 'Compliance Status', 'Suppression Check',
        'Date Added', 'Personalized Hook'
      ]
    },
    {
      name: 'Campaign Configuration',
      headers: ['Parameter', 'Configured Value', 'Description']
    },
    {
      name: 'Company Database',
      headers: [
        'Company ID', 'Campaign ID', 'Company Name', 'Domain', 'Country',
        'Industry', 'Employee Size', 'Discovery Method', 'Discovery Date', 'Status'
      ]
    },
    {
      name: 'Research Evidence',
      headers: [
        'Evidence ID', 'Campaign ID', 'Entity Name', 'Entity Type',
        'Search Query Used', 'Grounding Evidence URL', 'Extracted Snippet / Signal',
        'Verification Score', 'Date Captured'
      ]
    },
    {
      name: 'Agent Settings',
      headers: ['Setting name', 'Value', 'Description']
    },
    {
      name: 'Execution Logs',
      headers: ['Timestamp', 'Campaign ID', 'Action', 'Target', 'Status', 'Details']
    },
    {
      name: 'Lists & Settings',
      headers: ['List name', 'Value', 'Active status']
    }
  ]
});