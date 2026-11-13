/**
 * ============================================================================
 * AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
 * RESEARCH SCHEMA DEFINITIONS (DUAL AI ROUTING & GROUNDING CONTRACTS)
 * ============================================================================
 * File: ResearchSchema.js
 * 
 * Defines strict JSON schemas for structured LLM outputs across both
 * Google Gemini (gemini-3.8-flash) and OpenAI (gpt-4o). Enforces required
 * company firmographics, persona signals, and citation evidence contracts.
 * ============================================================================
 */

/**
 * Returns the raw JSON Schema for Account-level Company Research and Grounded Evidence.
 * Compatible directly with Gemini responseSchema and OpenAI json_schema.
 * 
 * @return {Object} Raw JSON Schema object
 */
function getCompanyResearchRawSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['company', 'evidence'],
    properties: {
      company: {
        type: 'object',
        additionalProperties: false,
        required: ['name', 'website', 'headquarters', 'industry', 'company_size', 'business_summary'],
        properties: {
          name: {
            type: 'string',
            description: 'Official company name. Use an empty string if not supported by evidence.'
          },
          website: {
            type: 'string',
            description: 'Official company website URL or domain, or an empty string if unknown.'
          },
          headquarters: {
            type: 'string',
            description: 'Country and region of headquarters, or an empty string if unknown.'
          },
          industry: {
            type: 'string',
            description: 'Main industry classification, or an empty string if unknown.'
          },
          company_size: {
            type: 'string',
            description: 'Publicly documented size range or employee count; otherwise use Unknown.'
          },
          business_summary: {
            type: 'string',
            description: 'Brief description of documented core business activities. Do not infer unsupported claims.'
          }
        }
      },
      evidence: {
        type: 'array',
        description: 'Verifiable factual claims supported by live search grounding sources.',
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['claim', 'extracted_information', 'source_url', 'source_type', 'suggested_quality', 'quality_reason'],
          properties: {
            claim: {
              type: 'string',
              description: 'One important factual claim about the company.'
            },
            extracted_information: {
              type: 'string',
              description: 'The relevant information stated by the source, summarized faithfully.'
            },
            source_url: {
              type: 'string',
              description: 'The exact public source URL supporting this claim. Use an empty string if no source is available.'
            },
            source_type: {
              type: 'string',
              enum: [
                'company_website',
                'annual_or_sustainability_report',
                'regulatory_or_public_registry',
                'reliable_industry_source',
                'other_public_source',
                'unknown'
              ]
            },
            suggested_quality: {
              type: 'string',
              enum: ['High', 'Medium', 'Low', 'Unknown']
            },
            quality_reason: {
              type: 'string',
              description: 'Brief reason for the suggested evidence quality. This is an AI suggestion, not independent verification.'
            }
          }
        }
      }
    }
  };
}

/**
 * Returns OpenAI-formatted structured output schema for Company Research.
 * Maintained for backward compatibility with OpenAI API and offline test suites.
 * 
 * @return {Object} OpenAI response_format payload
 */
function getCompanyResearchTextFormat_() {
  return {
    format: {
      type: 'json_schema',
      name: 'company_research',
      strict: true,
      schema: getCompanyResearchRawSchema_()
    }
  };
}

/**
 * Returns Gemini-formatted responseSchema for Company Research.
 * Passed to generationConfig.responseSchema when calling gemini-3.8-flash.
 * 
 * @return {Object} Gemini responseSchema definition
 */
function getGeminiCompanyResearchSchema_() {
  return getCompanyResearchRawSchema_();
}

/**
 * Returns the raw JSON Schema for Persona-level Research and Grounded Evidence.
 * Used by Persona Mapper (Agent 2) to log evidence under Entity Type = 'Persona'.
 * 
 * @return {Object} Raw JSON Schema object
 */
function getPersonaResearchRawSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['company_name', 'persona', 'evidence'],
    properties: {
      company_name: {
        type: 'string',
        description: 'The target company name the persona belongs to.'
      },
      persona: {
        type: 'object',
        additionalProperties: false,
        required: ['name', 'title', 'persona_type', 'linkedin_url', 'buying_signal', 'confidence'],
        properties: {
          name: { type: 'string', description: 'Full name of the executive or decision maker.' },
          title: { type: 'string', description: 'Exact job title of the decision maker.' },
          persona_type: { type: 'string', description: 'Functional classification (e.g., Executive Sponsor, Champion, Technical Evaluator).' },
          linkedin_url: { type: 'string', description: 'Public LinkedIn profile URL or empty string.' },
          buying_signal: { type: 'string', description: 'Documented trigger or buying signal relevant to the campaign focus.' },
          confidence: { type: 'string', enum: ['High', 'Medium', 'Low', 'Unknown'] }
        }
      },
      evidence: {
        type: 'array',
        description: 'Verifiable sources supporting the persona identification and buying signal.',
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['claim', 'extracted_information', 'source_url', 'source_type', 'suggested_quality', 'quality_reason'],
          properties: {
            claim: { type: 'string', description: 'Specific claim regarding role, responsibilities, or buying signal.' },
            extracted_information: { type: 'string', description: 'Quote or summarized statement from the source.' },
            source_url: { type: 'string', description: 'Public URL validating this persona or buying signal.' },
            source_type: {
              type: 'string',
              enum: [
                'company_website',
                'annual_or_sustainability_report',
                'regulatory_or_public_registry',
                'reliable_industry_source',
                'other_public_source',
                'unknown'
              ]
            },
            suggested_quality: { type: 'string', enum: ['High', 'Medium', 'Low', 'Unknown'] },
            quality_reason: { type: 'string', description: 'Reasoning for quality assessment.' }
          }
        }
      }
    }
  };
}

/**
 * Returns OpenAI-formatted structured output schema for Persona Research.
 * 
 * @return {Object} OpenAI response_format payload
 */
function getPersonaResearchTextFormat_() {
  return {
    format: {
      type: 'json_schema',
      name: 'persona_research',
      strict: true,
      schema: getPersonaResearchRawSchema_()
    }
  };
}

/**
 * Returns Gemini-formatted responseSchema for Persona Research.
 * 
 * @return {Object} Gemini responseSchema definition
 */
function getGeminiPersonaResearchSchema_() {
  return getPersonaResearchRawSchema_();
}