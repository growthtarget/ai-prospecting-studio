function getCompanyResearchTextFormat_() {
  return {
    format: {
      type: 'json_schema',
      name: 'company_research',
      strict: true,
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['company', 'evidence'],
        properties: {
          company: {
            type: 'object',
            additionalProperties: false,
            required: ['name', 'website', 'headquarters', 'industry', 'company_size', 'business_summary'],
            properties: {
              name: { type: 'string', description: 'Company name. Use an empty string if not supported by evidence.' },
              website: { type: 'string', description: 'Official company website URL, or an empty string if unknown.' },
              headquarters: { type: 'string', description: 'Country and region, or an empty string if unknown.' },
              industry: { type: 'string', description: 'Main industry or industries, or an empty string if unknown.' },
              company_size: { type: 'string', description: 'Publicly supported size range or employee count; otherwise use Unknown.' },
              business_summary: { type: 'string', description: 'Brief description of documented business activities. Do not infer unsupported claims.' }
            }
          },
          evidence: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['claim', 'extracted_information', 'source_url', 'source_type', 'suggested_quality', 'quality_reason'],
              properties: {
                claim: { type: 'string', description: 'One important factual claim about the company.' },
                extracted_information: { type: 'string', description: 'The relevant information stated by the source, summarized faithfully.' },
                source_url: { type: 'string', description: 'The exact public source URL supporting this claim. Use an empty string if no source is available.' },
                source_type: {
                  type: 'string',
                  enum: ['company_website', 'annual_or_sustainability_report', 'regulatory_or_public_registry', 'reliable_industry_source', 'other_public_source', 'unknown']
                },
                suggested_quality: { type: 'string', enum: ['High', 'Medium', 'Low', 'Unknown'] },
                quality_reason: { type: 'string', description: 'Brief reason for the suggested evidence quality. This is an AI suggestion, not independent verification.' }
              }
            }
          }
        }
      }
    }
  };
}