\# AI PROSPECTING STUDIO

\## Project Master Brief \& Development Specification



\*\*Project:\*\* AI Prospecting Studio  

\*\*Version:\*\* 2.0  

\*\*Development environment:\*\* ChatGPT Codex Desktop  

\*\*Primary AI provider:\*\* OpenAI API  

\*\*Application interface:\*\* Google Sheets  

\*\*Backend:\*\* Google Apps Script



\---



\# 1. Project Vision



Build a reusable, configurable, AI-powered company research and prospecting platform.



The platform will allow users to create different prospecting campaigns for different industries, markets, countries, business objectives and target customer profiles.



Each campaign will have its own research instructions, target company criteria, qualification rules, exclusions, scoring logic and output requirements.



The objective is to automate the process of identifying, researching, evaluating, verifying and prioritizing prospective companies using AI.



This must be a reusable platform, not a solution hardcoded for one industry, one company or one campaign.



\# 2. Core Technology Decisions



The following technology choices have been made:



| Component | Technology |

|---|---|

| Development partner | ChatGPT Codex Desktop |

| AI engine | OpenAI API |

| AI interaction interface | OpenAI Responses API |

| Application interface | Google Sheets |

| Backend | Google Apps Script |

| Data storage | Google Sheets |

| Document storage | Google Drive |

| API key security | Google Apps Script Properties |

| Local development | Git and clasp CLI, introduced when needed |



\*\*Important:\*\* The application must not depend on Gemini or Google-specific AI models.



Google products are being used for application infrastructure and data management, not as a requirement for the AI engine.



\# 3. AI Architecture



\## 3.1 OpenAI as the primary AI provider



Use the OpenAI API for:



\- Company research and information synthesis.

\- Business activity and industry classification.

\- Company qualification against campaign criteria.

\- Analysis of sustainability commitments and reporting obligations when relevant.

\- Evidence interpretation and verification assistance.

\- Company scoring and prioritization.

\- Structured data extraction.

\- Research summaries and recommendations.



Use the OpenAI Responses API as the initial integration approach.



Select the model based on the task, quality requirements, speed and cost. Model names must be configurable rather than hardcoded throughout the application.



\## 3.2 Modular AI provider architecture



Create a provider-independent AI integration layer.



Suggested architecture:



\- AI Provider Interface

\- OpenAI Provider (initial implementation)

\- Gemini Provider (optional future implementation)

\- Other Providers (future expansion)



All AI agents must communicate through this provider layer rather than calling provider-specific APIs directly.



This allows us to change models or introduce alternative providers without rewriting the application logic.



\## 3.3 Research sources



AI-generated answers must not be treated as independent proof of company facts.



Design a research-source layer that can support:



\- Company websites.

\- Annual reports and sustainability reports.

\- Regulatory filings and public registries.

\- Reliable industry databases.

\- Search and web research APIs.

\- Other authorized public information sources.



Keep research-source integrations separate from AI-provider integrations.



Do not assume the OpenAI API alone provides unrestricted live web access. Any external search or browsing capability must be explicitly configured through an appropriate tool or research provider.



\# 4. Application Architecture



The application should consist of three primary layers.



\### Layer 1: Google Sheets Interface



Google Sheets is the main user interface and structured data store.



Users should be able to:



\- Create campaigns.

\- Define target company profiles.

\- Configure research instructions.

\- Set qualification and exclusion criteria.

\- Start and monitor research.

\- Review results.

\- Edit and approve company records.

\- Filter, sort and export data.

\- Re-run research and update records.



Provide a custom menu and, where practical, a user-friendly sidebar.



\### Layer 2: Google Apps Script Backend



Google Apps Script will manage:



\- Application logic.

\- Campaign configuration.

\- AI provider integration.

\- Research orchestration.

\- Data validation.

\- Sheet updates.

\- Execution status.

\- Error handling.

\- Logging.

\- API configuration.



Use modular files and functions rather than one large script.



\### Layer 3: AI and Research Services



The backend communicates with:



\- OpenAI API.

\- Configured research and search services.

\- Other external data sources as required.



All integrations must use secure server-side requests.



\# 5. AI Agents



Build the agents incrementally. Do not attempt to implement all agents at once.



\## Agent 1: Research Agent



Purpose: Discover and research companies matching a campaign's target profile.



Responsibilities:



\- Identify relevant companies.

\- Collect company information.

\- Identify industry and business activities.

\- Find official company websites.

\- Gather relevant public evidence.

\- Record source URLs and research dates.

\- Return structured company records.



\## Agent 2: Qualification Agent



Purpose: Determine whether researched companies meet campaign requirements.



Responsibilities:



\- Evaluate company size, industry, geography and business profile.

\- Apply mandatory criteria.

\- Identify excluded companies.

\- Explain qualification decisions.

\- Distinguish confirmed information from assumptions.

\- Mark insufficient evidence as unknown.



\## Agent 3: Verification Agent



Purpose: Improve the reliability of company records.



Responsibilities:



\- Check important claims against source evidence.

\- Identify conflicting information.

\- Detect missing or outdated information.

\- Flag weak or unsupported evidence.

\- Identify possible duplicate companies.

\- Recommend records for manual review.



\## Agent 4: Scoring Agent



Purpose: Prioritize qualified companies.



Responsibilities:



\- Apply campaign-specific scoring rules.

\- Evaluate company fit.

\- Score individual criteria.

\- Explain scoring decisions.

\- Identify high-priority prospects.

\- Maintain transparent scoring logic.



Scoring rules must be configurable for each campaign.



\## Agent 5: Reporting Agent



Purpose: Turn research results into useful prospecting outputs.



Responsibilities:



\- Summarize campaign results.

\- Produce company profiles.

\- Generate research summaries.

\- Identify strongest prospects.

\- Highlight missing information.

\- Provide campaign-level statistics.



\# 6. Campaign Configuration



Every campaign must be independently configurable.



Campaign configuration should include:



\- Campaign name.

\- Campaign objective.

\- Target industries.

\- Target countries and regions.

\- Company size.

\- Target business activities.

\- Company qualification criteria.

\- Mandatory requirements.

\- Exclusion criteria.

\- Research instructions.

\- Evidence requirements.

\- Scoring criteria and weights.

\- Required output fields.

\- Maximum number of companies.

\- Research depth.

\- AI model selection.

\- Research-source configuration.



Do not hardcode campaign-specific criteria into the application.



Create reusable campaign templates where practical.



\# 7. Proposed Google Sheets Structure



Create the following worksheets:



| Worksheet | Purpose |

|---|---|

| Dashboard | Campaign overview and performance |

| Campaigns | Campaign creation and management |

| Campaign Configuration | Detailed campaign criteria |

| Company Database | Main company records |

| Research Evidence | Sources and supporting evidence |

| Agent Settings | AI and agent configuration |

| Execution Logs | Activity, errors and progress |

| Lists \& Settings | Reusable dropdowns and system settings |



The exact schema should be proposed and approved before implementation.



\# 8. Data Quality and Research Standards



Data reliability is a core product requirement.



The platform must:



1\. Record the source URL for researched information.

2\. Record when information was collected.

3\. Distinguish verified facts from AI interpretation.

4\. Mark unavailable information as unknown rather than inventing it.

5\. Identify conflicting evidence.

6\. Avoid fabricated companies, contacts, financial figures or commitments.

7\. Detect duplicate companies.

8\. Maintain traceability between research results and their supporting evidence.

9\. Allow human review of uncertain or important results.

10\. Avoid presenting AI-generated scores as objectively verified facts.



Where possible, prioritize official company and regulatory sources.



\# 9. User Experience



The platform should be accessible to a non-technical user.



The user should not need to edit code to operate the application.



Prioritize:



\- Guided campaign creation.

\- Clear input fields.

\- Simple configuration options.

\- Start and stop controls where technically supported.

\- Visible execution progress.

\- Clear error messages.

\- Searchable and filterable company tables.

\- Easy review and editing.

\- Campaign duplication.

\- Research reruns.

\- CSV export.



Avoid unnecessary complexity in the initial version.



\# 10. Security and API Management



Security is mandatory.



OpenAI API keys and other credentials must:



\- Never be placed in spreadsheet cells.

\- Never be exposed in client-side interfaces.

\- Never be hardcoded into source files.

\- Be stored in secure server-side configuration, initially Google Apps Script Properties.

\- Never appear in execution logs or error messages.



Use least-privilege access wherever possible.



The application must request approval before sending external communications or performing consequential actions.



\# 11. Reliability, Limits and Cost Management



Design around the practical limitations of Google Apps Script and external APIs.



Consider:



\- Execution time limits.

\- API rate limits.

\- Request failures.

\- Large research batches.

\- Retry logic.

\- Partial execution.

\- Duplicate processing.

\- API usage costs.

\- Logging and recovery.



Implement batch processing and resumable workflows where appropriate.



Provide configurable limits for research volume and AI usage.



Do not assume every campaign can be completed in a single Apps Script execution.



\# 12. Development Roadmap



Develop the application in controlled phases.



\## Phase 1: Foundation



\- Establish project folder and code structure.

\- Connect to Google Sheets.

\- Create the initial spreadsheet architecture.

\- Implement configuration management.

\- Establish secure API key storage.

\- Create the OpenAI provider integration.

\- Test a basic OpenAI API request.



\## Phase 2: First Research Agent



\- Build the Research Agent.

\- Define structured output schemas.

\- Integrate an appropriate research source.

\- Write company records into Google Sheets.

\- Store source evidence.

\- Validate results.



\## Phase 3: Qualification



\- Build configurable qualification logic.

\- Add inclusion and exclusion rules.

\- Separate mandatory criteria from weighted criteria.

\- Explain qualification decisions.



\## Phase 4: Verification and Scoring



\- Add evidence checks.

\- Implement duplicate detection.

\- Build campaign-specific scoring.

\- Add confidence and evidence-quality indicators.



\## Phase 5: Reporting and Automation



\- Create the Dashboard.

\- Add campaign summaries.

\- Implement resumable batch processing.

\- Add export functionality.

\- Improve error handling and monitoring.



\## Phase 6: Advanced Capabilities



Potential future additions:



\- Multiple AI providers.

\- Multiple research providers.

\- Contact discovery.

\- CRM integrations.

\- Scheduled campaigns.

\- Advanced reporting.

\- Team collaboration.

\- User permissions.

\- Campaign performance analytics.



These features should not delay the initial working version.



\# 13. Development Rules for Codex



Codex is the primary development partner.



Follow these rules throughout development:



1\. Assume the project owner has no programming experience.

2\. Explain technical concepts in plain language.

3\. Work in small, manageable steps.

4\. Do not make major architectural changes without explaining them.

5\. Ask for approval before implementing major features.

6\. Never expose or request API secrets in ordinary chat messages.

7\. Explain which accounts, services and permissions are needed before requesting access.

8\. Create maintainable, modular code.

9\. Test each component before moving to the next phase.

10\. Explain how to operate and troubleshoot the application.

11\. Maintain a project README and update documentation as the application evolves.

12\. Do not silently replace working code or remove existing features.



\# 14. First Assignment for Codex



Before writing application code, review this entire project specification.



Prepare a development plan that addresses:



1\. Recommended technical architecture.

2\. Proposed Google Sheets structure and data schema.

3\. Recommended Apps Script project structure.

4\. OpenAI API integration approach.

5\. Recommended OpenAI model choices for the initial MVP, with cost and capability trade-offs.

6\. Research-source integration options.

7\. Secure API key configuration.

8\. Google account and service permissions required.

9\. Apps Script limitations and how to manage them.

10\. Development phases and milestones.

11\. Estimated complexity and potential risks.

12\. The simplest practical approach for a non-technical owner.



Identify any missing decisions or potential problems.



\*\*Do not start coding yet.\*\*



Present the proposed architecture and implementation plan in plain English. Wait for the project owner's approval before proceeding.



\---



\# 15. Definition of Success



The first successful version should allow a non-technical user to:



1\. Open a Google Sheet.

2\. Create a new prospecting campaign.

3\. Define the target company profile.

4\. Configure the research and qualification criteria.

5\. Start an AI-powered company research process.

6\. Receive structured company records.

7\. Review supporting sources.

8\. See qualification results and scores.

9\. Filter and prioritize prospective companies.

10\. Export the results.



The application must be reusable for different industries, countries, business models and prospecting objectives.



\*\*Guiding principle: Build a reliable, reusable prospecting platform—not a collection of one-off AI prompts.\*\*

