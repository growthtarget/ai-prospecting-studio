Here is the complete markdown document reflecting the full architecture, 12-tab schema, dual-AI engine, multi-channel dispatch, suppression guard, and operational workflows of **AI Prospecting Studio + Persona Mapper (v2.0)**.

***

# AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
## System Master Brief, Architecture & Technical Specification

**Project:** AI Prospecting Studio + Persona Mapper  
**Version:** 2.0 (Production Architecture)  
**Repository:** [https://github.com/growthtarget/ai-prospecting-studio](https://github.com/growthtarget/ai-prospecting-studio)  
**Primary AI Engine:** Google Gemini API (`gemini-3.8-flash` with Google Search Grounding)  
**Fallback AI Engine:** OpenAI API (`gpt-4o`)  
**Application Interface:** Google Sheets (12 Dedicated Synchronized Tabs)  
**Backend:** Google Apps Script (Modular Enterprise ES6+)  
**Verification Providers:** ZeroBounce, MillionVerifier, NeverBounce, Custom HTTPS  
**Dispatch Engines:** Google Workspace (Direct & Drafts), Smartlead, Instantly, Custom Webhook  

---

# 1. Executive Summary & System Vision

**AI Prospecting Studio + Persona Mapper (v2.0)** is an autonomous, enterprise-grade B2B outbound intelligence, persona discovery, and hyper-personalized outreach synthesis platform built natively inside Google Sheets and powered by Google Apps Script.

The platform eliminates disconnected SaaS prospecting stacks by consolidating **Market Research (Agent 1: Company Discovery)** and **Decision-Maker Persona Mapping & Draft Synthesis (Agent 2: Persona Mapper)** into a continuous, rule-governed pipeline.

### Core Capabilities:
1. **Dynamic Multi-Department Targeting:** Configurable outbound strategies across ESG, Cybersecurity, Procurement, Enterprise IT, Sales, and Human Resources with custom buying signals and value propositions.
2. **Dual-AI Provider Routing:** Native primary execution on Google Gemini (`gemini-3.8-flash`) utilizing real-time Google Search Grounding with dynamic fallback to OpenAI (`gpt-4o`).
3. **Strict Validation & Suppression Guard:** Active screening against blacklisted domains, global opt-out registries, competitor lists, and generic role-based inboxes (`info@`, `admin@`, `support@`) prior to queueing.
4. **Isolated Multi-Campaign Management:** Robust 28-column master campaign persistence preventing cross-campaign data pollution, backed by dynamic vertical configuration mapping (38 parameters).
5. **Multi-Channel Dispatch & Human-in-the-Loop Review:** Dedicated 45-column prospect review queue with 1-click approvals, automated Gmail Draft generation, and support for enterprise cold email infrastructure (Smartlead, Instantly).
6. **Full-Pipeline Sandbox Simulation:** Integrated pilot runner (`PilotSandboxRunner.gs`) for dry-run verification across all 12 operational tabs with automated cleanup.

---

# 2. 12-Tab System Architecture & Master Schemas

The application operates across 12 tightly synchronized sheets styled with high-contrast Dark Navy headers (`#174ea6`) on soft Light Blue backgrounds (`#e8f0fe`), frozen headers, and native Google Sheets cell validations.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                               12 OPERATIONAL SHEETS                                    │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  01. Dashboard               │ Real-time campaign KPIs, AI health, grounded telemetry │
│  02. Campaigns (28 Cols)     │ Master campaign registry & isolated JSON state storage  │
│  03. Company Intelligence    │ Agent 1 research output, firmographics & fit scores     │
│  04. Persona Leads (29 Cols) │ Agent 2 identified decision-makers & contact details    │
│  05. Persona Mapper Settings │ Master 39-column runtime parameters for Agent 2         │
│  06. Prospect Review Queue   │ 45-column human-in-the-loop triage & draft synthesis    │
│  07. Campaign Configuration  │ 38 vertical runtime parameters for Agent 1              │
│  08. Company Database        │ Historical master account repository                    │
│  09. Research Evidence       │ Grounded URLs, citations, timestamps & audit trail      │
│  10. Agent Settings          │ Provider priorities, temperature, tokens & thresholds   │
│  11. Execution Logs          │ Detailed execution logging, debug metrics & errors      │
│  12. Lists & Settings        │ Validation dropdowns, suppressions & blacklist rules    │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

### Detailed Tab Specifications

#### 1. Dashboard
- **Purpose:** Central operational telemetry, real-time KPI aggregation, and AI provider diagnostics.
- **Key Metrics:** Active Campaign selector, Primary AI Engine status (`gemini-3.8-flash` / Grounding active), API connection health indicator, Total Companies Identified, Qualified Targets, Leads Mapped, Review Queue Count, Approved Leads, and Drafts Created.

#### 2. Campaigns (28 Columns)
- **Purpose:** Master repository maintaining complete state isolation between campaigns.
- **Columns (1–28):** `Campaign ID`, `Campaign Name`, `Status`, `Target Industry`, `Target Geography`, `Target FTE Min`, `Target FTE Max`, `Target Revenue Min`, `Target Revenue Max`, `Departments Selected`, `Custom Business Needs`, `Department Priority 1`, `Department Priority 2`, `Department Priority 3`, `Sender Name`, `Sender Title`, `Sender Company`, `Sender Email`, `AI Model Assigned`, `Email Verification Provider`, `Email Sender Provider`, `Created Date`, `Last Run Date`, `Total Accounts Discovered`, `Total Leads Qualified`, `Total Drafts Created`, `Raw Config JSON Backup`, `Notes`.

#### 3. Company Intelligence (25 Columns)
- **Purpose:** Storage for Agent 1 discovery outputs with enriched firmographic data.
- **Columns (1–25):** `Company ID`, `Campaign ID`, `Company Name`, `Website`, `HQ Country`, `HQ City`, `Primary Industry`, `Secondary Industry`, `Estimated FTE`, `Annual Revenue Estimate`, `Ownership Type`, `Business Model`, `Core Products/Services`, `Strategic Priorities`, `ESG / Tech Focus Area`, `Identified Pain Points / Buying Signals`, `eValuater Solution Fit` (*High, Medium, Low, Disqualified - Direct Competitor*), `Fit Score (1-100)`, `Qualification Reason`, `Exclusion Flags`, `Primary Contact Persona`, `Target Dept`, `Verification Status`, `Discovery Timestamp`, `Audit Notes`.

#### 4. Persona Leads (29 Columns)
- **Purpose:** Decision-makers identified by Agent 2 matching campaign-level department parameters.
- **Columns (1–29):** `Lead ID`, `Company ID`, `Campaign ID`, `Full Name`, `First Name`, `Last Name`, `Job Title`, `Seniority Level`, `Department`, `Sub-Department / Function`, `LinkedIn Profile URL`, `Email Address`, `Email Status` (*Valid, Risky, Catch-All, Invalid, Unknown*), `Phone Number`, `Location / Country`, `Relevance Score (1-100)`, `Decision Authority Level`, `Specific Role Priorities`, `Identified Pain Trigger`, `Personalized Icebreaker Hook`, `Recommended CTA`, `Suppression Check` (*Passed / Flagged*), `Review Status`, `Assigned Sender Account`, `Draft Created Flag`, `Sent Timestamp`, `Email Verification Timestamp`, `Verification Provider Used`, `Audit Notes`.

#### 5. Persona Mapper Settings (39 Canonical Columns)
- **Purpose:** Normalized horizontal parameters feeding Agent 2 synthesis.
- **Canonical Structure:** Direct-mapped from Row 2 onward to guarantee synchronization with `ControlPanel.html` and bypass cell-offset collisions during live execution. Contains comprehensive sender identity records, prompt templates, pain point matrices, and dispatch rules.

#### 6. Prospect Review Queue (45 Columns)
- **Purpose:** Human-in-the-loop review station for staged outreach drafts prior to dispatch.
- **Key Fields:** Lead demographics, company context, generated Subject Lines (Angles A, B, C), Personalized Email Body, Value Proposition Fit, Suppression Check Status, Human Action Dropdown (`Pending Review`, `Approve`, `Edit Required`, `Reject`), Approval Timestamp, Draft ID, and Dispatch Status.

#### 7. Campaign Configuration (38 Vertical Parameters)
- **Purpose:** Normalized vertical key-value pair engine reading runtime parameters for Agent 1.
- **Key Parameters:** `CAMPAIGN_ID` (Row 2), `CAMPAIGN_NAME`, `PRIMARY_AI_PROVIDER`, `GEMINI_MODEL`, `OPENAI_MODEL`, `SEARCH_GROUNDING_ENABLED`, `TARGET_INDUSTRIES`, `TARGET_GEOGRAPHIES`, `MIN_EMPLOYEES`, `MAX_EMPLOYEES`, `CUSTOM_BUSINESS_NEEDS`, `DEPARTMENT_FOCUS`, `SENDER_NAME`, `SENDER_TITLE`, `SENDER_COMPANY`, `SENDER_EMAIL`, `VERIFICATION_SERVICE`, `DISPATCH_PROVIDER`.

#### 8. Company Database
- **Purpose:** Long-term repository of all discovered entities across all campaigns, serving as a deduplication index for subsequent prospecting cycles.

#### 9. Research Evidence
- **Purpose:** Unaltered audit trail recording grounded research URLs, search queries issued, citation snippets, and model retrieval timestamps.

#### 10. Agent Settings
- **Purpose:** Global operational boundaries including API temperature controls, max output token limits, backoff retry multipliers, and provider routing hierarchies.

#### 11. Execution Logs
- **Purpose:** Real-time runtime telemetry tracking script execution times, batch counts, API payload sizes, error stack traces, and debug markers.

#### 12. Lists & Settings
- **Purpose:** System validation references, in-cell dropdown source arrays, domain suppression blacklists, competitor name patterns, and generic inbox filter lists.

---

# 3. Dual-AI Engine & Search Grounding

The system implements an abstracted provider layer (`AIProvider.gs`) decoupling prompt design from specific model architectures.

```
                     ┌───────────────────────────────┐
                     │   Autonomous Prompt Router    │
                     └───────────────┬───────────────┘
                                     │
                    Try Primary      ▼      On Quota/Error
             ┌──────────────────────────────────────────────┐
             │                                              │
             ▼                                              ▼
┌─────────────────────────┐                    ┌─────────────────────────┐
│   Google Gemini API     │                    │       OpenAI API        │
│  (`gemini-3.8-flash`)   │                    │       (`gpt-4o`)        │
├─────────────────────────┤                    ├─────────────────────────┤
│ • Native Search Ground  │                    │ • Fallback Architecture │
│ • Live Web Synthesis    │                    │ • Structured JSON Mode  │
│ • Low Latency Pipeline  │                    │ • High-Reasoning Matrix │
└─────────────────────────┘                    └─────────────────────────┘
```

### Search Grounding & Parsing Workflow
1. **Direct API Callers:** Built natively via `UrlFetchApp` avoiding heavy external SDK dependencies and minimizing runtime overhead.
2. **Search Grounding Injection:** Agent 1 dynamically queries Google Search Grounding to evaluate real-world company data (recent acquisitions, public filings, press releases, live leadership changes).
3. **Structured Response Extraction:** Custom parsers (`parsePersonaMapperResponse_` and `collectPersonaMapperUrls_`) automatically sanitize conversational markdown, extract grounded web sources, and populate `Research Evidence` with live source URLs.

---

# 4. Agent Architecture & Pipeline Execution

```
┌────────────────────────────────────────────────────────────────────────────────┐
│                           END-TO-END EXECUTION FLOW                            │
└────────────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 1. CAMPAIGN SETUP (ControlPanel.html / Campaign Configuration)         │
   │    Select departments, define Custom Business Needs, set sender profile │
   └────────────────────────────────────┬───────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 2. AGENT 1: COMPANY DISCOVERY (CompanyDiscovery.gs)                   │
   │    Queries Gemini 3.8 Flash + Search Grounding                         │
   │    Evaluates firmographics, fit scores & qualification criteria        │
   │    Outputs to: Company Intelligence & Company Database                 │
   └────────────────────────────────────┬───────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 3. AGENT 2: PERSONA MAPPER (PersonaMapper.gs)                         │
   │    Discovers target job titles across designated departments           │
   │    Synthesizes custom icebreakers using Custom Business Needs          │
   │    Applies Active Suppression Guard (Blacklist / Role check)           │
   │    Outputs to: Persona Leads & Prospect Review Queue (45 cols)         │
   └────────────────────────────────────┬───────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 4. HUMAN-IN-THE-LOOP TRIAGE (Prospect Review Queue)                   │
   │    Reviewers evaluate generated copy & select: Approve / Edit / Reject │
   └────────────────────────────────────┬───────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 5. DISPATCH ENGINE (EmailDispatcher.gs / GmailApp)                    │
   │    Generates formatted Gmail Drafts signed by configured Sender Profile│
   │    Or pushes approved leads to Smartlead / Instantly via Webhooks      │
   └────────────────────────────────────┬───────────────────────────────────┘
                                        │
                                        ▼
   ┌────────────────────────────────────────────────────────────────────────┐
   │ 6. TELEMETRY & REPORTING (Dashboard / Execution Logs)                  │
   │    Aggregates conversion metrics, API health & pipeline audit trails   │
   └────────────────────────────────────────────────────────────────────────┘
```

---

# 5. Active Suppression Guard & Data Validation

To protect sender reputation and adhere to compliance standards (GDPR, CAN-SPAM, CASL), the system enforces automated validation boundaries:

1. **Active Suppression Guard (`loadSuppressionRulesFromSettings_`):**
   - Automatically checks prospect domains against blacklisted entities in `Lists & Settings`.
   - Filters out generic role-based inboxes (`info@`, `sales@`, `support@`, `admin@`, `billing@`, `contact@`, `help@`, `jobs@`).
   - Automatically tags violators as `Rejected (Suppression Guard)` and prevents downstream draft generation.
2. **Normalized Cell Validations:**
   - Enforces strict dropdown values in Google Sheets via `SheetSetup.gs` across all 12 tabs.
   - Company Fit Ratings strictly constrained to: `High`, `Medium`, `Low`, `Disqualified - Direct Competitor`.
   - Lead Status strictly constrained to: `Pending Review`, `Approved`, `Edit Required`, `Rejected`.

---

# 6. Pilot Sandbox Runner & Verification System

The platform includes an automated simulation and testing suite (`PilotSandboxRunner.gs`):

- **End-to-End Simulation:** Generates mock accounts, executes suppression screenings, qualifies leads, writes structured rows to `Persona Mapper Settings`, updates `Campaign Configuration`, and constructs actual drafts in Gmail Drafts.
- **Zero-Pollution Standby Cleanup:** Includes a 1-click cleanup routine (`cleanupPilotDataFromRealSheets`) that purges mock pilot records across all 12 operational sheets without altering configuration rules or dropdown validations.
- **Audit Verification:** The menu action **"⚙️ Audit Setup Status"** (`showSetupStatus`) performs real-time validation of tab structures, Script Properties, model connectivity, and header formatting.

---

# 7. Menu Navigation & Operational Controls

The platform features a consolidated Google Sheets menu bar:

```
🚀 AI Prospecting Studio
├── 🎛️ Open Control Panel               -> Launches ControlPanel.html Sidebar
├── 🔄 Sync Active Campaign             -> Synchronizes Campaign Config across sheets
├── 🎨 Fix Tab Header Contrast & Styling -> Applies High-Contrast Navy/Blue UI styling
├── ──────────────
├── 🏢 Run Agent 1: Company Discovery   -> Executes CompanyDiscovery.gs for active criteria
├── 👥 Run Agent 2: Persona Mapper       -> Executes PersonaMapper.gs for qualified accounts
├── ──────────────
├── ✅ Review Queue: Approve All Staged -> Batch marks pending leads as Approved
├── ✉️ Safe Dispatch: Create Gmail Drafts-> Converts approved leads into Gmail drafts
├── ──────────────
├── 🧪 Run End-to-End Pilot Test        -> Runs comprehensive 12-tab pipeline simulation
├── 🧹 Clean Up Pilot Data              -> Cleans all mock pilot records across tabs
└── ⚙️ Audit Setup Status               -> Displays system diagnostics & API connectivity
```

---

# 8. Security & Credential Storage

The application enforces zero-leakage security boundaries:
- **No Hardcoded Keys:** API credentials for Google Gemini, OpenAI, and email verification services must never be written into spreadsheet cells.
- **Storage Hierarchy:**
  1. `Script Properties` (Primary: Configured via Google Apps Script Project Settings)
  2. `User Properties` (Secondary)
  3. `Agent Settings` Sheet (Fallback for testing, protected range)
- **Least-Privilege Execution:** Cold email dispatch defaults to **"Workspace Gmail (Drafts Only)"**, requiring explicit user confirmation before any live outbound communication is transmitted.

---

# 9. Technical Stack & Repository Structure

- **Repository:** `https://github.com/growthtarget/ai-prospecting-studio`
- **Core Script Files:**
  - `Code.js` / `Code.gs`: Master menu controllers, sidebar launcher, and entry-point bindings.
  - `SheetSetup.js` / `SheetSetup.gs`: 12-tab schema builder, dynamic validation rules, contrast styler.
  - `Config.js` / `Config.gs`: Dynamic 38-parameter loader and dynamic Column A index scanner.
  - `AIProvider.js` / `AIProvider.gs`: Dual-engine API caller (Gemini 3.8 Flash + OpenAI 4o) and response parsers.
  - `CompanyDiscovery.js` / `CompanyDiscovery.gs`: Agent 1 discovery engine with search grounding.
  - `PersonaMapper.js` / `PersonaMapper.gs`: Agent 2 persona identification, icebreaker synthesis, and suppression guard.
  - `EmailDispatcher.js` / `EmailDispatcher.gs`: Multi-channel draft creator and webhook dispatcher.
  - `PilotSandboxRunner.js` / `PilotSandboxRunner.gs`: End-to-end sandbox runner and cleanup engine.
  - `ControlPanel.html`: Modern responsive user interface for campaign creation, department targeting, and pipeline monitoring.

---

# 10. Operational Playbook for Non-Technical Users

1. **Initialization:**
   - From the top menu, click **🚀 AI Prospecting Studio > 🎨 Fix Tab Header Contrast & Styling** to establish high-contrast headers.
   - Click **⚙️ Audit Setup Status** to confirm that API keys and all 12 sheets are online.
2. **Campaign Creation:**
   - Click **🎛️ Open Control Panel** to open the sidebar.
   - Enter your Campaign Name, Target Industry, Geography, Employee Sizes, and select target departments (e.g., ESG, Procurement, IT).
   - Enter your **Custom Business Needs / Buying Signals** and configure your Sender Identity.
   - Click **Save Campaign & Sync**.
3. **Execution:**
   - Run **🏢 Run Agent 1: Company Discovery** from the menu or sidebar to find qualified accounts.
   - Review findings in `Company Intelligence`.
   - Run **👥 Run Agent 2: Persona Mapper** to locate decision-makers and generate personalized copy.
4. **Triage & Draft Creation:**
   - Open `Prospect Review Queue`. Inspect icebreakers and value propositions.
   - Set status to **Approve** (or use **✅ Review Queue: Approve All Staged**).
   - Click **✉️ Safe Dispatch: Create Gmail Drafts** to populate your Gmail drafts folder for review and dispatch.
```