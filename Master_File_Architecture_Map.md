# AI Prospecting Studio + Persona Mapper (v2.0) — Master File & Architecture Map

## 1. Core Navigation, UI & State Management

### `Code.js` (Core & Navigation Engine)
* **Role:** Manages the Google Sheets top menu (`onOpen()`), active campaign tracking, and the Native View Filter Engine (`filterOperationalSheetsByCampaign_`) that isolates operational tabs per campaign (including `API_Usage_Log`).

### `SheetSetup.js` (Interface & Validation Formatter)
* **Role:** Formats all 13 tabs with high-contrast styling (`#174ea6` on `#e8f0fe`), locks the 38 Campaign Configuration parameters, pre-populates Agent Settings with budget guardrail rows (`DAILY_SPEND_CAP_USD`, `MIN_VERIFIER_CREDITS`), and populates native dropdowns across the workspace.

### `ControlPanel.html` (Sidebar UI)
* **Role:** The user interface where you configure campaign parameters, select departments, trigger agents, and save configurations via `studioDirectSaveCampaignPipeline()`.

---

## 2. AI Engine & Multi-Provider Routing Layer

### `AIProvider.js` (Universal AI Router)
* **Role:** Acts as the universal gateway. Dispatches requests dynamically to Gemini or OpenAI based on campaign/agent configuration with structured JSON extraction and citation collection.

### `GeminiProvider.js` (Google Gemini Engine)
* **Role:** Connects directly to the Google Generative Language API (`v1beta`) using `gemini-3.8-flash`. Automatically injects real-time Google Search Grounding (`google_search: {}`), extracts web chunk citations, checks pre-flight budget hard caps, and logs token usage/costs via `ApiTracker.js`.

### `OpenAIProvider.js` (OpenAI Fallback Engine)
* **Role:** Handles OpenAI requests (`requestOpenAIResponse_` and direct chat completions) using securely retrieved API keys from Script Properties, with integrated pre-flight budget checks and token usage logging via `ApiTracker.js`.

---

## 3. Financial Guardrails & API Tracking Layer

### `ApiTracker.js` (Centralized Usage & Cost Engine) — *NEW*
* **Role:** Records every external API request (LLMs and email verifiers) to Tab 2 (`API_Usage_Log`). Automatically calculates estimated dollar costs based on model token rates or verification credit pricing, enforces pre-flight daily spend hard caps (`checkApiBudgetGuardrails_`), and dispatches automated low-balance email alerts via `MailApp`.

---

## 4. Research, Schema & Grounding Parsers

### `ResearchSchema.js` (Strict Data Contracts)
* **Role:** Defines strict JSON schemas for structured LLM outputs across Gemini and OpenAI, enforcing required company firmographics, persona signals, and citation evidence contracts.

### `ResearchResults.js` (Grounding & Citation Parser)
* **Role:** Transforms raw AI responses and live search grounding metadata into canonical spreadsheet rows matching Tab 9 (Company Database - 10 cols) and Tab 10 (Research Evidence - 9 cols) with anti-hallucination quality downgrade guardrails.

### `ResearchStorage.js` (Atomic Storage & Deduplication)
* **Role:** Manages atomic persistence under `ScriptLock` and Scenario B Campaign-scoped deduplication (allowing the same company to be targeted under separate campaigns without collision).

### `ResearchChecks.js` (Offline Test Harness)
* **Role:** Executes offline unit test suites and grounding verification assertions across Gemini, OpenAI, and canonical table schemas without altering live spreadsheet data.

---

## 5. Operational Pipeline & Agents

### `CompanyDiscovery.js` (Agent 1: Account Discovery)
* **Role:** Executes AI-powered company discovery using Gemini/OpenAI and live Google Search Grounding. Captures account evidence, populates Tab 9 (Company Database) and Tab 10 (Research Evidence), and applies Scenario B multi-campaign deduplication.

### `PersonaMapper.js` (Agent 2: Decision-Maker Persona & Email Mapping)
* **Role:** Takes discovered companies, identifies target decision-makers, generates email patterns, checks the active suppression guard (`Lists & Settings`), runs email verifications (logging verifier credit usage to `API_Usage_Log`), captures persona initiative signals into Tab 10, and populates Tab 7 (Prospect Review Queue).

### `EmailSender.js` (Outreach Dispatch & Sequencer Integration)
* **Role:** Resolves sender profiles, stamps custom display names (`{ name: senderName }`), creates authentic Gmail drafts (`GmailApp.createDraft()`), or dispatches to Smartlead, Instantly, and Custom Webhooks.

### `PilotSandboxRunner.js` (End-to-End Simulator)
* **Role:** Runs an authentic simulation across all 13 tabs for testing and demonstration purposes with non-destructive cleanup and sample API usage logging.

---

## Confirmed Workbook Tab Order (13 Total Tabs)
1. **Dashboard** (KPI Matrix & Performance Rollup)
2. **API_Usage_Log** (Centralized API consumption, token counts, and cost tracking)
3. **Campaigns** (Master 28-column storage)
4. **Company Intelligence** (25-column account profiles)
5. **Persona Leads** (29-column decision-maker records)
6. **Persona Mapper Settings** (39 canonical parameters)
7. **Prospect Review Queue** (45-column outreach staging)
8. **Campaign Configuration** (Locked 38 vertical parameters)
9. **Company Database** (10 columns, Col J = Campaign ID)
10. **Research Evidence** (9 canonical citation columns)
11. **Agent Settings** (4 columns: Category, Key, Value, Status)
12. **Execution Logs** (Timestamped pipeline audits)
13. **Lists & Settings** (Validation arrays & global suppression rules)

---

## System Architecture Data Flow

```text
[User / Control Panel / Menu Trigger]
                 │
                 ▼
    ┌───────────────────────┐
    │      AIProvider.js    │ ──► Universal Router: dispatches requests 
    └───────────────────────┘     to Gemini or OpenAI dynamically.
                 │
       ┌────────┴────────┐
       ▼                 ▼
┌───────────────┐ ┌─────────────────┐
│GeminiProvider │ │OpenAIProvider.js│ ──► Handles direct REST API calls, API keys,
└───────────────┘ └─────────────────┘     and live Google Search Grounding (Gemini).
       │                 │
       └────────┬────────┘
                │ (API Request / Response & Token Usage)
                ▼
    ┌───────────────────────┐
    │    ApiTracker.js      │ ──► Pre-flight budget hard-cap checks, cost calculation,
    └───────────────────────┘     and atomic logging to Tab 2 (API_Usage_Log).
                │
                ▼
    ┌───────────────────────┐
    │   ResearchSchema.js   │ ──► Enforces strict JSON contracts & schemas 
    └───────────────────────┘     for account & persona research.
                │
                ▼
    ┌───────────────────────┐
    │  ResearchResults.js   │ ──► Transforms AI text into canonical rows:
    └───────────────────────┘     • Tab 9: Company Database (10 cols)
                │                 • Tab 10: Research Evidence (9 cols)
                │                   (Includes Anti-Hallucination Guardrails)
                ▼
    ┌───────────────────────┐
    │  ResearchStorage.js   │ ──► Handles atomic saving under ScriptLock &
    └───────────────────────┘     Scenario B Campaign-Scoped Deduplication.