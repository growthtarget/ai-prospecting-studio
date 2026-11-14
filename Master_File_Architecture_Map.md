# AI Prospecting Studio + Persona Mapper (v2.0) — Master File & Architecture Map

## 1. Core Navigation, UI & State Management

### `Code.js` (Core & Navigation Engine)
* **Role:** Manages the Google Sheets top menu (`onOpen()`), active campaign tracking, and the Native View Filter Engine (`filterOperationalSheetsByCampaign_`) that isolates operational tabs per campaign.

### `SheetSetup.js` (Interface & Validation Formatter)
* **Role:** Formats all 12 tabs with high-contrast styling (`#174ea6` on `#e8f0fe`), locks the 38 Campaign Configuration parameters, and populates native dropdowns across the workspace.

### `ControlPanel.html` (Sidebar UI)
* **Role:** The user interface where you configure campaign parameters, select departments, trigger agents, and save configurations via `studioDirectSaveCampaignPipeline()`.

---

## 2. AI Engine & Multi-Provider Routing Layer

### `AIProvider.js` (Universal AI Router)
* **Role:** Acts as the universal gateway. Dispatches requests dynamically to Gemini or OpenAI based on campaign/agent configuration with structured JSON extraction and citation collection.

### `GeminiProvider.js` (Google Gemini Engine)
* **Role:** Connects directly to the Google Generative Language API (`v1beta`) using `gemini-3.8-flash`. Automatically injects real-time Google Search Grounding (`google_search: {}`) and extracts web chunk citations.

### `OpenAIProvider.js` (OpenAI Fallback Engine)
* **Role:** Handles OpenAI requests (`requestOpenAIResponse_` and direct chat completions) using securely retrieved API keys from Script Properties.

---

## 3. Research, Schema & Grounding Parsers

### `ResearchSchema.js` (Strict Data Contracts)
* **Role:** Defines strict JSON schemas for structured LLM outputs across Gemini and OpenAI, enforcing required company firmographics, persona signals, and citation evidence contracts.

### `ResearchResults.js` (Grounding & Citation Parser)
* **Role:** Transforms raw AI responses and live search grounding metadata into canonical spreadsheet rows matching Tab 8 (Company Database - 10 cols) and Tab 9 (Research Evidence - 9 cols) with anti-hallucination quality downgrade guardrails.

### `ResearchStorage.js` (Atomic Storage & Deduplication)
* **Role:** Manages atomic persistence under `ScriptLock` and Scenario B Campaign-scoped deduplication (allowing the same company to be targeted under separate campaigns without collision).

### `ResearchChecks.js` (Offline Test Harness)
* **Role:** Executes offline unit test suites and grounding verification assertions across Gemini, OpenAI, and canonical table schemas without altering live spreadsheet data.

---

## 4. Operational Pipeline & Agents

### `CompanyDiscovery.js` (Agent 1: Account Discovery)
* **Role:** Executes AI-powered company discovery using Gemini/OpenAI and live Google Search Grounding. Captures account evidence, populates Tab 8 (Company Database) and Tab 9 (Research Evidence), and applies Scenario B multi-campaign deduplication.

### `PersonaMapper.js` (Agent 2: Decision-Maker Persona & Email Mapping)
* **Role:** Takes discovered companies, identifies target decision-makers, generates email patterns, checks the active suppression guard (`Lists & Settings`), runs email verifications, captures persona initiative signals into Tab 9, and populates Tab 6 (Prospect Review Queue).

### `EmailSender.js` (Outreach Dispatch & Sequencer Integration)
* **Role:** Resolves sender profiles, stamps custom display names (`{ name: senderName }`), creates authentic Gmail drafts (`GmailApp.createDraft()`), or dispatches to Smartlead, Instantly, and Custom Webhooks.

### `PilotSandboxRunner.js` (End-to-End Simulator)
* **Role:** Runs an authentic simulation across all 12 tabs for testing and demonstration purposes with non-destructive cleanup.

---

## Detailed File Correlation & Agent Integration

| File Name | Primary Role in System | How it Connects to Agents / Pipeline |
| :--- | :--- | :--- |
| **`AIProvider.js`** | Universal AI Gateway | Acts as the single entry point for Agent 1 and Agent 2. It reads your active campaign settings (`GEMINI` or `OPENAI`), routes the prompt, and extracts clean text, structured JSON, and web citations. |
| **`GeminiProvider.js`** | Google Gemini Engine (`v1beta`) | Directly powers `gemini-3.8-flash`. Automatically appends `tools: [{ google_search: {} }]` for real-time web search grounding and captures `groundingMetadata` (URLs and queries). |
| **`OpenAIProvider.js`** | OpenAI Fallback Engine (`GPT-4o`) | Powers fallback execution via standard chat completions or responses API using securely retrieved API keys from Script Properties. |
| **`ResearchSchema.js`** | Strict Data Contracts | Defines the exact JSON schemas required by Agent 1 (Company firmographics & website evidence) and Agent 2 (Decision-maker titles, buying signals, and persona evidence). |
| **`ResearchResults.js`** | Data Normalizer & Grounding Parser | Takes raw AI JSON outputs and extracts live web search URLs. Features an Anti-Hallucination Quality Guardrail that downgrades verification scores to `'Unknown'` if an ungrounded URL is injected. Maps output directly into Tab 8 and Tab 9. |
| **`ResearchStorage.js`** | Atomic Save & Deduplication Engine | Executes Scenario B Multi-Campaign Deduplication (allowing the same account to be targeted under separate campaigns without collision, while blocking duplicates within the same campaign). Saves atomically using `LockService`. |
| **`ResearchChecks.js`** | Automated Test Harness | Runs internal assertions (23+ checks) on schema mapping, column length validation, and deduplication logic without altering live spreadsheet data. |

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
                │ (Raw AI Response & Grounding Metadata)
                ▼
    ┌───────────────────────┐
    │   ResearchSchema.js   │ ──► Enforces strict JSON contracts & schemas 
    └───────────────────────┘     for account & persona research.
                │
                ▼
    ┌───────────────────────┐
    │  ResearchResults.js   │ ──► Transforms AI text into canonical rows:
    └───────────────────────┘     • Tab 8: Company Database (10 cols)
                │                 • Tab 9: Research Evidence (9 cols)
                │                   (Includes Anti-Hallucination Guardrails)
                ▼
    ┌───────────────────────┐
    │  ResearchStorage.js   │ ──► Handles atomic saving under ScriptLock &
    └───────────────────────┘     Scenario B Campaign-Scoped Deduplication.
                │
                ▼
    ┌───────────────────────┐
    │   ResearchChecks.js   │ ──► Offline unit test harness verifying 
    └───────────────────────┘     schemas, grounding, and deduplication.