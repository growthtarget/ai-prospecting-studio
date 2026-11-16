# AI Prospecting Studio + Persona Mapper (v2.0) — Master File & Architecture Map

## 1. Core Navigation, UI & State Management

### `Code.js` (Core, Multi-Client Isolation & Native Filter Engine)
* **Role:** Manages the Google Sheets top menu (`onOpen()`), `onEdit()` bi-directional sync (Dashboard status/ID & `Lists & Settings` auto-sorting), active campaign state (`getStudioActiveCampaignId_()`), and the complete **Multi-Client Isolation & Native View Filter Engine**.
* **Client Credential Isolation:** Implements `resolveClientScopedSecret_()`, `resolveCampaignClientName_()`, `isLocalAgencyClient_()`, `normalizeClientPropertyPrefix_()`, and `isStudioValidSecretValue_()` to enforce zero-fallback `CLIENT_{NAME}_*` Script Property isolation for external clients while restricting `"Growth Target"` to local default keys.
* **Client Workspace Management:** Implements `addClientAccountToSettings()`, `getRegisteredClientsList_()`, `getInactiveClientsList_()`, and `filterAgentSettingsForClientWorkspace()` (allowing instant workspace filtering and auto-suggested IDs even for 0-campaign clients).
* **Multi-Tab Native View Filter Engine:**
  * `filterOperationalSheetsByCampaign_()`: Auto-filters `Prospect Review Queue`, `Persona Leads`, `Company Intelligence`, `Company Database`, `Research Evidence`, `Execution Logs`, `API_Usage_Log`, and `Persona Mapper Settings` by the Active **Campaign ID**.
  * `filterCampaignsSheetByClient_()` & `filterDashboardRollupByClient_()`: Auto-filters `Campaigns` (Col C) and `Dashboard` → `Campaign Performance Rollup` (Row 17+, Col C) by the Selected **Client / Owner**.
* **Campaign State Sync:** `selectAndActivateCampaign()` and `studioDirectSaveCampaignPipeline()` sync the locked 38-row `Campaign Configuration`, 28-column `Campaigns` (storing extended fields like `senderEmail` and `replyToEmail` in `Campaign Config JSON`), 39-column `Persona Mapper Settings`, and 4-column `Dashboard` KPIs (with client-scoped AI key status).

### `SheetSetup.js` (Interface, Registry & Validation Formatter)
* **Role:** Formats all 13 tabs with high-contrast styling (`#174ea6` on `#e8f0fe`), freezes header rows, preserves the locked 38 `Campaign Configuration` parameters and 28 `Campaigns` columns, and enforces plain number formats (`0` / `0.00`) on numeric limits, scores, and caps to prevent 1900 date conversions.
* **Tab 13 (`Lists & Settings`) Registry Engine:**
  * `sortListsAndSettingsTab_()` pins all `Client - Account` rows at the top (`Growth Target` #1, Active clients A–Z, Inactive archived clients A–Z), followed by Whitelist and Suppression rules.
  * `refreshClientDropdownValidations_()` dynamically binds active `Client - Account` entries to `Campaign Configuration` (Row 4) and `Campaigns` (Col C).
  * `getInactiveClientsFromSettings_()` enables true client archival without deleting historical `Campaigns` rows.
* **Tab 11 (`Agent Settings`) Security Vault & Auto-Filter:** Pre-populates `Growth Target` local defaults and auto-generates persistent `"Client: {Name}"` blocks for every active external client—displaying `"Configured in Script Properties (Protected)"` status indicators, editable per-client budget guardrails (`CLIENT_{NAME}_DAILY_SPEND_CAP_USD`, `CLIENT_{NAME}_MIN_VERIFIER_CREDITS`), and filtering via `filterAgentSettingsByClient_()`.

### `ControlPanel.html` (Client-First Sidebar UI)
* **Role:** Interactive Google Sheets sidebar (`380px`) driving campaign creation, configuration, agent triggers, and client workspace switching.
* **Client-First Cascading Workspace:** Placing `CLIENT / OWNER` at the top acts as the primary workspace switcher—filtering `ACTIVE CAMPAIGN` strictly to the selected client's campaigns (preventing accidental cross-client overwrites), excluding `Inactive` archived clients, and supporting inline `[ + Add New Client... ]` creation.
* **0-Campaign Instant Workspace & Smart UI Controls:** Auto-filters `Agent Settings`, `Campaigns`, and `Dashboard` Rollup when switching to a 0-campaign client and auto-suggests unique IDs/names (e.g., `PA-001` / `Pilot Acme - Campaign 1`); includes sticky pain-point presets, AI Provider–filtered model dropdowns, busy concurrency locks, and dedicated `Sender Email` / `Reply-To Email` fields saved in `Campaign Config JSON`.

### `Config.js` (Canonical Schema & App Definitions)
* **Role:** Holds `APP_CONFIG` (`sheetDefinitions` and canonical tab headers/configuration constants) used across setup and validation workflows.

---

## 2. AI Engine & Multi-Provider Routing Layer

### `AIProvider.js` (Universal AI Router)
* **Role:** Acts as the universal gateway. Dispatches Agent 1 and Agent 2 requests dynamically to Gemini or OpenAI using campaign-level `clientName` and `campaignId` context, with structured JSON extraction and citation collection.

### `GeminiProvider.js` (Google Gemini Engine & Resiliency Cascade)
* **Role:** Connects to the Google Generative Language API (`v1beta`) using client-scoped `GEMINI_API_KEY` (`CLIENT_{NAME}_GEMINI_API_KEY` for external clients, default key for `Growth Target`).
* **Smart Resiliency & Grounding:** Automatically injects live Google Search Grounding (`google_search: {}`), extracts web chunk citations, includes a model cascade (`gemini-3.8-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-flash`), handles 429 grounding fallback to ungrounded JSON, applies `thinkingConfig` speedups and thought filtering, checks pre-flight budget caps, and logs token usage/costs via `ApiTracker.js`.

### `OpenAIProvider.js` (OpenAI Fallback Engine)
* **Role:** Handles OpenAI requests via the universal `/v1/responses` API and guarded Chat Completions fallback using strictly client-scoped `OPENAI_API_KEY` (`CLIENT_{NAME}_OPENAI_API_KEY` for external clients), with integrated pre-flight budget checks and token/cost logging via `ApiTracker.js`.

---

## 3. Financial Guardrails & API Tracking Layer

### `ApiTracker.js` (Centralized Usage, Client Tagging & Cost Engine)
* **Role:** Atomically records every external API request (Gemini, OpenAI, and Email Verifiers) to Tab 2 (`API_Usage_Log`), tagging `Details / Error` with `[Client: {ClientName}]`.
* **Per-Client Budget Guardrails:** Automatically calculates estimated USD costs, enforces pre-flight per-client daily spend hard caps (`CLIENT_{NAME}_DAILY_SPEND_CAP_USD` vs. local `DAILY_SPEND_CAP_USD` via `checkApiBudgetGuardrails_()`), and dispatches automated low-verifier-credit alerts (`CLIENT_{NAME}_MIN_VERIFIER_CREDITS`) via `MailApp`.

---

## 4. Research, Schema & Grounding Parsers

### `ResearchSchema.js` (Strict Data Contracts)
* **Role:** Defines strict JSON schemas for structured LLM outputs across Gemini and OpenAI, enforcing required company firmographics, persona signals, and citation evidence contracts.

### `ResearchResults.js` (Grounding & Citation Parser)
* **Role:** Transforms raw AI responses and live search grounding metadata into canonical spreadsheet rows matching Tab 9 (`Company Database` - 10 cols) and Tab 10 (`Research Evidence` - 9 cols) with anti-hallucination quality downgrade guardrails.

### `ResearchStorage.js` (Atomic Storage & Deduplication)
* **Role:** Manages atomic persistence under `ScriptLock` and Scenario B Campaign-scoped deduplication (allowing the same company to be targeted under separate campaigns without collision).

### `ResearchChecks.js` (Offline Test Harness)
* **Role:** Executes offline unit test suites and grounding verification assertions across Gemini, OpenAI, and canonical table schemas without altering live spreadsheet data.

---

## 5. Operational Pipeline, Agents & Client Outreach Relay

### `CompanyDiscovery.js` (Agent 1: Account Discovery)
* **Role:** Executes AI-powered company discovery scoped to `clientName` and `campaignId` using Gemini/OpenAI and live Google Search Grounding. Inspects campaign focus/keywords/industry for ESG detection (decoupled from client/owner name), captures account evidence into Tab 10 (`Research Evidence`), populates Tab 9 (`Company Database` Col J = `Campaign ID`), and applies Scenario B multi-campaign deduplication.

### `PersonaMapper.js` (Agent 2: Decision-Maker Persona & Email Mapping)
* **Role:** Reads prior `Research Evidence` clues, identifies decision-makers using client-scoped AI keys, generates/validates email patterns, enforces the active suppression guard (`Lists & Settings`), verifies emails via `ZeroBounce` / `MillionVerifier` / `NeverBounce` / `Custom HTTPS` (strict fail-closed enforcement for external clients when verifier != `None`; safe unverified pass-through when local agency keys are unconfigured), logs verifier credits to `API_Usage_Log`, allows `catch-all` / `risky` pass-through with soft warnings, captures persona signals into Tab 10, stages leads in Tab 7 (`Prospect Review Queue`), and syncs live `Dashboard` KPIs.

### `EmailSender.js` (Client-Isolated Outreach Dispatch Engine)
* **Role:** Dispatches approved/queued outreach strictly isolated by Client / Owner with custom sender identity (`senderName`, `senderEmail`, `replyToEmail`):
  * **Local Agency (`Growth Target`):** Uses native Google Workspace `GmailApp` (`Drafts Only` and `Direct Send`) with alias/reply-to support, plus local `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`, or webhook endpoints.
  * **External Clients — Option A (Workspace Gmail Relay):** Routes `Drafts Only` and `Direct Send` strictly to the client's deployed `ClientGmailRelay.gs` Web App via `CLIENT_{NAME}_GMAIL_RELAY_URL` and `CLIENT_{NAME}_GMAIL_RELAY_SECRET`. **Fail-closed:** external client drafts/sends never fall back to the agency operator's Gmail mailbox.
  * **External Clients — Option B (Sequencers & Webhooks):** Enforces strict `CLIENT_{NAME}_SMARTLEAD_API_KEY`, `CLIENT_{NAME}_INSTANTLY_API_KEY`, and `CLIENT_{NAME}_SENDER_WEBHOOK_URL` keys with zero local fallback.

### `ClientGmailRelay.gs` (Standalone External Client Onboarding Template — v2.1)
* **Role:** Deployed as a standalone Google Apps Script Web App inside an external client's own Google Workspace account so outreach drafts (`createDraft`) and sends (`sendEmail`) execute inside the **client's own Gmail mailbox and domain** without granting agency access to their full Workspace.
* **v2.1 Capabilities:** Authenticates requests against the client's `RELAY_SECRET` Script Property (`CLIENT_{NAME}_GMAIL_RELAY_SECRET`), honors `senderName`, `senderEmail` (verifying `GmailApp.getAliases()` with safe primary mailbox fallback), `replyTo` (`replyToEmail`), `campaignId`, and `clientName`.

### `PilotSandboxRunner.js` (Client-Isolated Simulator & Option A Full Wipe)
* **Role:** Runs an authentic end-to-end simulation (`PILOT-CMP-001` under client `"Pilot Acme"`) across all 13 tabs, optionally tests client Gmail relay routing, logs sample API/verifier usage tagged `[Client: Pilot Acme]`, and appends/updates `Dashboard` Rollup non-destructively.
* **Option A Full Wipe (`cleanupPilotDataFromRealSheets()`):** Removes zero-trace of all `PILOT-*` and `Pilot Acme` test data across all 13 tabs (including `Dashboard` Rollup rows, `Agent Settings` client block, `Lists & Settings` client/suppression entries, `API_Usage_Log`, and `Execution Logs`), refreshes sheet validations, and cleanly restores the active real campaign (or resets to `Standby` if none remain).

---

## Multi-Client / Owner Isolation Architecture

| Operating Tier | Example Owner | Credential Namespace (Script Properties) | Email Dispatch Isolation | Local Fallback Allowed? |
| :--- | :--- | :--- | :--- | :--- |
| **Local / Agency Brand** | `Growth Target` | Default keys (`GEMINI_API_KEY`, `OPENAI_API_KEY`, `VERIFIER_API_KEY`, `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`) — explicitly ignores `CLIENT_*` | Native Workspace `GmailApp` (`Drafts Only` / `Direct Send`) or local sequencer keys | N/A (Local Scope) |
| **External Client** | `eValuater`, `Pilot Acme` | Strictly prefixed keys (`CLIENT_{CLIENT_NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`, `_GMAIL_RELAY_URL`, `_GMAIL_RELAY_SECRET`, `_SMARTLEAD_API_KEY`, `_INSTANTLY_API_KEY`, `_SENDER_WEBHOOK_URL`) | **Option A:** Client's standalone `ClientGmailRelay.gs` Web App<br>**Option B:** Client-scoped Smartlead / Instantly / Webhook | **NO (Zero Fallback / Fail-Closed)** |

### Credential Security & Archival Rules
* **Server-Side Encryption Only:** Raw API keys and relay secrets live strictly in **Apps Script Script Properties** — never plain text in Google Sheets cells.
* **Placeholder Rejection:** `isStudioValidSecretValue_()` rejects placeholder strings (`PASTE_`, `Configured in Script Properties`, `Protected`) so status labels in `Agent Settings` are never transmitted as live API keys.
* **True Archival:** Setting a client's `Active status` to `Inactive` in Tab 13 (`Lists & Settings`) removes that client from the Control Panel dropdown and sheet validations while preserving historical campaign rows in Tab 3 (`Campaigns`).

---

## Workspace Gmail Relay Setup for Clients (`ClientGmailRelay.gs` v2.1)

1. **Client-Side Deployment:** Open Google Apps Script inside the **client's Google Workspace account**, paste `ClientGmailRelay.gs` (v2.1), set `RELAY_SECRET` in the client project's **Script Properties**, and deploy as a **Web App** (*Execute as: Me*, *Who has access: Anyone* — protected by `RELAY_SECRET` in the JSON POST body).
2. **Agency Master Binding:** In the Master AI Prospecting Studio Apps Script **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_URL` = Client Web App `/exec` URL
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_SECRET` = Matching `RELAY_SECRET`
3. **Campaign Identity Binding:** In the Control Panel sidebar, configure `Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` (mailbox or Workspace Send-As alias), and `Reply-To Email` (stored inside `Campaigns` Col AA `Campaign Config JSON`).
4. **Verification:** Tab 11 (`Agent Settings`) automatically displays `Configured in Script Properties (Protected)` for `CLIENT_{NAME}_GMAIL_RELAY_URL` once both `_URL` and `_SECRET` are present.

---

## Confirmed Workbook Tab Order & Schema (13 Total Tabs)

1. **`Dashboard`** *(Cols A–G)*
   * **Rows 1–2:** Green Banner (`A1:D1`), Active Campaign Name (`C1`), Status (`D1`), Active Campaign ID (`B2`).
   * **Rows 3–13:** 10-Row KPI Matrix (`A3:D13`) with Operating Department (`Client Owner` / `Local Owner`), live pipeline metrics, and client-scoped AI Key connection status.
   * **Rows 16–17+:** `Campaign Performance Rollup` (`A17:G...`) **auto-filtered on Col C (`Department`) by Selected Client / Owner**.
2. **`API_Usage_Log`** *(10 Canonical Cols — Auto-Filtered by Active Campaign ID)*
   * `Timestamp`, `Campaign ID`, `Service Name`, `Operation`, `HTTP Status`, `Input Tokens / Units`, `Output Tokens`, `Est. Cost ($)`, `Status`, `Details / Error` (tagged `[Client: {Name}]`).
3. **`Campaigns`** *(Locked 28 Cols — Auto-Filtered by Client / Owner Col C)*
   * Master campaign registry. Col C = `Department` (dynamic Client dropdown from `Lists & Settings`), Col D = `Status`, Cols H & J–M formatted as plain integer (`0`), Col AA = `Campaign Config JSON` (extended attributes including `senderEmail`, `replyToEmail`, `presetNeed`).
4. **`Company Intelligence`** *(25 Cols — Auto-Filtered by Active Campaign ID)*
   * Account enrichment profiles. Col J = 6 allowed `ESG Reporting Maturity` states; Col Q = `eValuater Solution Fit` (`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`).
5. **`Persona Leads`** *(29 Cols — Auto-Filtered by Active Campaign ID)*
   * Decision-maker profiles and email intelligence. Col Y = `Email Verification Status` (`valid`, `risky`, `invalid`, `catch-all`, `unverified`).
6. **`Persona Mapper Settings`** *(Locked 39 Canonical Cols — Auto-Filtered by Active Campaign ID Col A)*
   * Agent 2 parameters & scoring weights (Row 2+ direct write; Cols I, Z–AB, AI formatted as `0`; weight cols formatted as `0%`).
7. **`Prospect Review Queue`** *(45 Cols — Auto-Filtered by Active Campaign ID)*
   * Human-in-the-loop outreach staging. Col R `Review Status`, Col V `Human Approval Status`, Col Y `Send Status` (`Queued`, `Draft Created in Gmail`, `Draft Created`, `Sent`, `Pushed to Sequencer`, `Suppressed - Not Sent`, `Draft Failed`).
8. **`Campaign Configuration`** *(Locked 38 Vertical Parameters, Rows 2–39)*
   * Active campaign configuration key-value view. Row 2 = `CAMPAIGN_ID`, Row 4 = `DEPARTMENT` (dynamic Client dropdown), plain integer format (`0`) on numeric parameters (`B8`, `B10:B13`).
9. **`Company Database`** *(10 Cols — Auto-Filtered by Active Campaign ID Col J)*
   * Master discovered account directory (Cols A–I standard firmographics + Col J = `Campaign ID` for Scenario B multi-campaign deduplication).
10. **`Research Evidence`** *(9 Canonical Cols — Auto-Filtered by Active Campaign ID)*
    * Grounded audit trail (`Evidence ID`, `Campaign ID`, `Entity Name`, `Entity Type` (`Company` / `Persona`), `Search Query Used`, `Grounding Evidence URL`, `Extracted Snippet / Signal`, `Verification Score`, `Date Captured`).
11. **`Agent Settings`** *(4 Cols — Auto-Filtered by Active Client / Owner Col A)*
    * `Category`, `Setting Key`, `Value`, `Status / Description`. Holds `Growth Target` defaults + auto-generated `"Client: {Name}"` blocks with protected credential badges and editable per-client daily spend caps / min verifier credits.
12. **`Execution Logs`** *(6 Cols — Auto-Filtered by Active Campaign ID)*
    * Timestamped pipeline execution audits (`Timestamp`, `Campaign ID`, `Action`, `Target Lead`, `Status`, `Details`).
13. **`Lists & Settings`** *(3 Cols: `List name`, `Value`, `Active status`)*
    * Client registry (`Client - Account` pinned cleanly at the top: `Growth Target` -> Active A–Z -> Inactive A–Z) + Whitelist and Suppression rules (`Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, `Suppression - Internal Exclusion`).

---

## System Architecture Data Flow

```text
[User / ControlPanel.html / Top Menu Trigger]
                 │  (Selects Client/Owner -> Cascading Campaign + Sender/Reply-To)
                 ▼
    ┌───────────────────────┐
    │        Code.js        │ ──► Resolves Client Scope (Growth Target vs. CLIENT_{NAME}_*)
    └───────────────────────┘     & Updates Native Auto-Filters across Tabs 1-12.
                 │
                 ▼
    ┌───────────────────────┐
    │  CompanyDiscovery.js  │ ──► Agent 1: Discovers target accounts + ESG/Need signals
    │   & PersonaMapper.js  │ ──► Agent 2: Maps personas, emails, suppression & verifiers
    └───────────────────────┘
                 │
                 ▼
    ┌───────────────────────┐
    │      AIProvider.js    │ ──► Universal Router: dispatches requests dynamically
    └───────────────────────┘     using strict client-scoped credentials (zero fallback).
                 │
       ┌────────┴────────┐
       ▼                 ▼
┌───────────────┐ ┌─────────────────┐
│GeminiProvider │ │OpenAIProvider.js│ ──► Live Google Search Grounding + Model Cascade
└───────────────┘ └─────────────────┘     or OpenAI /v1/responses + Chat fallback.
       │                 │
       └────────┬────────┘
                ▼
    ┌───────────────────────┐
    │     ApiTracker.js     │ ──► Pre-flight client spend hard-caps, cost calculation,
    └───────────────────────┘     & [Client: ...] logging to Tab 2 (API_Usage_Log).
                │
                ▼
    ┌───────────────────────┐
    │   ResearchSchema.js   │ ──► Enforces strict JSON contracts & schemas
    │  & ResearchResults.js │     + anti-hallucination quality downgrade guardrails.
    └───────────────────────┘
                │
                ▼
    ┌───────────────────────┐
    │   ResearchStorage.js  │ ──► Atomic ScriptLock writes + Scenario B dedup to:
    └───────────────────────┘     Tab 4, Tab 5, Tab 7, Tab 9 (Col J), Tab 10 (Evidence).
                │
                ▼  (After Human Review / Approval in Tab 7: Prospect Review Queue)
    ┌───────────────────────┐
    │     EmailSender.js    │ ──► Strict Client-Isolated Dispatch Routing:
    └───────────────────────┘     • Growth Target -> Local Workspace GmailApp / Local Sequencers
                                  • External Clients (Option A) -> ClientGmailRelay.gs (v2.1)
                                  • External Clients (Option B) -> CLIENT_{NAME}_* Sequencer/Webhook
                │
                ▼
    ┌───────────────────────┐
    │  Dashboard & Sheets   │ ──► Updates live KPIs, Client-Filtered Performance Rollup,
    └───────────────────────┘     Send Status, Execution Logs & Verifier Low-Balance Alerts.
```