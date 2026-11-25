# AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
## System Master Brief, Architecture & Technical Specification

**Project:** AI Prospecting Studio + Persona Mapper  
**Version:** 2.0 (Multi-Client Production + Email Deliverability Architecture)  
**Repository:** [https://github.com/growthtarget/ai-prospecting-studio](https://github.com/growthtarget/ai-prospecting-studio)  
**Primary AI Engine:** Google Gemini API (`gemini-3.8-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-flash` cascade with Google Search Grounding & ungrounded 429 fallback)  
**Fallback AI Engine:** OpenAI API (`gpt-4o` / `gpt-4o-mini` via universal `/v1/responses` API & guarded Chat Completions fallback)  
**Application Interface:** Google Sheets (15 Dedicated Synchronized Tabs)  
**Backend:** Google Apps Script (17 Modular Project Files + `appsscript.json` + standalone `ClientGmailRelay.gs` v2.2 onboarding template)  
**Operating Model:** Multi-Client Agency Isolation (`Growth Target` local brand + zero-fallback `CLIENT_{NAME}_*` external client workspaces)  
**Verification Providers:** ZeroBounce, MillionVerifier, NeverBounce, Custom HTTPS  
**Dispatch Engines:** Native Google Workspace Gmail (`Growth Target` Drafts & Direct Send), Client Google Workspace Gmail Relay (`ClientGmailRelay.gs` v2.2 — Option A), Native Microsoft 365 / Exchange Online via Microsoft Graph API v1.0 (`Drafts Only` & `Direct Send` — Zero External Platform), Smartlead, Instantly, Custom Webhook (Option B)  
**Deliverability Engine:** Centralized `EmailEvents` Data Layer + Dynamic `Email Deliverability` Dashboard & Charts (`EmailEventsDeliverability.js` — Inline `Sent` logging + Hourly Gmail / Client Relay / Microsoft 365 Graph polling)  

---

# 1. Executive Summary & System Vision

**AI Prospecting Studio + Persona Mapper (v2.0)** is an autonomous, multi-client B2B outbound intelligence, persona discovery, hyper-personalized outreach, and email-deliverability analytics platform built natively inside Google Sheets and powered by Google Apps Script.

Operating from a single agency Google Workspace spreadsheet, the platform manages multiple independent client accounts under strict credential, mailbox, budget, and spreadsheet-view isolation while unifying **Market Research (Agent 1: Company Discovery)**, **Decision-Maker Persona Mapping & Draft Synthesis (Agent 2: Persona Mapper)**, **Multi-Platform Outreach Dispatch**, and **Client-Scoped Deliverability & Bounce Telemetry** into one governed pipeline.

### Core Capabilities:
1. **Multi-Client / Owner Isolation & Client-First Workspace:** Manages the local agency brand (`Growth Target`) alongside external clients (e.g., `eValuater`) from one master workbook. External clients enforce strict `CLIENT_{NAME}_*` Script Properties with **zero fallback** to local keys or agency Gmail; clients marked `Inactive` in `Lists & Settings` are cleanly archived from sidebar and sheet dropdowns without deleting historical campaign records.
2. **Dynamic Dual-AI Routing & Resiliency Cascade:** Primary execution on Google Gemini (`gemini-3.8-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-flash`) with live Google Search Grounding, `thinkingConfig` speedups, and automatic 429 ungrounded JSON fallback—plus OpenAI `/v1/responses` support via `AIProvider.js`.
3. **Centralized API Cost Tracking & Per-Client Budget Guardrails:** `ApiTracker.js` logs every LLM and email verifier call to `API_Usage_Log` tagged with `[Client: {Name}]`, calculates estimated USD costs, enforces per-client daily spend hard caps, and dispatches low-balance alerts via `MailApp`.
4. **Active Suppression Guard & Email Verification:** Screens blacklisted domains, opt-out emails, competitor lists, and role-based accounts (`info@`, `admin@`, `support@`) via `Lists & Settings`. Verifies emails per client via ZeroBounce, MillionVerifier, NeverBounce, or Custom HTTPS (strict fail-closed enforcement for external clients when verifier != `None`; safe local unverified pass-through when agency keys are unconfigured; `catch-all`/`risky` soft-warning pass-through).
5. **Isolated Multi-Tab Schemas & Native Auto-Filter Engine:** Locks the 38 vertical parameters in `Campaign Configuration`, 28 columns in `Campaigns` (storing extended fields like `senderEmail` and `replyToEmail` in `Campaign Config JSON`), and 39 columns in `Persona Mapper Settings`. Native Google Sheets auto-filters isolate operational tabs (including `EmailEvents`) by Active Campaign ID and `Campaigns`, `Dashboard` Rollup, and `Agent Settings` by Selected Client / Owner.
6. **Multi-Platform Client Dispatch (Google Workspace Relay, Native Microsoft 365 / Exchange Online & Sequencers):**
   * **Google Workspace Clients (`ClientGmailRelay.gs` v2.2):** Creates drafts, sends emails (capturing `draftId`, `messageId`, `threadId`, and `rfcMessageId`), and polls bounces/replies/manual draft sends directly inside each external client's own Google Workspace mailbox via a standalone Web App relay with `senderEmail` alias fallback and `replyToEmail` support.
   * **Microsoft 365 / Exchange Online Clients (Native Microsoft Graph API v1.0):** Creates drafts (`POST /users/{mailbox}/messages` with `ImmutableId` -> `Draft Created`) or sends emails (`POST /messages/{id}/send` with `/sendMail` fallback -> `Sent`) directly in the client's Exchange Online mailbox via OAuth2 `client_credentials` using `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET` — requiring zero external middleware platforms (no Make, Zapier, n8n, or Power Automate needed).
   * **Sequencers & Webhooks:** Routes to client-scoped Smartlead, Instantly, or Custom HTTPS Webhooks.
7. **Centralized Email Performance & Deliverability Module (`EmailEventsDeliverability.js`):**
   * Records every lifecycle event (`Sent`, `Delivered`, `Bounced`, `Replied`) into the central 11-column **`EmailEvents`** sheet with deterministic `SHA-256` deduplication (`event_id`).
   * Captures `Thread / Message ID` during dispatch (`EmailSender.js`), logs `Sent` events inline, and runs an automated hourly sync (plus silent sidebar **`⚡ Sync Events`**) across Local Gmail, Client Gmail Relays (`v2.2`), and Microsoft 365 Graph mailboxes (with manual Draft-to-Sent promotion, SMTP DSN/NDR Hard vs. Soft bounce parsing, reply tracking, optional Exchange Message Trace v1.0, and 4-hour no-bounce delivery confirmation with late-bounce reconciliation).
   * Drives the **`Email Deliverability`** tab (Client & Campaign filter bar, Headline Tiles, Rate Breakdown, 3 tables, and 3 charts) and feeds Section 6 (`Email Deliverability Snapshot`) inside `ControlPanel.html` directly from the existing sidebar filters without restructuring how clients or campaigns are configured.
8. **Full-Pipeline Pilot Sandbox & Option A Full Wipe:** `PilotSandboxRunner.js` runs an end-to-end simulation (`PILOT-CMP-001` under `"Pilot Acme"`) across all 15 tabs (including 11 multi-day `EmailEvents` samples across Gmail & Microsoft 365) and includes a zero-trace Option A cleanup (`cleanupPilotDataFromRealSheets()`) that removes all pilot/client data and cleanly restores the active production workspace.

---

# 2. 15-Tab System Architecture & Master Schemas

The application operates across **15 tightly synchronized sheets** styled with high-contrast Dark Navy headers (`#174ea6`) on Light Blue (`#e8f0fe`), frozen header rows, plain integer/number formats (`0` / `0.00` / `@`) to prevent 1900 date conversion bugs, and native Google Sheets dropdown validations.

    ┌─────────────────────────────────────────────────────────────────────────────────────────────────┐
    │                                   15 SYNCHRONIZED WORKBOOK TABS                                 │
    ├─────────────────────────────────────────────────────────────────────────────────────────────────┤
    │  01. Dashboard                │ KPI Matrix (Cols A-D) & Client-Filtered Performance Rollup (17+) │
    │  02. EmailEvents (11 Cols)    │ Centralized Sent/Delivered/Bounced/Replied ledger (Camp Filter) │
    │  03. Email Deliverability     │ Client/Campaign filters, KPI tiles, rates, 3 tables & 3 charts  │
    │  04. API_Usage_Log (10 Cols)  │ Token/credit costs tagged [Client: ...] (Auto-filtered: Camp ID)│
    │  05. Campaigns (Locked 28)    │ Master campaign registry & JSON config (Auto-filtered: Owner)   │
    │  06. Company Intelligence(25) │ Enriched account profiles & scores (Auto-filtered: Camp ID)     │
    │  07. Persona Leads (29 Cols)  │ Decision-makers & email verification (Auto-filtered: Camp ID)   │
    │  08. Persona Mapper Settings  │ Locked 39 canonical Agent 2 settings (Auto-filtered: Camp ID)   │
    │  09. Prospect Review Queue(45)│ Human-in-the-loop triage & outreach (Auto-filtered: Camp ID)    │
    │  10. Campaign Configuration   │ Locked 38 vertical active campaign parameters (Rows 2-39)       │
    │  11. Company Database (10)    │ Master account directory w/ Col J Campaign ID (Filtered: Camp)  │
    │  12. Research Evidence (9)    │ Grounded URLs, snippets & scores (Auto-filtered: Camp ID)       │
    │  13. Agent Settings (4 Cols)  │ Local defaults & protected Client blocks (Filtered: Owner)      │
    │  14. Execution Logs (6 Cols)  │ Timestamped pipeline audits (Auto-filtered: Camp ID)            │
    │  15. Lists & Settings (3 Cols)│ Pinned Client - Account registry + Whitelist/Suppression rules  │
    └─────────────────────────────────────────────────────────────────────────────────────────────────┘

### Detailed Tab Specifications

#### 1. `Dashboard` *(Cols A–G)*
* **Rows 1–2:** Green Banner (`A1:D1`), Active Campaign Name (`C1`), Status (`D1`), and Active Campaign ID (`B2`) with bi-directional `onEdit()` sync to `Campaigns` and `Campaign Configuration`.
* **Rows 3–13:** 10-Row KPI Matrix (`A3:D13`) displaying Active Campaign Name, Campaign ID, Operating Department (`Client Owner` / `Local Owner`), Discovery Limit, Companies Discovered, Persona Leads, Outreach Drafts Created (counting `Draft Created`, `Draft Created in Gmail`, `Sent`, and `Pushed to Sequencer`), Suppressed Records, Dispatch Channel, and Primary AI Model with live client-scoped API key status (`Connected` or `Missing: CLIENT_{NAME}_...`).
* **Rows 16–17+ (`Campaign Performance Rollup`, `A17:G...`):** Summary table (`Campaign ID`, `Campaign Name`, `Department`, `Status`, `Companies Found`, `Gmail Drafts`, `Last Run`) **auto-filtered on Column C (`Department`) by Selected Client / Owner**.

#### 2. `EmailEvents` *(11 Canonical Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Central atomic event ledger powering `Email Deliverability` and `ControlPanel.html` deliverability metrics.
* **Columns (1–11):**
  1. `event_id` — Deterministic SHA-256 unique ID (`EVT-SNT-*`, `EVT-DLV-*`, `EVT-BNC-*`, `EVT-RPL-*`) preventing duplicate entries across syncs.
  2. `client_id` — Client ownership (`Growth Target`, `eValuater`, etc.).
  3. `campaign_id` — Campaign attribution (`EVA-CAN-001`, `CMP-005`, etc.).
  4. `provider` — Strict dropdown: `Gmail`, `Microsoft 365`.
  5. `provider_message_id` — RFC `Message-ID`, Gmail ID, or Microsoft Graph `internetMessageId` / `id`.
  6. `event_type` — Strict dropdown: `Sent`, `Delivered`, `Bounced`, `Replied`.
  7. `event_date` — Timestamp formatted `yyyy-mm-dd hh:mm:ss`.
  8. `recipient` — Lowercase recipient email address.
  9. `smtp_code` — Plain text (`@`) SMTP code (`250 2.0.0`, `550 5.1.1`, `451 4.7.1`, `UNKNOWN`, etc.).
  10. `bounce_type` — Dropdown: `Hard`, `Soft` (blank on non-bounce rows).
  11. `reason` — Parsed DSN/NDR diagnostic explanation, reply subject snippet, or delivery confirmation note.

#### 3. `Email Deliverability` *(Dynamic Analytics & Charts Tab)*
* **Filter Bar (Row 3):** `B3` **Client Filter** (`All Clients` + registered clients), `D3` **Campaign Filter** (`All Campaigns` + campaigns), `F3:J3` Active Scope label. Changing `B3` to a different client automatically resets `D3` to `All Campaigns` via `onEdit()` if the previous campaign belonged to a different client.
* **Headline Tiles (Rows 6–7):** `Sent` (`A7`), `Delivered` (`C7`), `Hard Bounces` (`E7`), `Soft Bounces` (`G7`), and `Replies` (`I7`) computed live via `COUNTIFS` respecting `$B$3` and `$D$3`.
* **Performance Breakdown (Rows 10–11):** `Bounce Rate` (`(Hard + Soft) / Sent` in `A11`), `Delivery Rate` (`Delivered / Sent` in `D11`), and `Reply Rate` (`Replied / Delivered` with `Sent` fallback in `G11`), formatted `0.0%`.
* **Tables & Embedded Charts (`Cols A–J` + Charts in `Col L`):**
  * **Campaign Comparison (`A13:J29` + Column Chart `L13`):** Per-campaign `Sent`, `Delivered`, `Hard Bounces`, `Soft Bounces`, `Replied`, `Client`, `Delivery Rate`, `Bounce Rate`, `Reply Rate` (falls back to `Campaigns` tab when 0 events exist).
  * **SMTP Error Distribution (`A31:E47` + Donut Chart `L31`):** `QUERY` + `FILTER` grouping `Bounced` rows by `smtp_code` with `Occurrences`, `Bounce Type`, `% of Bounces`, and `Sample Reason`.
  * **Daily Sending Volume (`A49:E80` + Column Chart `L49`):** Last 30 active days sorted chronologically ascending (`Date`, `Sent`, `Delivered`, `Bounced`, `Replied`).

#### 4. `API_Usage_Log` *(10 Canonical Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Atomic cost and consumption ledger for all LLM and email verifier API calls.
* **Columns (1–10):** `Timestamp`, `Campaign ID`, `Service Name`, `Operation`, `HTTP Status`, `Input Tokens / Units`, `Output Tokens`, `Est. Cost ($)`, `Status`, `Details / Error` *(automatically tagged with `[Client: {ClientName}]`)*.

#### 5. `Campaigns` *(Locked 28 Columns — Auto-Filtered by Client / Owner Col C)*
* **Purpose:** Master multi-campaign storage preserving complete isolation across campaigns and clients.
* **Columns (1–28):** `Campaign ID`, `Campaign Name`, `Department` *(Client / Owner with dynamic validation from `Lists & Settings`)*, `Status`, `Target Countries`, `Target Industries`, `Company Size`, `Discovery Limit`, `Research Depth`, `Evidence Period Months`, `Min Company Score`, `Min Persona Score`, `Min Email Priority Score`, `Target Titles`, `Campaign Focus`, `Search Keywords`, `Custom Business Needs`, `AI Provider`, `AI Model`, `Email Verifier Provider`, `Email Verifier Endpoint`, `Email Sender Provider` *(9 validated options including `Microsoft 365 (Drafts Only)` and `Microsoft 365 (Direct Send)`)*, `Email Sender Endpoint`, `Sender Name`, `Sender Title`, `Sender Company`, `Campaign Config JSON` *(extended settings backup including `senderEmail`, `replyToEmail`, and `presetNeed`)*, `Last Modified`.
* **Formatting:** Plain integer (`0`) enforced on `Discovery Limit` (Col H) and `Evidence Period` / `Score` columns (Cols J–M).

#### 6. `Company Intelligence` *(25 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Storage for enriched account profiles generated during Agent 2 deep research.
* **Columns (1–25):** `Campaign ID`, `Campaign Name`, `Company ID`, `Company Name`, `Country`, `Industry`, `Company Size`, `ESG Frameworks`, `Reporting Obligations`, `ESG Reporting Maturity` *(6 allowed states: `Early Stage`, `Developing / Spreadsheet-reliant`, `Developing`, `Mature`, `Advanced`, `N/A`)*, `Supply Chain Profile`, `Scope 3 Exposure`, `Sustainability Data Needs`, `Digital Transformation Signals`, `ESG Governance`, `Business Triggers`, `eValuater Solution Fit` *(`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`)*, `Company Opportunity Score`, `Company Opportunity`, `Evidence URLs`, `Evidence Checked Date`, `Verification Status`, `Human Review Status`, `Last Enriched At`, `Company Score Rationale`.

#### 7. `Persona Leads` *(29 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Decision-makers identified by Agent 2 with contact intelligence and email verification diagnostics.
* **Columns (1–29):** `Campaign ID`, `Campaign Name`, `Company ID`, `Company Name`, `Industry`, `Persona ID`, `Name`, `Title`, `Persona Type`, `LinkedIn URL`, `Signal`, `Signal Date`, `Evidence URL`, `Buying Signal`, `Verification Status`, `Confidence`, `Persona Fit Score`, `Personalized Hook`, `Review Status`, `Last Verified At`, `Business Email`, `Email Domain`, `Email Source URL`, `Email Discovery Method`, `Email Verification Status` *(`valid`, `risky`, `invalid`, `catch-all`, `unverified`)*, `Email Verification Date`, `Email Verification Provider`, `Email Identity Confidence`, `Email Verification Notes`.

#### 8. `Persona Mapper Settings` *(Locked 39 Canonical Columns — Auto-Filtered by Active Campaign ID Col A)*
* **Purpose:** Direct-mapped horizontal runtime configuration feeding Agent 2 (`Row 2+`).
* **Formatting:** Plain integer (`0`) enforced on Cols I, Z–AB, AI; percentage format (`0%`) on scoring weight columns.

#### 9. `Prospect Review Queue` *(45 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Human-in-the-loop review station staging personalized outreach drafts before dispatch.
* **Key Validated & Tracking Fields:** `Priority Tier` (`Tier 1`, `Tier 2`, `Tier 3`), `Review Status` (Col R: `Needs Review`, `Approved`, `Rejected`), `Human Approval Status` (Col V), `Send Status` (Col Y: `Queued`, `Draft Created in Gmail`, `Draft Created` *(used by Microsoft 365 Drafts Only)*, `Sent` *(used by Direct Send & promoted drafts)*, `Pushed to Sequencer`, `Suppressed - Not Sent`, `Draft Failed`), `Dispatch Mode` (Col AA), `Thread / Message ID` (Col AB — stores Gmail/M365 draft ID or sent message ID for promotion and bounce/reply correlation), `Compliance Status` (Col AP), and `Suppression Check` (Col AQ).

#### 10. `Campaign Configuration` *(Locked 38 Vertical Parameters, Rows 2–39)*
* **Purpose:** Active campaign key-value configuration view synced from `ControlPanel.html` and `Campaigns`.
* **Locked Parameters (Rows 2–39):** `CAMPAIGN_ID` (Row 2), `CAMPAIGN_NAME` (Row 3), `DEPARTMENT` (Row 4 — Client / Account Owner with dynamic dropdown), `TARGET_COUNTRIES`, `TARGET_INDUSTRIES`, `COMPANY_SIZE_RANGE`, `MAX_COMPANIES`, `RESEARCH_DEPTH`, `EVIDENCE_PERIOD_MONTHS`, `MIN_COMPANY_SCORE`, `MIN_PERSONA_SCORE`, `MIN_EMAIL_PRIORITY_SCORE`, `TARGET_TITLES`, `CAMPAIGN_FOCUS`, `SEARCH_KEYWORDS`, `CUSTOM_BUSINESS_NEEDS`, `AI_PROVIDER`, `PRIMARY_AI_PROVIDER`, `AI_MODEL`, `EMAIL_VERIFICATION_PROVIDER`, `EMAIL_VERIFICATION_ENDPOINT`, `EMAIL_SENDER_PROVIDER` & `DISPATCH_MODE` *(validated across `Workspace Gmail`, `Google Workspace`, `Microsoft 365 (Drafts Only)`, `Microsoft 365 (Direct Send)`, `Smartlead`, `Instantly`, `Custom Webhook`)*, `EMAIL_SENDER_ENDPOINT`, `SENDER_NAME`, `SENDER_TITLE`, `SENDER_COMPANY`, 9 `INCLUDE_*` boolean flags, `HUMAN_APPROVAL_REQUIRED`, and `AUTO_SEND_ENABLED`. Plain integer format (`0`) enforced on numeric rows.

#### 11. `Company Database` *(10 Columns — Auto-Filtered by Active Campaign ID Col J)*
* **Purpose:** Master discovered company directory supporting **Scenario B Campaign-Scoped Deduplication** (allowing the same account to be discovered cleanly in separate campaigns).
* **Columns (1–10):** `Company ID`, `Company Name`, `Domain`, `Country`, `Industry`, `Employee Size`, `Discovery Method`, `Discovery Date`, `Status`, `Campaign ID` (Col J).

#### 12. `Research Evidence` *(9 Canonical Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Auditable grounding log populated during both Agent 1 (`Entity Type: Company`) and Agent 2 (`Entity Type: Persona`) with anti-hallucination quality downgrade guardrails.
* **Columns (1–9):** `Evidence ID`, `Campaign ID`, `Entity Name`, `Entity Type`, `Search Query Used`, `Grounding Evidence URL`, `Extracted Snippet / Signal`, `Verification Score`, `Date Captured`.

#### 13. `Agent Settings` *(4 Columns — Auto-Filtered by Active Client / Owner Col A)*
* **Purpose:** Displays `Growth Target` local defaults (15 rows including `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`) plus auto-generated `"Client: {Name}"` blocks (12 rows per active external client including `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET`).
* **Security Indicators & Guardrails:** Displays `"Configured in Script Properties (Protected)"` badges for encrypted API keys, relay secrets, and Microsoft 365 credentials (never raw secrets) alongside editable per-client `DAILY_SPEND_CAP_USD` and `MIN_VERIFIER_CREDITS` thresholds.

#### 14. `Execution Logs` *(6 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Timestamped execution telemetry (`Timestamp`, `Campaign ID`, `Action`, `Target Lead`, `Status`, `Details`).

#### 15. `Lists & Settings` *(3 Columns: `List name`, `Value`, `Active status`)*
* **Purpose:** Global client registry and suppression rule store.
* **Pinned Client Sorting:** `sortListsAndSettingsTab_()` keeps all `Client - Account` rows pinned cleanly at the top (`Growth Target` first, Active clients A–Z, Inactive archived clients A–Z), followed by Whitelist and Suppression rules (`Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, `Suppression - Internal Exclusion`).

---

# 3. Dual-AI Engine, Grounding & Financial Guardrails

The AI layer decouples prompts from model providers through **`AIProvider.js`**, routing calls dynamically with strict client-scoped credentials:

                         ┌─────────────────────────────────────────┐
                         │     AIProvider.js (Universal Router)    │
                         │   Passes clientName & campaignId Scope  │
                         └────────────────────┬────────────────────┘
                                              │
                      Pre-Flight Budget Check │ (`ApiTracker.js`)
                                              ▼
              ┌──────────────────────────────────────────────────────────────┐
              │                                                              │
              ▼                                                              ▼
    ┌───────────────────────────────────┐          ┌───────────────────────────────────┐
    │        GeminiProvider.js          │          │        OpenAIProvider.js          │
    ├───────────────────────────────────┤          ├───────────────────────────────────┤
    │ • Model Cascade:                  │          │ • Universal /v1/responses API     │
    │   3.8-flash -> 3.5-flash-lite ->  │          │ • Guarded Chat Completions        │
    │   2.5-flash                       │          │   fallback (`gpt-4o`/`4o-mini`)   │
    │ • Google Search Grounding         │          │ • Strict CLIENT_{NAME}_OPENAI_KEY │
    │ • 429 Ungrounded JSON Fallback    │          │ • Token & Cost Logging            │
    │ • thinkingConfig & Thought Filter │          └───────────────────────────────────┘
    └─────────────────┬─────────────────┘                            │
                      └───────────────────────┬──────────────────────┘
                                              ▼
                         ┌─────────────────────────────────────────┐
                         │    ResearchSchema.js / Results / Storage│
                         │  (JSON Contracts, Grounding & Dedup)    │
                         └─────────────────────────────────────────┘

### Grounding, Parsing & Cost Tracking Workflow
1. **Direct REST Transport:** Built natively on `UrlFetchApp` with `ScriptLock` atomic writes (`ResearchStorage.js`) and offline unit checks (`ResearchChecks.js`).
2. **Search Grounding & Anti-Hallucination Guardrails:** `ResearchResults.js` parses Gemini Search Grounding chunks and OpenAI citations into `Research Evidence` (Tab 12), downgrading ungrounded claims automatically.
3. **Centralized Cost Engine (`ApiTracker.js`):** Every Gemini, OpenAI, and Email Verifier request checks `checkApiBudgetGuardrails_(campaignId, clientName)` before execution and logs tokens/units, USD cost, and `[Client: {Name}]` tags to `API_Usage_Log`.

---

# 4. End-to-End Agent Architecture & Email Deliverability Pipeline Flow

    [User / ControlPanel.html / Top Menu Trigger]
                     │  (Selects Client/Owner -> Cascading Campaign + Sender/Reply-To + Deliverability Snapshot)
                     ▼
        ┌───────────────────────┐
        │        Code.js        │ ──> Resolves Client Scope (Growth Target vs. CLIENT_{NAME}_*)
        └───────────────────────┘     & Updates Native Auto-Filters across Operational Tabs (incl. EmailEvents).
                     │
                     ▼
        ┌───────────────────────┐
        │  CompanyDiscovery.js  │ ──> Agent 1: Discovers target accounts + ESG/Need signals
        │   & PersonaMapper.js  │ ──> Agent 2: Maps personas, emails, suppression & verifiers
        └───────────────────────┘
                     │
                     ▼
        ┌───────────────────────┐
        │      AIProvider.js    │ ──> Universal Router: dispatches requests dynamically
        └───────────────────────┘     using strict client-scoped credentials (zero fallback).
                     │
           ┌─────────┴─────────┐
           ▼                   ▼
    ┌───────────────┐   ┌─────────────────┐
    │GeminiProvider │   │OpenAIProvider.js│ ──> Live Google Search Grounding + Model Cascade
    └───────────────┘   └─────────────────┘     or OpenAI /v1/responses + Chat fallback.
           │                   │
           └─────────┬─────────┘
                     ▼
        ┌───────────────────────┐
        │     ApiTracker.js     │ ──> Pre-flight client spend hard-caps, cost calculation,
        └───────────────────────┘     & [Client: ...] logging to Tab 4 (API_Usage_Log).
                     │
                     ▼
        ┌───────────────────────┐
        │   ResearchSchema.js   │ ──> Enforces strict JSON contracts & schemas
        │  & ResearchResults.js │     + anti-hallucination quality downgrade guardrails.
        └───────────────────────┘
                     │
                     ▼
        ┌───────────────────────┐
        │   ResearchStorage.js  │ ──> Atomic ScriptLock writes + Scenario B dedup to:
        └───────────────────────┘     Tabs 6, 7, 9, 11 (Col J), 12 (Research Evidence).
                     │
                     ▼  (After Human Review / Approval in Tab 9: Prospect Review Queue)
        ┌───────────────────────┐
        │     EmailSender.js    │ ──> Strict Client-Isolated Dispatch Routing + Message ID Capture:
        └───────────────────────┘     • Growth Target -> Local GmailApp / Local MS_* / Sequencers
                                      • External Google Workspace -> ClientGmailRelay.gs (v2.2)
                                      • External Microsoft 365 -> Graph API (CLIENT_{NAME}_MS_*)
                                      • External Sequencers -> CLIENT_{NAME}_* Smartlead/Instantly/Webhook
                     │
                     ├──> Logs inline 'Sent' event -> Tab 2 (EmailEvents) + Tab 9 (`Thread / Message ID`)
                     ▼
        ┌──────────────────────────────────────────────────────────────────────┐
        │                   EmailEventsDeliverability.js                       │
        ├──────────────────────────────────────────────────────────────────────┤
        │ • Hourly Trigger (`syncEmailDeliverabilityEvents`) & Silent Sidebar  │
        │   Sync (`syncEmailDeliverabilityFromSidebar`)                        │
        │ • Draft-to-Sent promotion (Gmail, Relay v2.2 & Microsoft 365)        │
        │ • SMTP DSN/NDR Hard vs. Soft bounce parsing + Reply detection        │
        │ • Exchange Message Trace v1.0 (optional) or 4h no-bounce Delivery    │
        │   confirmation + late-bounce reconciliation                          │
        │ • Powers Tab 3 (`Email Deliverability` + 3 Charts) & ControlPanel    │
        └──────────────────────────────────────────────────────────────────────┘
                     │
                     ▼
        ┌───────────────────────┐
        │  Dashboard & Sheets   │ ──> Updates live KPIs, Client-Filtered Performance Rollup,
        └───────────────────────┘     Send Status, Execution Logs & Verifier Low-Balance Alerts.

---

# 5. Multi-Client Isolation, Security Vault & Client Dispatch Engines

### Strict Zero-Fallback Credential Matrix

| Operating Tier | Example Owner | Credential Namespace (Script Properties) | Email Dispatch & Event Polling Isolation | Local Fallback Allowed? |
| :--- | :--- | :--- | :--- | :--- |
| **Local / Agency Brand** | `Growth Target` | Default keys (`GEMINI_API_KEY`, `OPENAI_API_KEY`, `VERIFIER_API_KEY`, `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`) — explicitly ignores `CLIENT_*` | Native Workspace `GmailApp` (`Drafts Only` / `Direct Send` + DSN/Reply polling), local Microsoft 365 Graph (`MS_*`), or local sequencer keys | N/A (Local Scope) |
| **External Client** | `eValuater`, `Pilot Acme` | Strictly prefixed keys (`CLIENT_{CLIENT_NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`, `_GMAIL_RELAY_URL`, `_GMAIL_RELAY_SECRET`, `_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET`, `_SMARTLEAD_API_KEY`, `_INSTANTLY_API_KEY`, `_WEBHOOK_URL` / `_EMAIL_SENDER_WEBHOOK_URL`, `_WEBHOOK_SECRET`) | **Google Workspace (Option A):** Client's standalone `ClientGmailRelay.gs` (v2.2) Web App (dispatch + `pollEvents`)<br>**Microsoft 365 (Native Graph):** Direct OAuth2 + Exchange Online `/messages`, `/send`, & optional Message Trace<br>**Sequencers (Option B):** Client-scoped Smartlead / Instantly / Webhook | **NO (Zero Fallback / Fail-Closed)** |

### Credential Security Rules & Access Controls
* **Server-Side Encrypted Storage Only:** Raw API keys, relay secrets, and Microsoft Entra Tenant/Client IDs and Secrets live strictly in **Apps Script Script Properties** — never plain text inside Google Sheets cells and never passed to `ControlPanel.html` or browser clients.
* **Script-Editor Access Restriction:** Because Script Properties live inside the Apps Script project attached to the spreadsheet, **only trusted agency administrators should be granted Editor access to the master Google Sheet** (any Editor can open **Extensions → Apps Script** and view Script Properties).
* **Placeholder Rejection:** `isStudioValidSecretValue_()` and `isValidSenderSecretValue_()` reject placeholder indicators (`PASTE_`, `Configured in Script Properties`, `Protected`) so status labels in `Agent Settings` are never sent as live API keys.
* **Least-Privilege Default:** Outreach dispatch defaults to **`Drafts Only`**, requiring explicit human approval before creating drafts or sending live outreach.

### Workspace Gmail Relay Setup for External Clients (`ClientGmailRelay.gs` v2.2 — Option A)
1. **Client Workspace Deployment:** Open Google Apps Script inside the **client's Google Workspace account**, paste `ClientGmailRelay.gs` (v2.2), set `RELAY_SECRET` in Script Properties, run `authorizeRelayPermissions()` once, and deploy as a **Web App** (*Execute as: Me*, *Who has access: Anyone* — protected by `RELAY_SECRET` in the JSON POST body; when upgrading from v2.1, click **Deploy → Manage deployments → Edit → New version → Deploy** to keep the same `/exec` URL).
2. **Master Studio Binding:** In the Master AI Prospecting Studio **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_URL` = Deployed client Web App `/exec` URL
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_SECRET` = Matching `RELAY_SECRET`
3. **Sender Mailbox & Reply-To Routing:** Configure `Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` (validated against `GmailApp.getAliases()` with safe primary mailbox fallback), and `Reply-To Email` inside `ControlPanel.html` (stored inside `Campaigns` Col AA `Campaign Config JSON`).
4. **Fail-Closed Verification & Polling:** External client drafts, sends, and `pollEvents` bounce/reply/draft-promotion queries execute exclusively inside the client's own Gmail mailbox and never leak into the agency operator's Gmail if relay keys are missing.

### Native Microsoft 365 / Exchange Online Setup for Clients (Direct Microsoft Graph API — Zero External Platform)
1. **Client Entra ID Setup (`entra.microsoft.com`):** Client's Microsoft 365 Global/App Admin registers a Single-Tenant App Registration, creates a Client Secret (copying the **Value**), and grants **Microsoft Graph Application permissions** `Mail.ReadWrite` (Drafts + NDR bounce & reply reading) and `Mail.Send` (Direct Send) with Admin Consent *(or restricts to a single mailbox via Exchange Online RBAC for Applications; optionally adds `ExchangeMessageTrace.Read.All` for raw server-level delivery logs)*.
2. **Master Studio Binding:** In the Master AI Prospecting Studio **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_MS_TENANT_ID` = Directory (tenant) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_ID` = Application (client) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_SECRET` = Client Secret Value
3. **Campaign Mailbox & Reply-To Binding:** In `ControlPanel.html` (Section 5), select **`Microsoft 365 (Drafts Only)`** or **`Microsoft 365 (Direct Send)`**, leave `Email Sender Endpoint` blank, enter the client's Exchange Online **Primary Mailbox / UPN** (e.g., `sales@evaluater.ai`) in `Sender Email / Mailbox` (validates before saving), and set `Reply-To Email Address`.
4. **Graph API Execution & Batch Token Caching (`EmailSender.js` & `EmailEventsDeliverability.js`):**
   * Authenticates via `login.microsoftonline.com/{tenantId}/oauth2/v2.0/token` (`client_credentials`, cached in memory per batch run).
   * **`Microsoft 365 (Drafts Only)`:** Calls `POST https://graph.microsoft.com/v1.0/users/{senderEmail}/messages` (`Prefer: IdType="ImmutableId"`, `isDraft: true`, `replyTo`, `toRecipients`), sets `Send Status = 'Draft Created'`, and records `draftId` in `Thread / Message ID`.
   * **`Microsoft 365 (Direct Send)`:** Creates draft + calls `POST https://graph.microsoft.com/v1.0/users/{senderEmail}/messages/{id}/send` (falling back to `/sendMail` if only `Mail.Send` is granted), sets `Send Status = 'Sent'`, stamps `Sent At` & `Thread / Message ID` (`internetMessageId`), and writes `'Sent'` to `EmailEvents`.
   * Safe-draft review (`createGmailDraftsForApprovedProspects()`) automatically honors Microsoft 365 campaigns by routing them to `Microsoft 365 (Drafts Only)` instead of Google Workspace Relay.

---

# 6. Active Suppression Guard, Data Validation, Native Auto-Filter Engine & Deliverability Quotas

1. **Active Suppression Guard (`loadSuppressionRulesFromSettings_`):**
   * Cross-references `Lists & Settings` for active `Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, and `Suppression - Internal Exclusion` entries without colliding with `Client - Account` rows.
   * Blocks role-based addresses (`info@`, `sales@`, `support@`, `admin@`, `billing@`, `contact@`, `help@`, `office@`, `jobs@`, `careers@`) and malformed email syntax, tagging rows as `Rejected (Suppression Guard)` / `Suppressed - Not Sent`.
2. **Strict Data Validation & Formatting (`SheetSetup.js` & `EmailEventsDeliverability.js`):**
   * Enforces strict dropdown validations on `Department / Client Owner`, `Status`, `Email Sender Provider` / `DISPATCH_MODE` (9 options including Microsoft 365), `ESG Reporting Maturity` (6 states), `eValuater Solution Fit` (`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`), `Email Verification Status` (`valid`, `risky`, `invalid`, `catch-all`, `unverified`), `Prospect Review Queue` status columns, `EmailEvents` (`provider`, `event_type`, `bounce_type`), and `Email Deliverability` filter bar (`B3`, `D3`).
   * Enforces plain number formats (`0` / `0.00`) on limits, scores, and caps, and `@` plain text on `EmailEvents!I` (`smtp_code`).
3. **Native Auto-Filter Engine (`Code.js`):**
   * **Filtered by Active Campaign ID:** `EmailEvents`, `Prospect Review Queue`, `Persona Leads`, `Company Intelligence`, `Company Database` (Col J), `Research Evidence`, `Execution Logs`, `API_Usage_Log`, and `Persona Mapper Settings` (Col A).
   * **Filtered by Selected Client / Owner:** `Campaigns` (Col C), `Dashboard` → `Campaign Performance Rollup` (Row 17+, Col C), and `Agent Settings` (Col A).
4. **Deliverability Scale & Apps Script Quota Considerations:**
   * Sync runs under `ScriptLock`, caps lookback to 7 days (`<= 300` Graph messages / `<= 50` Gmail threads per mailbox), and batches writes in a single `setValues()`. Beyond ~25 active client mailboxes, split sync triggers per client group to stay comfortably inside Apps Script's 6-minute execution limit and 20,000/day Workspace `UrlFetchApp` quota.

---

# 7. Pilot Sandbox Runner & Verification System

The platform includes an end-to-end simulation and zero-trace cleanup engine (**`PilotSandboxRunner.js`**):

* **Client-Isolated Simulation (`runPilotOnRealSheets()`):** Generates `PILOT-CMP-001` under sandbox client `"Pilot Acme"` across all 15 tabs, tests suppression screening and `catch-all` soft warnings, logs sample Gemini & ZeroBounce usage tagged `[Client: Pilot Acme]` to `API_Usage_Log`, seeds 11 multi-day deliverability events (`PILOT-EVT-001`..`011`) into `EmailEvents` (verifying `60.0%` Delivery Rate, `40.0%` Bounce Rate, `33.3%` Reply Rate, and all 3 `Email Deliverability` charts), appends/updates `Dashboard` Rollup non-destructively, and creates safe Gmail drafts.
* **Option A Full Wipe (`cleanupPilotDataFromRealSheets()`):** Purges zero-trace of all `PILOT-*` and `Pilot Acme` records across all 15 tabs (including `EmailEvents`, `Dashboard` Rollup rows, `Agent Settings` client block including `CLIENT_PILOT_ACME_*`, `Lists & Settings` client/suppression entries, `API_Usage_Log`, and `Execution Logs`), resets `Email Deliverability` `B3`/`D3` if pointing to `Pilot Acme` / `PILOT-*`, refreshes client & deliverability dropdown validations, and cleanly restores your active production campaign (or resets cleanly to `Standby` if no real campaigns remain).
* **System Health Check (`runStudioSpreadsheetAudit()`):** Audits all 13 core tabs, both Deliverability tabs (`EmailEvents` & `Email Deliverability`), hourly sync trigger status, local vs. client-isolated Script Properties, active client registry, and dispatch mode configuration.

---

# 8. Menu Navigation & Operational Controls

The platform provides a streamlined Google Sheets top menu (`onOpen()` in `Code.js`):

    🎯 AI Prospecting Studio
    ├── 🚀 Open Campaign Control Panel          -> Launches ControlPanel.html Client-First Sidebar
    ├── 🔄 Filter Tabs to Active Campaign       -> Refreshes all Campaign & Client native filters
    ├── ──────────────
    ├── 🔍 1. Discover Target Companies         -> Runs Agent 1 (CompanyDiscovery.js)
    ├── 👤 2. Run Persona Mapper                -> Runs Agent 2 (PersonaMapper.js)
    ├── ──────────────
    ├── 📋 Go to Prospect Review Queue          -> Navigates to Tab 9 (Prospect Review Queue)
    ├── ✅ Approve Selected Prospect            -> Marks selected row Approved
    ├── ✏️ Mark Selected: Needs Editing         -> Marks selected row Needs Review
    ├── ❌ Reject Selected Prospect             -> Marks selected row Rejected
    ├── ──────────────
    ├── 📝 Create Outreach Drafts (Safe Review) -> Dispatches drafts via Workspace Gmail or Microsoft 365
    ├── ✉️ Send Approved Outreach Emails        -> Sends approved outreach via configured channel
    ├── ──────────────
    ├── 📊 Open Email Deliverability            -> Navigates to Tab 3 (Email Deliverability)
    ├── 🔄 Sync Email Deliverability Events     -> Runs immediate Gmail/Relay/M365 event sync
    ├── ⏱️ Enable Hourly Deliverability Sync    -> Installs hourly time-driven sync trigger
    ├── ──────────────
    ├── ❓ Check Spreadsheet Setup              -> Runs 15-tab, Trigger & Script Properties health check
    ├── ──────────────
    ├── 🧪 Run Authentic Pilot Pipeline         -> Runs client-isolated 15-tab pilot simulation
    └── 🧹 Clean Up Pilot Test Data             -> Executes Option A Full Wipe & restores active state

---

# 9. Technical Stack & Repository Structure

* **Repository:** `https://github.com/growthtarget/ai-prospecting-studio`
* **17 Core Apps Script Project Files + Manifest + Client Template:**
  * **Core, UI & Setup:** `Code.js`, `SheetSetup.js`, `Config.js`, `ControlPanel.html`
  * **AI & Cost Engine:** `AIProvider.js`, `GeminiProvider.js`, `OpenAIProvider.js`, `ApiTracker.js`
  * **Research & Grounding Subsystem:** `ResearchSchema.js`, `ResearchResults.js`, `ResearchStorage.js`, `ResearchChecks.js`
  * **Operational Pipeline, Deliverability & Simulation:** `CompanyDiscovery.js`, `PersonaMapper.js`, `EmailSender.js`, `EmailEventsDeliverability.js`, `PilotSandboxRunner.js`
  * **Standalone Client Onboarding Template:** `ClientGmailRelay.gs` (v2.2)
  * **Apps Script Manifest:** `appsscript.json` *(includes `script.scriptapp` and `https://mail.google.com/` scopes)*

---

# 10. Operational Playbook for Non-Technical Users

1. **System Health Check & Permissions:**
   * In **Extensions → Apps Script**, open `Code.gs` and run `authorizeStudioPermissions` once to verify Gmail, Sheets, External Request, and Trigger permissions.
   * Click **🎯 AI Prospecting Studio → ❓ Check Spreadsheet Setup** to verify all 13 core tabs, both Deliverability tabs, hourly trigger status, local agency keys, and client-isolated Script Properties are recognized.
   * Click **🎯 AI Prospecting Studio → ⏱️ Enable Hourly Deliverability Sync** once to start automated hourly deliverability ingestion.
2. **Client Onboarding (Optional for External Clients):**
   * In the **Campaign Control Panel** sidebar (`🚀 Open Campaign Control Panel`), select an existing client or click **`[ + Add New Client... ]`** under **`CLIENT / OWNER`**.
   * Add the client's encrypted keys inside **Extensions → Apps Script → Project Settings → Script Properties**:
     * **AI & Verification:** `CLIENT_{NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`
     * **If Google Workspace Outreach:** `_GMAIL_RELAY_URL` + `_GMAIL_RELAY_SECRET` (from `ClientGmailRelay.gs` v2.2)
     * **If Microsoft 365 / Exchange Online Outreach:** `_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET` (from client's Entra ID App Registration)
   * Check Tab 13 (`Agent Settings`) to verify `"Configured in Script Properties (Protected)"` badges appear.
3. **Campaign Setup (Client-First Cascading Workspace):**
   * Select your **`CLIENT / OWNER`** first (the **`ACTIVE CAMPAIGN`** menu filters strictly to that client, or auto-suggests a new ID like `PA-001` if the client has 0 saved campaigns).
   * Configure geography, industry, company size, limits, sticky **Pain-Point Presets**, **AI Provider & Model**, **Email Verifier**, **Email Sender Provider** (test **`Drafts Only`** first!), and **Sender Identity** (`Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` [required primary mailbox/UPN for Microsoft 365], `Reply-To Email`).
   * Click **Save Campaign Settings** to synchronize all tabs and apply native auto-filters.
4. **Pipeline Execution:**
   * Run **Step 1: Discover Target Companies** and inspect results in `Company Database` and `Research Evidence`.
   * Run **Step 2: Run Persona Mapper** to populate `Company Intelligence`, `Persona Leads`, `Research Evidence`, and `Prospect Review Queue`, while syncing live KPIs on `Dashboard`.
5. **Review, Dispatch & Deliverability Monitoring:**
   * Open `Prospect Review Queue` (`📋 Go to Prospect Review Queue`), review copy and verification statuses, and mark rows **Approved** (`✅ Approve Selected Prospect`).
   * Run **📝 Create Outreach Drafts (Safe Review)** (or **✉️ Send Approved Outreach Emails**) — drafts/emails execute safely inside `Growth Target`'s mailbox, the external client's Google Workspace via `ClientGmailRelay.gs` v2.2, or the external client's Exchange Online mailbox via Microsoft Graph API (`Draft Created` / `Sent`), while immediately logging `Sent` events into `EmailEvents`.
   * Monitor live **Sent / Delivered / Hard & Soft Bounces / Replies / SMTP Codes / Daily Volume** inside Section 6 of `ControlPanel.html` (filtered automatically by your selected Client & Campaign) or on Tab 3 (`Email Deliverability`), and click **`⚡ Sync Events`** anytime for on-demand refreshes.
6. **Sandbox Testing & Zero-Trace Cleanup:**
   * Run **🧪 Run Authentic Pilot Pipeline** anytime to verify end-to-end mechanics (including sample `EmailEvents` & `Email Deliverability` charts), and run **🧹 Clean Up Pilot Test Data** afterward for an Option A full wipe back to your production state.