# AI PROSPECTING STUDIO + PERSONA MAPPER (v2.0)
## System Master Brief, Architecture & Technical Specification

**Project:** AI Prospecting Studio + Persona Mapper  
**Version:** 2.0 (Multi-Client Production Architecture)  
**Repository:** [https://github.com/growthtarget/ai-prospecting-studio](https://github.com/growthtarget/ai-prospecting-studio)  
**Primary AI Engine:** Google Gemini API (`gemini-3.8-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-flash` cascade with Google Search Grounding & ungrounded 429 fallback)  
**Fallback AI Engine:** OpenAI API (`gpt-4o` / `gpt-4o-mini` via universal `/v1/responses` API & guarded Chat Completions fallback)  
**Application Interface:** Google Sheets (13 Dedicated Synchronized Tabs)  
**Backend:** Google Apps Script (16 Modular Project Files + `appsscript.json` + standalone `ClientGmailRelay.gs` v2.1 onboarding template)  
**Operating Model:** Multi-Client Agency Isolation (`Growth Target` local brand + zero-fallback `CLIENT_{NAME}_*` external client workspaces)  
**Verification Providers:** ZeroBounce, MillionVerifier, NeverBounce, Custom HTTPS  
**Dispatch Engines:** Native Google Workspace Gmail (`Growth Target` Drafts & Direct Send), Client Google Workspace Gmail Relay (`ClientGmailRelay.gs` v2.1 — Option A), Native Microsoft 365 / Exchange Online via Microsoft Graph API v1.0 (`Drafts Only` & `Direct Send` — Zero External Platform), Smartlead, Instantly, Custom Webhook (Option B)  

---

# 1. Executive Summary & System Vision

**AI Prospecting Studio + Persona Mapper (v2.0)** is an autonomous, multi-client B2B outbound intelligence, persona discovery, and hyper-personalized outreach platform built natively inside Google Sheets and powered by Google Apps Script.

Operating from a single agency Google Workspace spreadsheet, the platform manages multiple independent client accounts under strict credential, mailbox, budget, and spreadsheet-view isolation while unifying **Market Research (Agent 1: Company Discovery)** and **Decision-Maker Persona Mapping & Draft Synthesis (Agent 2: Persona Mapper)** into one governed pipeline.

### Core Capabilities:
1. **Multi-Client / Owner Isolation & Client-First Workspace:** Manages the local agency brand (`Growth Target`) alongside external clients (e.g., `eValuater`) from one master workbook. External clients enforce strict `CLIENT_{NAME}_*` Script Properties with **zero fallback** to local keys or agency Gmail; clients marked `Inactive` in `Lists & Settings` are cleanly archived from sidebar and sheet dropdowns without deleting historical campaign records.
2. **Dynamic Dual-AI Routing & Resiliency Cascade:** Primary execution on Google Gemini (`gemini-3.8-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-flash`) with live Google Search Grounding, `thinkingConfig` speedups, and automatic 429 ungrounded JSON fallback—plus OpenAI `/v1/responses` support via `AIProvider.js`.
3. **Centralized API Cost Tracking & Per-Client Budget Guardrails:** `ApiTracker.js` logs every LLM and email verifier call to `API_Usage_Log` tagged with `[Client: {Name}]`, calculates estimated USD costs, enforces per-client daily spend hard caps, and dispatches low-balance alerts via `MailApp`.
4. **Active Suppression Guard & Email Verification:** Screens blacklisted domains, opt-out emails, competitor lists, and role-based accounts (`info@`, `admin@`, `support@`) via `Lists & Settings`. Verifies emails per client via ZeroBounce, MillionVerifier, NeverBounce, or Custom HTTPS (strict fail-closed enforcement for external clients when verifier != `None`; safe local unverified pass-through when agency keys are unconfigured; `catch-all`/`risky` soft-warning pass-through).
5. **Isolated Multi-Tab Schemas & Native Auto-Filter Engine:** Locks the 38 vertical parameters in `Campaign Configuration`, 28 columns in `Campaigns` (storing extended fields like `senderEmail` and `replyToEmail` in `Campaign Config JSON`), and 39 columns in `Persona Mapper Settings`. Native Google Sheets auto-filters isolate operational tabs by Active Campaign ID and `Campaigns`, `Dashboard` Rollup, and `Agent Settings` by Selected Client / Owner.
6. **Multi-Platform Client Dispatch (Google Workspace Relay, Native Microsoft 365 / Exchange Online & Sequencers):**
   * **Google Workspace Clients (`ClientGmailRelay.gs` v2.1):** Creates drafts or sends emails directly inside each external client's own Google Workspace mailbox via a standalone Web App relay with `senderEmail` alias fallback and `replyToEmail` support.
   * **Microsoft 365 / Exchange Online Clients (Native Microsoft Graph API v1.0):** Creates drafts (`POST /users/{mailbox}/messages` -> `Draft Created`) or sends emails (`POST /users/{mailbox}/sendMail` -> `Sent`) directly in the client's Exchange Online mailbox via OAuth2 `client_credentials` using `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET` — requiring zero external middleware platforms (no Make, Zapier, n8n, or Power Automate needed).
   * **Sequencers & Webhooks:** Routes to client-scoped Smartlead, Instantly, or Custom HTTPS Webhooks.
7. **Full-Pipeline Pilot Sandbox & Option A Full Wipe:** `PilotSandboxRunner.js` runs an end-to-end simulation (`PILOT-CMP-001` under `"Pilot Acme"`) across all 13 tabs and includes a zero-trace Option A cleanup (`cleanupPilotDataFromRealSheets()`) that removes all pilot/client data and cleanly restores the active production workspace.

---

# 2. 13-Tab System Architecture & Master Schemas

The application operates across **13 tightly synchronized sheets** styled with high-contrast Dark Navy headers (`#174ea6`) on Light Blue (`#e8f0fe`), frozen header rows, plain integer/number formats (`0` / `0.00`) to prevent 1900 date conversion bugs, and native Google Sheets dropdown validations.

```text
┌─────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                   13 SYNCHRONIZED WORKBOOK TABS                                 │
├─────────────────────────────────────────────────────────────────────────────────────────────────┤
│  01. Dashboard                │ KPI Matrix (Cols A-D) & Client-Filtered Performance Rollup (17+) │
│  02. API_Usage_Log (10 Cols)  │ Token/credit costs tagged [Client: ...] (Auto-filtered: Camp ID)│
│  03. Campaigns (Locked 28)    │ Master campaign registry & JSON config (Auto-filtered: Owner)   │
│  04. Company Intelligence(25) │ Enriched account profiles & scores (Auto-filtered: Camp ID)     │
│  05. Persona Leads (29 Cols)  │ Decision-makers & email verification (Auto-filtered: Camp ID)   │
│  06. Persona Mapper Settings  │ Locked 39 canonical Agent 2 settings (Auto-filtered: Camp ID)   │
│  07. Prospect Review Queue(45)│ Human-in-the-loop triage & outreach (Auto-filtered: Camp ID)    │
│  08. Campaign Configuration   │ Locked 38 vertical active campaign parameters (Rows 2-39)       │
│  09. Company Database (10)    │ Master account directory w/ Col J Campaign ID (Filtered: Camp)  │
│  10. Research Evidence (9)    │ Grounded URLs, snippets & scores (Auto-filtered: Camp ID)       │
│  11. Agent Settings (4 Cols)  │ Local defaults & protected Client blocks (Filtered: Owner)      │
│  12. Execution Logs (6 Cols)  │ Timestamped pipeline audits (Auto-filtered: Camp ID)            │
│  13. Lists & Settings (3 Cols)│ Pinned Client - Account registry + Whitelist/Suppression rules  │
└─────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### Detailed Tab Specifications

#### 1. `Dashboard` *(Cols A–G)*
* **Rows 1–2:** Green Banner (`A1:D1`), Active Campaign Name (`C1`), Status (`D1`), and Active Campaign ID (`B2`) with bi-directional `onEdit()` sync to `Campaigns` and `Campaign Configuration`.
* **Rows 3–13:** 10-Row KPI Matrix (`A3:D13`) displaying Active Campaign Name, Campaign ID, Operating Department (`Client Owner` / `Local Owner`), Discovery Limit, Companies Discovered, Persona Leads, Outreach Drafts Created (counting `Draft Created`, `Draft Created in Gmail`, `Sent`, and `Pushed to Sequencer`), Suppressed Records, Dispatch Channel, and Primary AI Model with live client-scoped API key status (`Connected` or `Missing: CLIENT_{NAME}_...`).
* **Rows 16–17+ (`Campaign Performance Rollup`, `A17:G...`):** Summary table (`Campaign ID`, `Campaign Name`, `Department`, `Status`, `Companies Found`, `Gmail Drafts`, `Last Run`) **auto-filtered on Column C (`Department`) by Selected Client / Owner**.

#### 2. `API_Usage_Log` *(10 Canonical Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Atomic cost and consumption ledger for all LLM and email verifier API calls.
* **Columns (1–10):** `Timestamp`, `Campaign ID`, `Service Name`, `Operation`, `HTTP Status`, `Input Tokens / Units`, `Output Tokens`, `Est. Cost ($)`, `Status`, `Details / Error` *(automatically tagged with `[Client: {ClientName}]`)*.

#### 3. `Campaigns` *(Locked 28 Columns — Auto-Filtered by Client / Owner Col C)*
* **Purpose:** Master multi-campaign storage preserving complete isolation across campaigns and clients.
* **Columns (1–28):** `Campaign ID`, `Campaign Name`, `Department` *(Client / Owner with dynamic validation from `Lists & Settings`)*, `Status`, `Target Countries`, `Target Industries`, `Company Size`, `Discovery Limit`, `Research Depth`, `Evidence Period Months`, `Min Company Score`, `Min Persona Score`, `Min Email Priority Score`, `Target Titles`, `Campaign Focus`, `Search Keywords`, `Custom Business Needs`, `AI Provider`, `AI Model`, `Email Verifier Provider`, `Email Verifier Endpoint`, `Email Sender Provider` *(9 validated options including `Microsoft 365 (Drafts Only)` and `Microsoft 365 (Direct Send)`)*, `Email Sender Endpoint`, `Sender Name`, `Sender Title`, `Sender Company`, `Campaign Config JSON` *(extended settings backup including `senderEmail`, `replyToEmail`, and `presetNeed`)*, `Last Modified`.
* **Formatting:** Plain integer (`0`) enforced on `Discovery Limit` (Col H) and `Evidence Period` / `Score` columns (Cols J–M).

#### 4. `Company Intelligence` *(25 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Storage for enriched account profiles generated during Agent 2 deep research.
* **Columns (1–25):** `Campaign ID`, `Campaign Name`, `Company ID`, `Company Name`, `Country`, `Industry`, `Company Size`, `ESG Frameworks`, `Reporting Obligations`, `ESG Reporting Maturity` *(6 allowed states: `Early Stage`, `Developing / Spreadsheet-reliant`, `Developing`, `Mature`, `Advanced`, `N/A`)*, `Supply Chain Profile`, `Scope 3 Exposure`, `Sustainability Data Needs`, `Digital Transformation Signals`, `ESG Governance`, `Business Triggers`, `eValuater Solution Fit` *(`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`)*, `Company Opportunity Score`, `Company Opportunity`, `Evidence URLs`, `Evidence Checked Date`, `Verification Status`, `Human Review Status`, `Last Enriched At`, `Company Score Rationale`.

#### 5. `Persona Leads` *(29 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Decision-makers identified by Agent 2 with contact intelligence and email verification diagnostics.
* **Columns (1–29):** `Campaign ID`, `Campaign Name`, `Company ID`, `Company Name`, `Industry`, `Persona ID`, `Name`, `Title`, `Persona Type`, `LinkedIn URL`, `Signal`, `Signal Date`, `Evidence URL`, `Buying Signal`, `Verification Status`, `Confidence`, `Persona Fit Score`, `Personalized Hook`, `Review Status`, `Last Verified At`, `Business Email`, `Email Domain`, `Email Source URL`, `Email Discovery Method`, `Email Verification Status` *(`valid`, `risky`, `invalid`, `catch-all`, `unverified`)*, `Email Verification Date`, `Email Verification Provider`, `Email Identity Confidence`, `Email Verification Notes`.

#### 6. `Persona Mapper Settings` *(Locked 39 Canonical Columns — Auto-Filtered by Active Campaign ID Col A)*
* **Purpose:** Direct-mapped horizontal runtime configuration feeding Agent 2 (`Row 2+`).
* **Formatting:** Plain integer (`0`) enforced on Cols I, Z–AB, AI; percentage format (`0%`) on scoring weight columns.

#### 7. `Prospect Review Queue` *(45 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Human-in-the-loop review station staging personalized outreach drafts before dispatch.
* **Key Validated Status Fields:** `Priority Tier` (`Tier 1`, `Tier 2`, `Tier 3`), `Review Status` (Col R: `Needs Review`, `Approved`, `Rejected`), `Human Approval Status` (Col V), `Send Status` (Col Y: `Queued`, `Draft Created in Gmail`, `Draft Created` *(used by Microsoft 365 Drafts Only)*, `Sent` *(used by Direct Send)*, `Pushed to Sequencer`, `Suppressed - Not Sent`, `Draft Failed`), `Compliance Status` (Col AP), and `Suppression Check` (Col AQ).

#### 8. `Campaign Configuration` *(Locked 38 Vertical Parameters, Rows 2–39)*
* **Purpose:** Active campaign key-value configuration view synced from `ControlPanel.html` and `Campaigns`.
* **Locked Parameters (Rows 2–39):** `CAMPAIGN_ID` (Row 2), `CAMPAIGN_NAME` (Row 3), `DEPARTMENT` (Row 4 — Client / Account Owner with dynamic dropdown), `TARGET_COUNTRIES`, `TARGET_INDUSTRIES`, `COMPANY_SIZE_RANGE`, `MAX_COMPANIES`, `RESEARCH_DEPTH`, `EVIDENCE_PERIOD_MONTHS`, `MIN_COMPANY_SCORE`, `MIN_PERSONA_SCORE`, `MIN_EMAIL_PRIORITY_SCORE`, `TARGET_TITLES`, `CAMPAIGN_FOCUS`, `SEARCH_KEYWORDS`, `CUSTOM_BUSINESS_NEEDS`, `AI_PROVIDER`, `PRIMARY_AI_PROVIDER`, `AI_MODEL`, `EMAIL_VERIFICATION_PROVIDER`, `EMAIL_VERIFICATION_ENDPOINT`, `EMAIL_SENDER_PROVIDER` & `DISPATCH_MODE` *(validated across `Workspace Gmail`, `Google Workspace`, `Microsoft 365 (Drafts Only)`, `Microsoft 365 (Direct Send)`, `Smartlead`, `Instantly`, `Custom Webhook`)*, `EMAIL_SENDER_ENDPOINT`, `SENDER_NAME`, `SENDER_TITLE`, `SENDER_COMPANY`, 9 `INCLUDE_*` boolean flags, `HUMAN_APPROVAL_REQUIRED`, and `AUTO_SEND_ENABLED`. Plain integer format (`0`) enforced on numeric rows.

#### 9. `Company Database` *(10 Columns — Auto-Filtered by Active Campaign ID Col J)*
* **Purpose:** Master discovered company directory supporting **Scenario B Campaign-Scoped Deduplication** (allowing the same account to be discovered cleanly in separate campaigns).
* **Columns (1–10):** `Company ID`, `Company Name`, `Domain`, `Country`, `Industry`, `Employee Size`, `Discovery Method`, `Discovery Date`, `Status`, `Campaign ID` (Col J).

#### 10. `Research Evidence` *(9 Canonical Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Auditable grounding log populated during both Agent 1 (`Entity Type: Company`) and Agent 2 (`Entity Type: Persona`) with anti-hallucination quality downgrade guardrails.
* **Columns (1–9):** `Evidence ID`, `Campaign ID`, `Entity Name`, `Entity Type`, `Search Query Used`, `Grounding Evidence URL`, `Extracted Snippet / Signal`, `Verification Score`, `Date Captured`.

#### 11. `Agent Settings` *(4 Columns — Auto-Filtered by Active Client / Owner Col A)*
* **Purpose:** Displays `Growth Target` local defaults (15 rows including `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`) plus auto-generated `"Client: {Name}"` blocks (12 rows per active external client including `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET`).
* **Security Indicators & Guardrails:** Displays `"Configured in Script Properties (Protected)"` badges for encrypted API keys, relay secrets, and Microsoft 365 credentials (never raw secrets) alongside editable per-client `DAILY_SPEND_CAP_USD` and `MIN_VERIFIER_CREDITS` thresholds.

#### 12. `Execution Logs` *(6 Columns — Auto-Filtered by Active Campaign ID)*
* **Purpose:** Timestamped execution telemetry (`Timestamp`, `Campaign ID`, `Action`, `Target Lead`, `Status`, `Details`).

#### 13. `Lists & Settings` *(3 Columns: `List name`, `Value`, `Active status`)*
* **Purpose:** Global client registry and suppression rule store.
* **Pinned Client Sorting:** `sortListsAndSettingsTab_()` keeps all `Client - Account` rows pinned cleanly at the top (`Growth Target` first, Active clients A–Z, Inactive archived clients A–Z), followed by Whitelist and Suppression rules (`Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, `Suppression - Internal Exclusion`).

---

# 3. Dual-AI Engine, Grounding & Financial Guardrails

The AI layer decouples prompts from model providers through **`AIProvider.js`**, routing calls dynamically with strict client-scoped credentials:

```text
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
                     │    ResearchSchema.js / Results / Storage   │
                     │  (JSON Contracts, Grounding & Dedup)    │
                     └─────────────────────────────────────────┘
```

### Grounding, Parsing & Cost Tracking Workflow
1. **Direct REST Transport:** Built natively on `UrlFetchApp` with `ScriptLock` atomic writes (`ResearchStorage.js`) and offline unit checks (`ResearchChecks.js`).
2. **Search Grounding & Anti-Hallucination Guardrails:** `ResearchResults.js` parses Gemini Search Grounding chunks and OpenAI citations into `Research Evidence` (Tab 10), downgrading ungrounded claims automatically.
3. **Centralized Cost Engine (`ApiTracker.js`):** Every Gemini, OpenAI, and Email Verifier request checks `checkApiBudgetGuardrails_(campaignId, clientName)` before execution and logs tokens/units, USD cost, and `[Client: {Name}]` tags to `API_Usage_Log`.

---

# 4. End-to-End Agent Architecture & Pipeline Flow

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
    └───────────────────────┘     • Growth Target -> Local GmailApp / Local MS_* / Sequencers
                                  • External Google Workspace -> ClientGmailRelay.gs (v2.1)
                                  • External Microsoft 365 -> Graph API (CLIENT_{NAME}_MS_*)
                                  • External Sequencers -> CLIENT_{NAME}_* Smartlead/Instantly/Webhook
                │
                ▼
    ┌───────────────────────┐
    │  Dashboard & Sheets   │ ──► Updates live KPIs, Client-Filtered Performance Rollup,
    └───────────────────────┘     Send Status, Execution Logs & Verifier Low-Balance Alerts.
```

---

# 5. Multi-Client Isolation, Security Vault & Client Dispatch Engines

### Strict Zero-Fallback Credential Matrix

| Operating Tier | Example Owner | Credential Namespace (Script Properties) | Email Dispatch Isolation | Local Fallback Allowed? |
| :--- | :--- | :--- | :--- | :--- |
| **Local / Agency Brand** | `Growth Target` | Default keys (`GEMINI_API_KEY`, `OPENAI_API_KEY`, `VERIFIER_API_KEY`, `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`) — explicitly ignores `CLIENT_*` | Native Workspace `GmailApp` (`Drafts Only` / `Direct Send`), local Microsoft 365 Graph (`MS_*`), or local sequencer keys | N/A (Local Scope) |
| **External Client** | `eValuater`, `Pilot Acme` | Strictly prefixed keys (`CLIENT_{CLIENT_NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`, `_GMAIL_RELAY_URL`, `_GMAIL_RELAY_SECRET`, `_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET`, `_SMARTLEAD_API_KEY`, `_INSTANTLY_API_KEY`, `_WEBHOOK_URL` / `_EMAIL_SENDER_WEBHOOK_URL`, `_WEBHOOK_SECRET`) | **Google Workspace (Option A):** Client's standalone `ClientGmailRelay.gs` Web App<br>**Microsoft 365 (Native Graph):** Direct OAuth2 + Exchange Online `/messages` & `/sendMail`<br>**Sequencers (Option B):** Client-scoped Smartlead / Instantly / Webhook | **NO (Zero Fallback / Fail-Closed)** |

### Credential Security Rules
* **Server-Side Encrypted Storage Only:** Raw API keys, relay secrets, and Microsoft Entra Tenant/Client IDs and Secrets live strictly in **Apps Script Script Properties** — never plain text inside Google Sheets cells.
* **Placeholder Rejection:** `isStudioValidSecretValue_()` and `isValidSenderSecretValue_()` reject placeholder indicators (`PASTE_`, `Configured in Script Properties`, `Protected`) so status labels in `Agent Settings` are never sent as live API keys.
* **Least-Privilege Default:** Outreach dispatch defaults to **`Drafts Only`**, requiring explicit human approval before creating drafts or sending live outreach.

### Workspace Gmail Relay Setup for External Clients (`ClientGmailRelay.gs` v2.1 — Option A)
1. **Client Workspace Deployment:** Open Google Apps Script inside the **client's Google Workspace account**, paste `ClientGmailRelay.gs` (v2.1), set `RELAY_SECRET` in Script Properties, and deploy as a **Web App** (*Execute as: Me*, *Who has access: Anyone* — protected by `RELAY_SECRET` in the JSON POST body).
2. **Master Studio Binding:** In the Master AI Prospecting Studio **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_URL` = Deployed client Web App `/exec` URL
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_SECRET` = Matching `RELAY_SECRET`
3. **Sender Mailbox & Reply-To Routing:** Configure `Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` (validated against `GmailApp.getAliases()` with safe primary mailbox fallback), and `Reply-To Email` inside `ControlPanel.html` (stored inside `Campaigns` Col AA `Campaign Config JSON`).
4. **Fail-Closed Verification:** External client drafts and sends execute exclusively inside the client's own Gmail mailbox and never leak into the agency operator's Gmail if relay keys are missing.

### Native Microsoft 365 / Exchange Online Setup for Clients (Direct Microsoft Graph API — Zero External Platform)
1. **Client Entra ID Setup (`entra.microsoft.com`):** Client's Microsoft 365 Global/App Admin registers a Single-Tenant App Registration, creates a Client Secret (copying the **Value**), and grants **Microsoft Graph Application permissions** `Mail.ReadWrite` (Drafts) and `Mail.Send` (Direct Send) with Admin Consent *(or restricts to a single mailbox via Exchange Online RBAC for Applications without tenant-wide Entra permissions)*.
2. **Master Studio Binding:** In the Master AI Prospecting Studio **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_MS_TENANT_ID` = Directory (tenant) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_ID` = Application (client) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_SECRET` = Client Secret Value
3. **Campaign Mailbox & Reply-To Binding:** In `ControlPanel.html` (Section 5), select **`Microsoft 365 (Drafts Only)`** or **`Microsoft 365 (Direct Send)`**, leave `Email Sender Endpoint` blank, enter the client's Exchange Online **Primary Mailbox / UPN** (e.g., `sales@evaluater.ai`) in `Sender Email / Mailbox` (validates before saving), and set `Reply-To Email Address`.
4. **Graph API Execution & Batch Token Caching (`EmailSender.js`):**
   * Authenticates via `login.microsoftonline.com/{tenantId}/oauth2/v2.0/token` (`client_credentials`, cached in memory per batch run).
   * **`Microsoft 365 (Drafts Only)`:** Calls `POST https://graph.microsoft.com/v1.0/users/{senderEmail}/messages` (`isDraft: true`, `replyTo`, `toRecipients`) and sets `Send Status = 'Draft Created'`.
   * **`Microsoft 365 (Direct Send)`:** Calls `POST https://graph.microsoft.com/v1.0/users/{senderEmail}/sendMail` (`saveToSentItems: true`, `replyTo`) and sets `Send Status = 'Sent'` + stamps `Sent At`.
   * Safe-draft review (`createGmailDraftsForApprovedProspects()`) automatically honors Microsoft 365 campaigns by routing them to `Microsoft 365 (Drafts Only)` instead of Google Workspace Relay.

---

# 6. Active Suppression Guard, Data Validation & Native Auto-Filter Engine

1. **Active Suppression Guard (`loadSuppressionRulesFromSettings_`):**
   * Cross-references `Lists & Settings` for active `Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, and `Suppression - Internal Exclusion` entries without colliding with `Client - Account` rows.
   * Blocks role-based addresses (`info@`, `sales@`, `support@`, `admin@`, `billing@`, `contact@`, `help@`, `office@`, `jobs@`, `careers@`) and malformed email syntax, tagging rows as `Rejected (Suppression Guard)` / `Suppressed - Not Sent`.
2. **Strict Data Validation & Formatting (`SheetSetup.js`):**
   * Enforces strict dropdown validations on `Department / Client Owner`, `Status`, `Email Sender Provider` / `DISPATCH_MODE` (9 options including Microsoft 365), `ESG Reporting Maturity` (6 states), `eValuater Solution Fit` (`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`), `Email Verification Status` (`valid`, `risky`, `invalid`, `catch-all`, `unverified`), and `Prospect Review Queue` status columns.
   * Enforces plain number formats (`0` / `0.00`) on limits, scores, and caps so Google Sheets never converts integers (such as `80` or `85`) into 1900 dates.
3. **Native Auto-Filter Engine (`Code.js`):**
   * **Filtered by Active Campaign ID:** `Prospect Review Queue`, `Persona Leads`, `Company Intelligence`, `Company Database` (Col J), `Research Evidence`, `Execution Logs`, `API_Usage_Log`, and `Persona Mapper Settings` (Col A).
   * **Filtered by Selected Client / Owner:** `Campaigns` (Col C), `Dashboard` → `Campaign Performance Rollup` (Row 17+, Col C), and `Agent Settings` (Col A).

---

# 7. Pilot Sandbox Runner & Verification System

The platform includes an end-to-end simulation and zero-trace cleanup engine (**`PilotSandboxRunner.js`**):

* **Client-Isolated Simulation (`runPilotOnRealSheets()`):** Generates `PILOT-CMP-001` under sandbox client `"Pilot Acme"` across all 13 tabs, tests suppression screening and `catch-all` soft warnings, logs sample Gemini & ZeroBounce usage tagged `[Client: Pilot Acme]` to `API_Usage_Log`, appends/updates `Dashboard` Rollup non-destructively, and creates safe Gmail drafts (via optional Pilot Acme Relay or local sandbox Gmail).
* **Option A Full Wipe (`cleanupPilotDataFromRealSheets()`):** Purges zero-trace of all `PILOT-*` and `Pilot Acme` records across all 13 tabs (including `Dashboard` Rollup rows, `Agent Settings` client block including `CLIENT_PILOT_ACME_*`, `Lists & Settings` client/suppression entries, `API_Usage_Log`, and `Execution Logs`), refreshes client dropdown validations, and cleanly restores your active production campaign (or resets cleanly to `Standby` if no real campaigns remain).
* **System Health Check (`runStudioSpreadsheetAudit()`):** Audits all 13 tabs, local vs. client-isolated Script Properties, active client registry, and dispatch mode configuration.

---

# 8. Menu Navigation & Operational Controls

The platform provides a streamlined Google Sheets top menu (`onOpen()` in `Code.js`):

```text
🎯 AI Prospecting Studio
├── 🚀 Open Campaign Control Panel          -> Launches ControlPanel.html Client-First Sidebar
├── 🔄 Filter Tabs to Active Campaign       -> Refreshes all Campaign & Client native filters
├── ──────────────
├── 🔍 1. Discover Target Companies         -> Runs Agent 1 (CompanyDiscovery.js)
├── 👤 2. Run Persona Mapper                -> Runs Agent 2 (PersonaMapper.js)
├── ──────────────
├── 📋 Go to Prospect Review Queue          -> Navigates to Tab 7 (Prospect Review Queue)
├── ✅ Approve Selected Prospect            -> Marks selected row Approved
├── ✏️ Mark Selected: Needs Editing         -> Marks selected row Needs Review
├── ❌ Reject Selected Prospect             -> Marks selected row Rejected
├── ──────────────
├── 📝 Create Outreach Drafts (Safe Review) -> Dispatches drafts via Workspace Gmail or Microsoft 365
├── ✉️ Send Approved Outreach Emails        -> Sends approved outreach via configured channel
├── ──────────────
├── ❓ Check Spreadsheet Setup              -> Runs 13-tab & Script Properties health check
├── ──────────────
├── 🧪 Run Authentic Pilot Pipeline         -> Runs client-isolated 13-tab pilot simulation
└── 🧹 Clean Up Pilot Test Data             -> Executes Option A Full Wipe & restores active state
```

---

# 9. Technical Stack & Repository Structure

* **Repository:** `https://github.com/growthtarget/ai-prospecting-studio`
* **16 Core Apps Script Project Files + Manifest + Client Template:**
  * **Core, UI & Setup:** `Code.js`, `SheetSetup.js`, `Config.js`, `ControlPanel.html`
  * **AI & Cost Engine:** `AIProvider.js`, `GeminiProvider.js`, `OpenAIProvider.js`, `ApiTracker.js`
  * **Research & Grounding Subsystem:** `ResearchSchema.js`, `ResearchResults.js`, `ResearchStorage.js`, `ResearchChecks.js`
  * **Operational Pipeline & Simulation:** `CompanyDiscovery.js`, `PersonaMapper.js`, `EmailSender.js`, `PilotSandboxRunner.js`
  * **Standalone Client Onboarding Template:** `ClientGmailRelay.gs` (v2.1)
  * **Apps Script Manifest:** `appsscript.json`

---

# 10. Operational Playbook for Non-Technical Users

1. **System Health Check:**
   * Click **🎯 AI Prospecting Studio → ❓ Check Spreadsheet Setup** to verify all 13 tabs, local agency keys, and client-isolated Script Properties are recognized.
2. **Client Onboarding (Optional for External Clients):**
   * In the **Campaign Control Panel** sidebar (`🚀 Open Campaign Control Panel`), select an existing client or click **`[ + Add New Client... ]`** under **`CLIENT / OWNER`**.
   * Add the client's encrypted keys inside **Extensions → Apps Script → Project Settings → Script Properties**:
     * **AI & Verification:** `CLIENT_{NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`
     * **If Google Workspace Outreach:** `_GMAIL_RELAY_URL` + `_GMAIL_RELAY_SECRET` (from `ClientGmailRelay.gs` v2.1)
     * **If Microsoft 365 / Exchange Online Outreach:** `_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET` (from client's Entra ID App Registration)
   * Check Tab 11 (`Agent Settings`) to verify `"Configured in Script Properties (Protected)"` badges appear.
3. **Campaign Setup (Client-First Cascading Workspace):**
   * Select your **`CLIENT / OWNER`** first (the **`ACTIVE CAMPAIGN`** menu filters strictly to that client, or auto-suggests a new ID like `PA-001` if the client has 0 saved campaigns).
   * Configure geography, industry, company size, limits, sticky **Pain-Point Presets**, **AI Provider & Model**, **Email Verifier**, **Email Sender Provider** (test **`Drafts Only`** first!), and **Sender Identity** (`Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` [required primary mailbox/UPN for Microsoft 365], `Reply-To Email`).
   * Click **Save Campaign Settings** to synchronize all tabs and apply native auto-filters.
4. **Pipeline Execution:**
   * Run **Step 1: Discover Target Companies** and inspect results in `Company Database` and `Research Evidence`.
   * Run **Step 2: Run Persona Mapper** to populate `Company Intelligence`, `Persona Leads`, `Research Evidence`, and `Prospect Review Queue`, while syncing live KPIs on `Dashboard`.
5. **Review & Client-Isolated Dispatch:**
   * Open `Prospect Review Queue` (`📋 Go to Prospect Review Queue`), review copy and verification statuses, and mark rows **Approved** (`✅ Approve Selected Prospect`).
   * Run **📝 Create Outreach Drafts (Safe Review)** (or **✉️ Send Approved Outreach Emails**) — drafts/emails execute safely inside `Growth Target`'s mailbox, the external client's Google Workspace via `ClientGmailRelay.gs`, or the external client's Exchange Online mailbox via Microsoft Graph API (`Draft Created` / `Sent`).
6. **Sandbox Testing & Zero-Trace Cleanup:**
   * Run **🧪 Run Authentic Pilot Pipeline** anytime to verify end-to-end mechanics, and run **🧹 Clean Up Pilot Test Data** afterward for an Option A full wipe back to your production state.