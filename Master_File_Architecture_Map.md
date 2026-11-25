# AI Prospecting Studio + Persona Mapper (v2.0) — Master File & Architecture Map

## 1. Core Navigation, UI & State Management

### `Code.js` (Core, Multi-Client Isolation & Native Filter Engine)
* **Role:** Manages the Google Sheets top menu (`onOpen()`, including `'📝 Create Outreach Drafts (Safe Review)'`, `'📊 Open Email Deliverability'`, `'🔄 Sync Email Deliverability Events'`, and `'⏱️ Enable Hourly Deliverability Sync'`), `onEdit()` bi-directional sync (Dashboard status/ID, `Email Deliverability` smart filter reset on `B3` change, & `Lists & Settings` auto-sorting), active campaign state (`getStudioActiveCampaignId_()`), and the complete **Multi-Client Isolation & Native View Filter Engine**.
* **Client Credential Isolation:** Implements `resolveClientScopedSecret_()`, `resolveCampaignClientName_()`, `isLocalAgencyClient_()`, `normalizeClientPropertyPrefix_()`, and `isStudioValidSecretValue_()` to enforce zero-fallback `CLIENT_{NAME}_*` Script Property isolation for external clients while restricting `"Growth Target"` to local default keys.
* **Client Workspace Management:** Implements `addClientAccountToSettings()`, `getRegisteredClientsList_()`, `getInactiveClientsList_()`, and `filterAgentSettingsForClientWorkspace()` (allowing instant workspace filtering and auto-suggested IDs even for 0-campaign clients), while refreshing `Email Deliverability` filter dropdowns (`refreshDeliverabilityFilterDropdowns_()`).
* **Multi-Tab Native View Filter Engine:**
  * `filterOperationalSheetsByCampaign_()`: Auto-filters `Prospect Review Queue`, `Persona Leads`, `Company Intelligence`, `Company Database`, `Research Evidence`, `Execution Logs`, `API_Usage_Log`, `Persona Mapper Settings`, and `EmailEvents` by the Active **Campaign ID**.
  * `filterCampaignsSheetByClient_()` & `filterDashboardRollupByClient_()`: Auto-filters `Campaigns` (Col C) and `Dashboard` → `Campaign Performance Rollup` (Row 17+, Col C) by the Selected **Client / Owner**.
* **Campaign State Sync & Health Audit:** `selectAndActivateCampaign()` and `studioDirectSaveCampaignPipeline()` sync the locked 38-row `Campaign Configuration`, 28-column `Campaigns` (storing extended fields like `senderEmail` and `replyToEmail` in `Campaign Config JSON`), 39-column `Persona Mapper Settings`, and 4-column `Dashboard` KPIs (with client-scoped AI key status and `Draft Created` / `Sent` counting). `runStudioSpreadsheetAudit()` verifies all 13 core tabs plus the 2 Deliverability tabs (`EmailEvents`, `Email Deliverability`) and hourly trigger status; `authorizeStudioPermissions()` verifies Gmail read/search/compose and trigger permissions.

### `SheetSetup.js` (Interface, Registry & Validation Formatter)
* **Role:** Initializes and formats all **15 tabs** (13 core tabs + `EmailEvents` and `Email Deliverability`) with high-contrast styling (`#174ea6` on `#e8f0fe`), freezes header rows, preserves the locked 38 `Campaign Configuration` parameters and 28 `Campaigns` columns, and enforces plain number formats (`0` / `0.00`) on numeric limits, scores, and caps to prevent 1900 date conversions.
* **Email Provider Validation Sync:** Applies the 9 validated `EMAIL_SENDER_PROVIDER` / `DISPATCH_MODE` options (`Workspace Gmail (Drafts Only)`, `Workspace Gmail (Direct Send)`, `Google Workspace (Drafts Only)`, `Google Workspace (Direct Send)`, `Microsoft 365 (Drafts Only)`, `Microsoft 365 (Direct Send)`, `Smartlead`, `Instantly`, `Custom Webhook`) to `Campaign Configuration` (Rows 23–24) and `Campaigns` (Col V).
* **Tab 15 (`Lists & Settings`) Registry Engine:**
  * `sortListsAndSettingsTab_()` pins all `Client - Account` rows at the top (`Growth Target` #1, Active clients A–Z, Inactive archived clients A–Z), followed by Whitelist and Suppression rules.
  * `refreshClientDropdownValidations_()` dynamically binds active `Client - Account` entries to `Campaign Configuration` (Row 4), `Campaigns` (Col C), and `Email Deliverability` (`B3` / `D3`).
  * `getInactiveClientsFromSettings_()` enables true client archival without deleting historical `Campaigns` rows.
* **Tab 13 (`Agent Settings`) Security Vault & Auto-Filter:** Pre-populates `Growth Target` local defaults (including `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`) and auto-generates persistent `"Client: {Name}"` blocks (12 rows per active external client, including `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET`)—displaying `"Configured in Script Properties (Protected)"` status indicators, editable per-client budget guardrails (`CLIENT_{NAME}_DAILY_SPEND_CAP_USD`, `CLIENT_{NAME}_MIN_VERIFIER_CREDITS`), and filtering via `filterAgentSettingsByClient_()`.

### `ControlPanel.html` (Client-First Sidebar UI)
* **Role:** Interactive Google Sheets sidebar (`380px`) driving campaign creation, configuration, agent triggers, client workspace switching, and live deliverability inspection.
* **Client-First Cascading Workspace:** Placing `CLIENT / OWNER` at the top acts as the primary workspace switcher—filtering `ACTIVE CAMPAIGN` strictly to the selected client's campaigns (preventing accidental cross-client overwrites), excluding `Inactive` archived clients, and supporting inline `[ + Add New Client... ]` creation.
* **0-Campaign Instant Workspace & Smart UI Controls:** Auto-filters `Agent Settings`, `Campaigns`, and `Dashboard` Rollup when switching to a 0-campaign client and auto-suggests unique IDs/names (e.g., `PA-001` / `Pilot Acme - Campaign 1`); includes sticky pain-point presets, AI Provider–filtered model dropdowns, busy concurrency locks (`setPipelineBusy_()`), `Microsoft 365 (Drafts Only)` / `Microsoft 365 (Direct Send)` provider options with dynamic Script Property & primary mailbox hints in `onSenderChanged()`, save-time mailbox validation for Microsoft 365, and dedicated `Sender Email` / `Reply-To Email` fields saved in `Campaign Config JSON`.
* **Section 6 (`Email Deliverability Snapshot`):** Read-only metrics grid querying `EmailEvents` via `queryEmailDeliverabilityMetrics(clientFilter, campaignFilter)` directly from the active `Client / Owner` and `Active Campaign` filters (with scope switcher: `Match Control Panel Filters`, `All Campaigns for Selected Client`, or `Agency-Wide`), plus **`🔄 Refresh Stats`**, **`⚡ Sync Events`** (`syncEmailDeliverabilityFromSidebar()`), and **`📊 Open Email Deliverability Tab`**. Zero credentials or secrets are ever exposed to the HTML client.

### `Config.js` (Canonical Schema & App Definitions)
* **Role:** Holds `APP_CONFIG` (model defaults, pricing rates, budget guardrail defaults, and core `sheetDefinitions` header schemas) used across setup and validation workflows.

### `appsscript.json` (Apps Script Manifest & OAuth Scopes)
* **Role:** Defines V8 runtime, `America/Vancouver` timezone, Stackdriver logging, and required OAuth scopes: `spreadsheets.currentonly`, `script.container.ui`, `script.external_request`, `script.scriptapp` (hourly trigger management), `gmail.send`, `gmail.compose`, and `https://mail.google.com/` (local Gmail draft promotion, DSN bounce parsing, `Message-ID` header capture, and reply polling).

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
* **Role:** Atomically records every external API request (Gemini, OpenAI, and Email Verifiers) to Tab 4 (`API_Usage_Log`), tagging `Details / Error` with `[Client: {ClientName}]`.
* **Per-Client Budget Guardrails:** Automatically calculates estimated USD costs, enforces pre-flight per-client daily spend hard caps (`CLIENT_{NAME}_DAILY_SPEND_CAP_USD` vs. local `DAILY_SPEND_CAP_USD` via `checkApiBudgetGuardrails_()`), and dispatches automated low-verifier-credit alerts (`CLIENT_{NAME}_MIN_VERIFIER_CREDITS`) via `MailApp`.

---

## 4. Research, Schema & Grounding Parsers

### `ResearchSchema.js` (Strict Data Contracts)
* **Role:** Defines strict JSON schemas for structured LLM outputs across Gemini and OpenAI, enforcing required company firmographics, persona signals, and citation evidence contracts.

### `ResearchResults.js` (Grounding & Citation Parser)
* **Role:** Transforms raw AI responses and live search grounding metadata into canonical spreadsheet rows matching Tab 11 (`Company Database` - 10 cols) and Tab 12 (`Research Evidence` - 9 cols) with anti-hallucination quality downgrade guardrails.

### `ResearchStorage.js` (Atomic Storage & Deduplication)
* **Role:** Manages atomic persistence under `ScriptLock` and Scenario B Campaign-scoped deduplication (allowing the same company to be targeted under separate campaigns without collision).

### `ResearchChecks.js` (Offline Test Harness)
* **Role:** Executes offline unit test suites and grounding verification assertions across Gemini, OpenAI, and canonical table schemas without altering live spreadsheet data.

---

## 5. Operational Pipeline, Agents, Outreach Relay & Email Deliverability Engine

### `CompanyDiscovery.js` (Agent 1: Account Discovery)
* **Role:** Executes AI-powered company discovery scoped to `clientName` and `campaignId` using Gemini/OpenAI and live Google Search Grounding. Inspects campaign focus/keywords/industry for ESG detection (decoupled from client/owner name), captures account evidence into Tab 12 (`Research Evidence`), populates Tab 11 (`Company Database` Col J = `Campaign ID`), and applies Scenario B multi-campaign deduplication.

### `PersonaMapper.js` (Agent 2: Decision-Maker Persona & Email Mapping)
* **Role:** Reads prior `Research Evidence` clues, identifies decision-makers using client-scoped AI keys, generates/validates email patterns, enforces the active suppression guard (`Lists & Settings`), verifies emails via `ZeroBounce` / `MillionVerifier` / `NeverBounce` / `Custom HTTPS` (strict fail-closed enforcement for external clients when verifier != `None`; safe unverified pass-through when local agency keys are unconfigured), logs verifier credits to `API_Usage_Log`, allows `catch-all` / `risky` pass-through with soft warnings, captures persona signals into Tab 12, stages leads in Tab 9 (`Prospect Review Queue`), and syncs live `Dashboard` KPIs.

### `EmailSender.js` (Client-Isolated Outreach Dispatch Engine)
* **Role:** Dispatches approved/queued outreach strictly isolated by Client / Owner with custom sender identity (`senderName`, `senderEmail`, `replyToEmail`), captures message/draft IDs in `Prospect Review Queue` (`Thread / Message ID`), and logs `Sent` events inline to `EmailEvents`:
  * **Header Alias Safety (`findSenderHeaderCol_()`):** Resolves both `Prospect Review Queue` header variants (`Draft Subject` / `Email Subject` / `Approved Email Subject`, `Draft Email` / `Email Body` / `Approved Email Body`, `Name` / `Persona Name`, `Send Error` / `Last Error`, `Dispatch Mode`, `Thread / Message ID`, `Draft Created At`).
  * **Microsoft 365 / Exchange Online (`isMicrosoft365Provider_()`, `dispatchMicrosoft365Outreach_()`, `getMicrosoftGraphAccessToken_()`, `parseMicrosoftGraphError_()`):** Checked first in `dispatchOutreach_()`. Authenticates directly via Microsoft Entra OAuth2 `client_credentials` (`login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`) with in-memory batch token caching using strictly client-scoped `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET` (or local `MS_*` for `Growth Target`; zero fallback). Requires a valid Exchange Online primary mailbox / UPN in `Sender Email / Mailbox` (with optional `MS_SENDER_EMAIL` property fallback). Uses `Prefer: IdType="ImmutableId"` on `POST /users/{senderEmail}/messages` (`isDraft: true` -> `Send Status = 'Draft Created'`) and `POST /users/{senderEmail}/messages/{id}/send` on Direct Send (falling back to `/sendMail` if an Entra app only granted `Mail.Send`) so `internetMessageId` & `id` are captured for tracking (`Send Status = 'Sent'`). Safe-draft mode (`createGmailDraftsForApprovedProspects()`) automatically routes Microsoft 365 campaigns to Microsoft 365 Drafts.
  * **Local Agency (`Growth Target` Google Workspace):** Uses native Google Workspace `GmailApp` (`createDraft` for Drafts Only, `createDraft(...).send()` for Direct Send to capture `Message-ID` & Gmail `messageId`) with alias/reply-to support, plus local `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`, or webhook endpoints.
  * **External Clients — Option A (Workspace Gmail Relay):** Routes `Drafts Only` and `Direct Send` strictly to the client's deployed `ClientGmailRelay.gs` (v2.2) Web App via `CLIENT_{NAME}_GMAIL_RELAY_URL` and `CLIENT_{NAME}_GMAIL_RELAY_SECRET`, capturing `draftId` / `messageId` / `rfcMessageId`. **Fail-closed:** external client drafts/sends never fall back to the agency operator's Gmail mailbox.
  * **External Clients — Option B (Sequencers & Webhooks):** Enforces strict `CLIENT_{NAME}_SMARTLEAD_API_KEY`, `CLIENT_{NAME}_INSTANTLY_API_KEY`, `CLIENT_{NAME}_WEBHOOK_URL` / `_EMAIL_SENDER_WEBHOOK_URL`, and `CLIENT_{NAME}_WEBHOOK_SECRET` keys with zero local fallback.
  * **Inline Deliverability Logging:** Every successful `Sent` dispatch via Gmail or Microsoft 365 immediately calls `recordEmailEvent_()` to write a deduplicated `Sent` row into `EmailEvents`.

### `EmailEventsDeliverability.js` *(New File)* — Email Performance & Deliverability Engine
* **Role:** Powers centralized deliverability tracking, ingestion, analytics, and sidebar/tab querying without modifying client or campaign configuration schemas.
* **Sheet & Chart Builders:**
  * `setupEmailDeliverabilityModule()`: Initializes/refreshes Tab 2 (`EmailEvents`) and Tab 3 (`Email Deliverability`).
  * `setupEmailEventsTab_()`: Enforces the 11-column event schema (`event_id`..`reason`), strict dropdown validations (`provider`, `event_type`, `bounce_type`), timestamp format (`yyyy-mm-dd hh:mm:ss`), and plain-text `@` format on `smtp_code`.
  * `setupEmailDeliverabilityDashboardTab_()`: Builds the dynamic `Email Deliverability` tab (`B3` Client Filter, `D3` Campaign Filter, 5 Headline KPI Tiles, 3 Rate Breakdown Tiles, 3 Breakdown Tables, and 3 Embedded Google Sheets Charts: Campaign Comparison Column Chart, SMTP Bounce Error Donut Chart, and Daily Sending Volume Column Chart).
  * `refreshDeliverabilityFilterDropdowns_()` & `syncDeliverabilityTabFilters_()`: Dynamically syncs Client & Campaign dropdowns from `Lists & Settings`, `Campaigns`, and `EmailEvents`.
* **Event Ingestion & Querying:**
  * `recordEmailEvent_()`: Atomic (`ScriptLock`) idempotent single-event writer using deterministic `SHA-256` `event_id`s (`buildEmailEventId_()`).
  * `syncEmailDeliverabilityEvents()` (top menu / hourly trigger) & `syncEmailDeliverabilityFromSidebar()` (silent sidebar wrapper with zero modal `ui.alert` blocking): Delegates to `runEmailDeliverabilitySync_()`, polling Local Gmail, External Client Gmail Relays (`action: 'pollEvents'`), and Microsoft 365 Graph mailboxes (`isDraft: false` promotion, optional Exchange Message Trace v1.0, and `Mail.ReadWrite` NDR bounce & reply scanning).
  * `parseSmtpBounceDetails_()`: Extracts recipient, original `Message-ID`, SMTP status codes (`550 5.1.1`, `451 4.7.1`, etc.), diagnostic reason, and classifies `Hard` vs. `Soft` bounces.
  * `installDeliverabilitySyncTrigger()`: Installs a duplicate-safe hourly time-driven trigger for `syncEmailDeliverabilityEvents`.
  * `queryEmailDeliverabilityMetrics(clientFilter, campaignFilter)` & `openEmailDeliverabilitySheet()`: Read-only metrics API consumed by `ControlPanel.html` and navigation helper for the top menu/sidebar.

### `ClientGmailRelay.gs` (Standalone External Client Onboarding Template — v2.2)
* **Role:** Deployed as a standalone Google Apps Script Web App inside an external client's own Google Workspace account so outreach drafts, direct sends, and deliverability event polling execute inside the **client's own Gmail mailbox and domain** without granting agency access to their full Workspace.
* **v2.2 Capabilities:**
  * Authenticates every POST request against `RELAY_SECRET` in the client's Script Properties (`CLIENT_{NAME}_GMAIL_RELAY_SECRET`).
  * Honors `senderName`, `senderEmail` (verifying aliases with safe primary mailbox fallback), and `replyTo` (`replyToEmail`).
  * `sendEmail` uses `GmailApp.createDraft(...).send()` and returns `draftId`, `messageId`, `threadId`, and `rfcMessageId`.
  * Adds `action: 'pollEvents'` (`handleRelayPollEvents_()`) to detect manually sent drafts (`promotedDrafts`), scan Gmail DSN bounces (`parseRelaySmtpBounceDetails_()` returning `smtpCode`, `bounceType`, `originalMessageId`, `reason`), and detect replies strictly from `trackedRecipients` (never exposing unrelated private inbox threads).

### `PilotSandboxRunner.js` (Client-Isolated Simulator & Option A Full Wipe)
* **Role:** Runs an authentic end-to-end simulation (`PILOT-CMP-001` under client `"Pilot Acme"`) across all 15 tabs, optionally tests client Gmail relay routing, logs sample API/verifier usage tagged `[Client: Pilot Acme]`, seeds 11 multi-day deliverability events (`PILOT-EVT-001`..`011`) into `EmailEvents` via `writePilotDeliverabilityEvents_()`, and appends/updates `Dashboard` Rollup non-destructively.
* **Option A Full Wipe (`cleanupPilotDataFromRealSheets()`):** Removes zero-trace of all `PILOT-*` and `Pilot Acme` test data across all 15 tabs (including `EmailEvents`, `Dashboard` Rollup rows, `Agent Settings` client block including `CLIENT_PILOT_ACME_*`, `Lists & Settings` client/suppression entries, `API_Usage_Log`, and `Execution Logs`), resets `Email Deliverability` `B3`/`D3` if pointing to `Pilot Acme` / `PILOT-*`, refreshes sheet & deliverability dropdown validations, and cleanly restores the active real campaign (or resets to `Standby` if none remain).

---

## Multi-Client / Owner Isolation Architecture

| Operating Tier | Example Owner | Credential Namespace (Script Properties) | Email Dispatch & Event Polling Isolation | Local Fallback Allowed? |
| :--- | :--- | :--- | :--- | :--- |
| **Local / Agency Brand** | `Growth Target` | Default keys (`GEMINI_API_KEY`, `OPENAI_API_KEY`, `VERIFIER_API_KEY`, `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `SMARTLEAD_API_KEY`, `INSTANTLY_API_KEY`, `WEBHOOK_URL`, `WEBHOOK_SECRET`) — explicitly ignores `CLIENT_*` | Native Workspace `GmailApp` (`Drafts Only` / `Direct Send` + DSN/Reply polling), local Microsoft 365 Graph (`MS_*`), or local sequencer keys | N/A (Local Scope) |
| **External Client** | `eValuater`, `Pilot Acme` | Strictly prefixed keys (`CLIENT_{CLIENT_NAME}_GEMINI_API_KEY`, `_OPENAI_API_KEY`, `_VERIFIER_API_KEY`, `_GMAIL_RELAY_URL`, `_GMAIL_RELAY_SECRET`, `_MS_TENANT_ID`, `_MS_CLIENT_ID`, `_MS_CLIENT_SECRET`, `_SMARTLEAD_API_KEY`, `_INSTANTLY_API_KEY`, `_WEBHOOK_URL`, `_EMAIL_SENDER_WEBHOOK_URL`, `_WEBHOOK_SECRET`) | **Option A:** Client's standalone `ClientGmailRelay.gs` (v2.2) Web App (dispatch + `pollEvents`)<br>**Native Microsoft 365:** Direct Microsoft Graph API (`/messages`, `/send`, `/messages` polling, optional Exchange Message Trace)<br>**Option B:** Client-scoped Smartlead / Instantly / Webhook | **NO (Zero Fallback / Fail-Closed)** |

### Credential Security, Script-Editor Protection & Archival Rules
* **Server-Side Encryption Only:** Raw API keys, relay secrets, and Microsoft Entra Tenant/Client IDs and Secrets live strictly in **Apps Script Script Properties** — never plain text in Google Sheets cells and never passed to `ControlPanel.html` or browser clients.
* **Script-Editor Access Warning:** Anyone granted **Editor** access on the master Google Sheet can open **Extensions → Apps Script** and view Script Properties. Restrict spreadsheet **Editor** permissions strictly to trusted agency administrators.
* **Placeholder Rejection:** `isStudioValidSecretValue_()` and `isValidSenderSecretValue_()` reject placeholder strings (`PASTE_`, `Configured in Script Properties`, `Protected`) so status labels in `Agent Settings` are never transmitted as live API keys.
* **True Archival:** Setting a client's `Active status` to `Inactive` in Tab 15 (`Lists & Settings`) removes that client from the Control Panel dropdown and sheet validations while preserving historical campaign rows in Tab 5 (`Campaigns`) and historical `EmailEvents` analytics.

---

## Workspace Gmail Relay Setup for Clients (`ClientGmailRelay.gs` v2.2 — Option A)

1. **Client-Side Deployment:** Open Google Apps Script inside the **client's Google Workspace account**, paste `ClientGmailRelay.gs` (v2.2), set `RELAY_SECRET` in the client project's **Script Properties**, run `authorizeRelayPermissions()` once, and deploy as a **Web App** (*Execute as: Me*, *Who has access: Anyone* — protected by `RELAY_SECRET` in the JSON POST body; when upgrading from v2.1, use **Manage deployments → Edit → New version** so the `/exec` URL stays unchanged).
2. **Agency Master Binding:** In the Master AI Prospecting Studio Apps Script **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_URL` = Client Web App `/exec` URL
   * `CLIENT_{CLIENT_NAME}_GMAIL_RELAY_SECRET` = Matching `RELAY_SECRET`
3. **Campaign Identity Binding:** In the Control Panel sidebar, configure `Sender Name`, `Sender Title`, `Sender Company`, `Sender Email` (mailbox or Workspace Send-As alias), and `Reply-To Email` (stored inside `Campaigns` Col AA `Campaign Config JSON`).
4. **Verification:** Tab 13 (`Agent Settings`) automatically displays `Configured in Script Properties (Protected)` for `CLIENT_{NAME}_GMAIL_RELAY_URL` once both `_URL` and `_SECRET` are present.

---

## Microsoft 365 / Exchange Online Setup for Clients (Direct Microsoft Graph API)

1. **Client-Side Entra ID Setup (`entra.microsoft.com`):** Client's M365 Global/App Admin registers a Single-Tenant App Registration, creates a Client Secret (copying the **Value**, not Secret ID), and grants **Microsoft Graph Application permissions** `Mail.ReadWrite` (Drafts + NDR bounce & reply reading) and `Mail.Send` (Direct Send) with Admin Consent *(or restricts to the single mailbox via Exchange Online RBAC for Applications; optionally adds `ExchangeMessageTrace.Read.All` for raw Exchange server-level delivery logs)*.
2. **Agency Master Binding:** In the Master AI Prospecting Studio **Script Properties**, add:
   * `CLIENT_{CLIENT_NAME}_MS_TENANT_ID` = Directory (tenant) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_ID` = Application (client) ID
   * `CLIENT_{CLIENT_NAME}_MS_CLIENT_SECRET` = Client Secret Value
3. **Campaign Identity Binding:** In the Control Panel sidebar (Section 5), select **`Microsoft 365 (Drafts Only)`** or **`Microsoft 365 (Direct Send)`**, leave `Email Sender Endpoint` blank, enter the client's Exchange Online **Primary Mailbox / UPN** (e.g., `sales@evaluater.ai`) in `Sender Email / Mailbox`, and set `Reply-To Email Address`.
4. **Verification:** Tab 13 (`Agent Settings`) automatically displays `Configured in Script Properties (Protected)` for `CLIENT_{NAME}_MS_TENANT_ID`, `_MS_CLIENT_ID`, and `_MS_CLIENT_SECRET`.

---

## Confirmed Workbook Tab Order & Schema (15 Total Tabs)

1. **`Dashboard`** *(Cols A–G)*
   * **Rows 1–2:** Green Banner (`A1:D1`), Active Campaign Name (`C1`), Status (`D1`), Active Campaign ID (`B2`).
   * **Rows 3–13:** 10-Row KPI Matrix (`A3:D13`) with Operating Department (`Client Owner` / `Local Owner`), live pipeline metrics, and client-scoped AI Key connection status.
   * **Rows 16–17+:** `Campaign Performance Rollup` (`A17:G...`) **auto-filtered on Col C (`Department`) by Selected Client / Owner**.
2. **`EmailEvents`** *(11 Canonical Cols — Auto-Filtered by Active Campaign ID)*
   * Centralized deliverability event ledger (`event_id`, `client_id`, `campaign_id`, `provider`, `provider_message_id`, `event_type`, `event_date`, `recipient`, `smtp_code`, `bounce_type`, `reason`). Strict validations: Col D (`Gmail`, `Microsoft 365`), Col F (`Sent`, `Delivered`, `Bounced`, `Replied`), Col J (`Hard`, `Soft`). Col G formatted `yyyy-mm-dd hh:mm:ss`; Col I (`smtp_code`) formatted plain text `@`.
3. **`Email Deliverability`** *(Dynamic Analytics & Charts Tab)*
   * **Filter Bar (Row 3):** `B3` Client Filter (`All Clients` + registered clients), `D3` Campaign Filter (`All Campaigns` + campaigns), `F3:J3` Active Scope.
   * **Headline Tiles (Rows 6–7):** `Sent` (`A7`), `Delivered` (`C7`), `Hard Bounces` (`E7`), `Soft Bounces` (`G7`), `Replies` (`I7`).
   * **Performance Breakdown (Rows 10–11):** `Bounce Rate` (`A11`), `Delivery Rate` (`D11`), `Reply Rate` (`G11`), formatted `0.0%`.
   * **Breakdown Tables & Charts (`Cols A–J` + Charts in `Col L`):**
     * **Campaign Comparison (`A13:J29` + Column Chart `L13`):** Per-campaign `Sent`, `Delivered`, `Hard Bounces`, `Soft Bounces`, `Replied`, `Client`, `Delivery Rate`, `Bounce Rate`, `Reply Rate`.
     * **SMTP Error Distribution (`A31:E47` + Donut Chart `L31`):** `QUERY`/`FILTER` grouping `Bounced` events by `smtp_code` with `Occurrences`, `Bounce Type`, `% of Bounces`, and `Sample Reason`.
     * **Daily Sending Volume (`A49:E80` + Column Chart `L49`):** Last 30 active days (`Sent`, `Delivered`, `Bounced`, `Replied`).
4. **`API_Usage_Log`** *(10 Canonical Cols — Auto-Filtered by Active Campaign ID)*
   * `Timestamp`, `Campaign ID`, `Service Name`, `Operation`, `HTTP Status`, `Input Tokens / Units`, `Output Tokens`, `Est. Cost ($)`, `Status`, `Details / Error` (tagged `[Client: {Name}]`).
5. **`Campaigns`** *(Locked 28 Cols — Auto-Filtered by Client / Owner Col C)*
   * Master campaign registry. Col C = `Department` (dynamic Client dropdown from `Lists & Settings`), Col D = `Status`, Cols H & J–M formatted as plain integer (`0`), Col V = `Email Sender Provider` (9 options including Microsoft 365), Col AA = `Campaign Config JSON` (extended attributes including `senderEmail`, `replyToEmail`, `presetNeed`).
6. **`Company Intelligence`** *(25 Cols — Auto-Filtered by Active Campaign ID)*
   * Account enrichment profiles. Col J = 6 allowed `ESG Reporting Maturity` states; Col Q = `eValuater Solution Fit` (`High`, `Medium`, `Low`, `Disqualified - Direct Competitor`).
7. **`Persona Leads`** *(29 Cols — Auto-Filtered by Active Campaign ID)*
   * Decision-maker profiles and email intelligence. Col Y = `Email Verification Status` (`valid`, `risky`, `invalid`, `catch-all`, `unverified`).
8. **`Persona Mapper Settings`** *(Locked 39 Canonical Cols — Auto-Filtered by Active Campaign ID Col A)*
   * Agent 2 parameters & scoring weights (Row 2+ direct write; Cols I, Z–AB, AI formatted as `0`; weight cols formatted as `0%`).
9. **`Prospect Review Queue`** *(45 Cols — Auto-Filtered by Active Campaign ID)*
   * Human-in-the-loop outreach staging. Col R `Review Status`, Col V `Human Approval Status`, Col Y `Send Status` (`Queued`, `Draft Created in Gmail`, `Draft Created`, `Sent`, `Pushed to Sequencer`, `Suppressed - Not Sent`, `Draft Failed`), Col AA `Dispatch Mode`, Col AB `Thread / Message ID`.
10. **`Campaign Configuration`** *(Locked 38 Vertical Parameters, Rows 2–39)*
    * Active campaign configuration key-value view. Row 2 = `CAMPAIGN_ID`, Row 4 = `DEPARTMENT` (dynamic Client dropdown), Rows 23–24 = `EMAIL_SENDER_PROVIDER` / `DISPATCH_MODE` (9 validated options including Microsoft 365), plain integer format (`0`) on numeric parameters (`B8`, `B10:B13`).
11. **`Company Database`** *(10 Cols — Auto-Filtered by Active Campaign ID Col J)*
    * Master discovered account directory (Cols A–I standard firmographics + Col J = `Campaign ID` for Scenario B multi-campaign deduplication).
12. **`Research Evidence`** *(9 Canonical Cols — Auto-Filtered by Active Campaign ID)*
    * Grounded audit trail (`Evidence ID`, `Campaign ID`, `Entity Name`, `Entity Type` (`Company` / `Persona`), `Search Query Used`, `Grounding Evidence URL`, `Extracted Snippet / Signal`, `Verification Score`, `Date Captured`).
13. **`Agent Settings`** *(4 Cols — Auto-Filtered by Active Client / Owner Col A)*
    * `Category`, `Setting Key`, `Value`, `Status / Description`. Holds `Growth Target` defaults (including `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`) + auto-generated `"Client: {Name}"` blocks (12 rows per client including `CLIENT_{NAME}_MS_*`) with protected credential badges and editable per-client daily spend caps / min verifier credits.
14. **`Execution Logs`** *(6 Cols — Auto-Filtered by Active Campaign ID)*
    * Timestamped pipeline execution audits (`Timestamp`, `Campaign ID`, `Action`, `Target Lead`, `Status`, `Details`).
15. **`Lists & Settings`** *(3 Cols: `List name`, `Value`, `Active status`)*
    * Client registry (`Client - Account` pinned cleanly at the top: `Growth Target` -> Active A–Z -> Inactive A–Z) + Whitelist and Suppression rules (`Suppression - Domain`, `Suppression - Email`, `Suppression - Competitor Domain`, `Suppression - Opt-Out Email`, `Suppression - Internal Exclusion`).

---

## Email Performance & Deliverability Workflow

    [1. Dispatch (`EmailSender.js`)]
       ├── Growth Target Gmail: `createDraft().send()` -> captures `Message-ID` / `draftId`
       ├── External Client Gmail: `ClientGmailRelay.gs` v2.2 -> returns `rfcMessageId` / `draftId`
       └── Microsoft 365: Graph `/messages` (`ImmutableId`) + `/send` -> captures `internetMessageId`
               │
               ├──> Updates Tab 9 (`Prospect Review Queue`: `Send Status`, `Dispatch Mode`, `Thread / Message ID`)
               └──> Writes `'Sent'` event immediately to Tab 2 (`EmailEvents`) via `recordEmailEvent_()`

    [2. Automated Hourly Trigger / Manual Sync (`EmailEventsDeliverability.js`)]
       │   (`syncEmailDeliverabilityEvents` or silent `syncEmailDeliverabilityFromSidebar`)
       ▼
       Acquires `ScriptLock` ─> Groups active mailboxes by `(clientName, provider, mailbox)`
       ├── Draft Promotion: Checks `Draft Created` / `Draft Created in Gmail` IDs; promotes manual sends to `'Sent'`
       ├── Microsoft 365 Poll: Optional `messageTraces` + `/messages` NDRs (`Undeliverable:`) & replies (`CLIENT_{NAME}_MS_*`)
       ├── Local Gmail Poll: `GmailApp.search` for `mailer-daemon`/`postmaster` DSNs & tracked-recipient replies
       ├── Client Relay Poll: `POST action:'pollEvents'` to `CLIENT_{NAME}_GMAIL_RELAY_URL` (`RELAY_SECRET` auth)
       └── Delivery Reconciliation:
             • Logs `'Bounced'` (`Hard` / `Soft` + SMTP code + reason) & `'Replied'`
             • Applies 4-hour no-bounce grace window (`Delivered` `250 2.0.0`), skipping already-delivered IDs
             • Removes superseded assumed `Delivered` rows if a delayed bounce arrives later

    [3. Analytics & Filtering]
       ├── Tab 3 (`Email Deliverability`): `B3` (Client) & `D3` (Campaign) drive live KPIs, rates, tables & 3 charts
       └── `ControlPanel.html` (Section 6): Queries `EmailEvents` via `queryEmailDeliverabilityMetrics()` using existing filters

### Scale & Apps Script Quota Considerations
* **6-Minute Execution Limit:** Sync batches read/write operations and caps polling to the last 7 days (`<= 300` Graph messages / `<= 50` Gmail threads per mailbox). At high volume (>25 active client mailboxes), split sync triggers per client group.
* **Daily Quotas:** Google Workspace allows 20,000 `UrlFetchApp` calls/day (consumer accounts allow 2,000/day).
* **Delivery Accuracy:** Without optional `ExchangeMessageTrace.Read.All` (M365) or Workspace Admin Email Log Search, `Delivered` represents an absence-of-bounce confirmation after 4 hours (`250 2.0.0`) rather than a recipient server receipt.

---

## System Architecture Data Flow

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
                     ▼
        ┌────────────────────────────────────────────────────────────────────┐
        │                 EmailEventsDeliverability.js                       │
        ├────────────────────────────────────────────────────────────────────┤
        │ • Inline Sent logging -> Tab 2 (EmailEvents)                       │
        │ • Hourly Trigger / Sidebar Sync -> Draft Promotion, Bounces,       │
        │   Replies & 4h Delivery Confirmation (strictly client-isolated)    │
        │ • Live Filters & Charts -> Tab 3 (Email Deliverability)            │
        │ • Sidebar Query API -> ControlPanel.html Section 6 Snapshot        │
        └────────────────────────────────────────────────────────────────────┘
                     │
                     ▼
        ┌───────────────────────┐
        │  Dashboard & Sheets   │ ──> Updates live KPIs, Client-Filtered Performance Rollup,
        └───────────────────────┘     Send Status, Execution Logs & Verifier Low-Balance Alerts.