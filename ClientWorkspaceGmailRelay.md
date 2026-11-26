How to Apply (For Each External Google Workspace Client Using Gmail Relay)

1. Open the client's standalone Relay project at https://script.google.com (inside the client's Google Workspace account).

2\. Replace all code in Code.gs with the complete script above and click Save.

3\. Run authorizeRelayPermissions once from the top toolbar if Apps Script asks to verify scopes.

4\. Click Deploy → Manage deployments → Edit (pencil icon) → Version: New version → Deploy (this preserves the exact same /exec Web App URL stored in your master Script Properties).

5\. (Security Reminder) Keep RELAY\_SECRET strictly inside Script Properties (never in spreadsheet cells), and ensure Editor access on your master Google Sheet is restricted strictly to trusted agency admins.





/\*\*

&#x20;\* ============================================================================

&#x20;\* CLIENT GOOGLE WORKSPACE GMAIL RELAY (STANDALONE WEB APP v2.3)

&#x20;\* ============================================================================

&#x20;\* Deploy this script INSIDE the client's own Google Workspace account at:

&#x20;\* https://script.google.com

&#x20;\*

&#x20;\* CLIENT: ClientName (example.com)

&#x20;\* DEFAULT OUTREACH: FirstName.LastName@example.com

&#x20;\* DEFAULT REPLY-TO: FirstName.LastName@example.com

&#x20;\*

&#x20;\* WHAT IT DOES:

&#x20;\* Supports both campaign-level dispatch modes chosen by the client:

&#x20;\* - Drafts Only ("createDraft"): Stages Gmail drafts and detects manual

&#x20;\*   Draft-to-Sent promotions via GmailApp.getDrafts() lookup.

&#x20;\* - Direct Send ("sendEmail"): Sends immediately from the client's Workspace

&#x20;\*   mailbox and returns draftId, messageId, threadId, and RFC Message-ID.

&#x20;\* - Deliverability Event Polling ("pollEvents"): Reports manual draft

&#x20;\*   promotions, verified SMTP DSN bounces (Hard/Soft), and replies strictly

&#x20;\*   from tracked campaign recipients (authenticated via RELAY\_SECRET).

&#x20;\*

&#x20;\* SETUP / UPGRADE STEPS:

&#x20;\* 1. Log in to the client's Google Workspace account (ideally FirstName.LastName@example.com).

&#x20;\* 2. Open the standalone project at https://script.google.com

&#x20;\* 3. Replace all code in Code.gs with this complete script and click Save.

&#x20;\* 4. Confirm Project Settings (gear icon) > Script Properties contains:

&#x20;\*      Property: RELAY\_SECRET

&#x20;\*      Value:    <strong random passphrase, minimum 16 characters (32+ recommended)>

&#x20;\* 5. Select "authorizeRelayPermissions" in the top toolbar and click Run once

&#x20;\*    if Apps Script asks to verify scopes.

&#x20;\* 6. Deploy:

&#x20;\*      - First-time setup: Deploy > New deployment > Web app:

&#x20;\*          \* Execute as: Me (the client's Google Workspace email)

&#x20;\*          \* Who has access: Anyone

&#x20;\*      - Upgrading an existing relay (keeps the exact same /exec URL):

&#x20;\*          Deploy > Manage deployments > Edit (pencil icon) > Version: New version > Deploy

&#x20;\* 7. Ensure YOUR main AI Prospecting Studio Script Properties contain:

&#x20;\*      - CLIENT\_{CLIENT\_NAME}\_GMAIL\_RELAY\_URL    = <Web App /exec URL>

&#x20;\*      - CLIENT\_{CLIENT\_NAME}\_GMAIL\_RELAY\_SECRET = <Same RELAY\_SECRET>

&#x20;\* ============================================================================

&#x20;\*/



function doGet() {

&#x20; return ContentService

&#x20;   .createTextOutput(JSON.stringify({

&#x20;     ok: true,

&#x20;     service: 'Client Gmail Relay (v2.3)',

&#x20;     mailbox: getRelayMailboxEmail\_(),

&#x20;     message: 'Web App is live. Use HTTP POST with a valid RELAY\_SECRET.'

&#x20;   }))

&#x20;   .setMimeType(ContentService.MimeType.JSON);

}



function doPost(e) {

&#x20; try {

&#x20;   if (!e || !e.postData || !e.postData.contents) {

&#x20;     return jsonResponse\_({ success: false, error: 'Missing JSON request body.' });

&#x20;   }



&#x20;   const expectedSecret = String(

&#x20;     PropertiesService.getScriptProperties().getProperty('RELAY\_SECRET') || ''

&#x20;   ).trim();



&#x20;   if (!expectedSecret || expectedSecret.length < 16) {

&#x20;     return jsonResponse\_({

&#x20;       success: false,

&#x20;       error: 'RELAY\_SECRET is not configured in client Script Properties (minimum 16 characters).'

&#x20;     });

&#x20;   }



&#x20;   const payload = JSON.parse(e.postData.contents);

&#x20;   const providedSecret = String(payload.secret || '').trim();



&#x20;   if (!providedSecret || providedSecret !== expectedSecret) {

&#x20;     return jsonResponse\_({ success: false, error: 'Unauthorized: Invalid RELAY\_SECRET.' });

&#x20;   }



&#x20;   const action = String(payload.action || 'createDraft').trim();

&#x20;   const primaryMailbox = getRelayMailboxEmail\_();

&#x20;   const ownInfo = getRelayOwnAddressesInfo\_(primaryMailbox);



&#x20;   // Deliverability Event Polling (v2.3)

&#x20;   if (action === 'pollEvents') {

&#x20;     return jsonResponse\_(handleRelayPollEvents\_(payload, primaryMailbox, ownInfo));

&#x20;   }



&#x20;   const to = String(payload.to || '').trim();

&#x20;   const subject = String(payload.subject || '').trim();

&#x20;   const body = String(payload.body || '');

&#x20;   const senderName = String(payload.senderName || '').trim();

&#x20;   const senderEmail = String(payload.senderEmail || payload.from || '').trim();

&#x20;   const replyTo = String(payload.replyTo || senderEmail || '').trim();



&#x20;   if (!to || to.indexOf('@') < 1) {

&#x20;     return jsonResponse\_({ success: false, error: 'Missing or invalid recipient email ("to").' });

&#x20;   }

&#x20;   if (!subject || !body) {

&#x20;     return jsonResponse\_({ success: false, error: 'Both "subject" and "body" are required.' });

&#x20;   }



&#x20;   const resolved = resolveRelayMailOptions\_(senderName, senderEmail, replyTo, primaryMailbox, ownInfo);

&#x20;   const mailOptions = resolved.mailOptions;

&#x20;   let effectiveSenderEmail = resolved.effectiveSenderEmail;

&#x20;   let aliasFallback = resolved.aliasFallback;



&#x20;   if (action === 'createDraft') {

&#x20;     let draft;

&#x20;     try {

&#x20;       draft = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;     } catch (aliasErr) {

&#x20;       // Fallback without 'from' if senderEmail is an unverified alias on this Workspace account

&#x20;       if (mailOptions.from) {

&#x20;         delete mailOptions.from;

&#x20;         effectiveSenderEmail = primaryMailbox;

&#x20;         aliasFallback = true;

&#x20;         draft = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;       } else {

&#x20;         throw aliasErr;

&#x20;       }

&#x20;     }



&#x20;     return jsonResponse\_({

&#x20;       success: true,

&#x20;       action: 'createDraft',

&#x20;       draftId: draft.getId(),

&#x20;       mailbox: primaryMailbox,

&#x20;       senderEmail: effectiveSenderEmail,

&#x20;       replyTo: mailOptions.replyTo || primaryMailbox,

&#x20;       senderName: senderName || 'Default',

&#x20;       aliasFallback: aliasFallback

&#x20;     });

&#x20;   }



&#x20;   if (action === 'sendEmail') {

&#x20;     let draftToSend = null;

&#x20;     let sentMsg = null;



&#x20;     try {

&#x20;       draftToSend = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;       sentMsg = draftToSend.send();

&#x20;     } catch (sendAliasErr) {

&#x20;       if (mailOptions.from) {

&#x20;         // Clean up any unsent draft before retrying with the primary mailbox

&#x20;         if (draftToSend) {

&#x20;           try { draftToSend.deleteDraft(); } catch (delErr) {}

&#x20;         }

&#x20;         delete mailOptions.from;

&#x20;         effectiveSenderEmail = primaryMailbox;

&#x20;         aliasFallback = true;

&#x20;         draftToSend = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;         sentMsg = draftToSend.send();

&#x20;       } else {

&#x20;         throw sendAliasErr;

&#x20;       }

&#x20;     }



&#x20;     const draftId = draftToSend ? draftToSend.getId() : '';

&#x20;     let messageId = '';

&#x20;     let threadId = '';

&#x20;     let rfcMessageId = '';



&#x20;     try {

&#x20;       if (sentMsg) {

&#x20;         messageId = sentMsg.getId() || '';

&#x20;         threadId = sentMsg.getThread() ? sentMsg.getThread().getId() : '';

&#x20;         rfcMessageId = sentMsg.getHeader('Message-ID') || '';

&#x20;       }

&#x20;     } catch (idErr) {}



&#x20;     if (!rfcMessageId \&\& messageId) {

&#x20;       try {

&#x20;         const fetchedMsg = GmailApp.getMessageById(messageId);

&#x20;         if (fetchedMsg) {

&#x20;           rfcMessageId = fetchedMsg.getHeader('Message-ID') || '';

&#x20;         }

&#x20;       } catch (fetchErr) {}

&#x20;     }



&#x20;     return jsonResponse\_({

&#x20;       success: true,

&#x20;       action: 'sendEmail',

&#x20;       draftId: draftId,

&#x20;       messageId: messageId,

&#x20;       threadId: threadId,

&#x20;       rfcMessageId: rfcMessageId,

&#x20;       mailbox: primaryMailbox,

&#x20;       senderEmail: effectiveSenderEmail,

&#x20;       replyTo: mailOptions.replyTo || primaryMailbox,

&#x20;       senderName: senderName || 'Default',

&#x20;       aliasFallback: aliasFallback

&#x20;     });

&#x20;   }



&#x20;   return jsonResponse\_({ success: false, error: 'Unsupported action: ' + action });

&#x20; } catch (err) {

&#x20;   return jsonResponse\_({ success: false, error: err \&\& err.message ? err.message : String(err) });

&#x20; }

}



/\*\*

&#x20;\* Resolves mailbox \& verified Send-As aliases so `mailOptions.from` never causes

&#x20;\* unexpected draft or Direct Send failures on unverified aliases.

&#x20;\*/

function getRelayOwnAddressesInfo\_(primaryMailbox) {

&#x20; const primaryLower = String(primaryMailbox || '').trim().toLowerCase();

&#x20; const ownMap = {};

&#x20; const aliasMap = {};

&#x20; let aliasesChecked = false;



&#x20; if (primaryLower \&\& primaryLower.indexOf('@') > 0) {

&#x20;   ownMap\[primaryLower] = true;

&#x20; }



&#x20; try {

&#x20;   const aliases = GmailApp.getAliases() || \[];

&#x20;   aliasesChecked = true;

&#x20;   aliases.forEach(function(alias) {

&#x20;     const trimmed = String(alias || '').trim();

&#x20;     const lower = trimmed.toLowerCase();

&#x20;     if (lower \&\& lower.indexOf('@') > 0) {

&#x20;       ownMap\[lower] = true;

&#x20;       aliasMap\[lower] = trimmed;

&#x20;     }

&#x20;   });

&#x20; } catch (aliasErr) {}



&#x20; return {

&#x20;   primaryLower: primaryLower,

&#x20;   ownMap: ownMap,

&#x20;   aliasMap: aliasMap,

&#x20;   aliasesChecked: aliasesChecked

&#x20; };

}



function resolveRelayMailOptions\_(senderName, senderEmail, replyTo, primaryMailbox, ownInfo) {

&#x20; const mailOptions = {};

&#x20; let effectiveSenderEmail = primaryMailbox;

&#x20; let aliasFallback = false;



&#x20; if (senderName) {

&#x20;   mailOptions.name = senderName;

&#x20; }

&#x20; if (replyTo \&\& replyTo.indexOf('@') > 0) {

&#x20;   mailOptions.replyTo = replyTo;

&#x20; }



&#x20; const cleanSender = String(senderEmail || '').trim();

&#x20; const cleanSenderLower = cleanSender.toLowerCase();



&#x20; if (cleanSender \&\& cleanSender.indexOf('@') > 0 \&\& cleanSenderLower !== ownInfo.primaryLower) {

&#x20;   if (ownInfo.aliasesChecked) {

&#x20;     if (ownInfo.aliasMap\[cleanSenderLower]) {

&#x20;       mailOptions.from = ownInfo.aliasMap\[cleanSenderLower];

&#x20;       effectiveSenderEmail = ownInfo.aliasMap\[cleanSenderLower];

&#x20;     } else {

&#x20;       // Sender email is not an active Send-As alias; safely fall back to primary mailbox

&#x20;       aliasFallback = true;

&#x20;       effectiveSenderEmail = primaryMailbox;

&#x20;     }

&#x20;   } else {

&#x20;     mailOptions.from = cleanSender;

&#x20;     effectiveSenderEmail = cleanSender;

&#x20;   }

&#x20; }



&#x20; return {

&#x20;   mailOptions: mailOptions,

&#x20;   effectiveSenderEmail: effectiveSenderEmail,

&#x20;   aliasFallback: aliasFallback

&#x20; };

}



/\*\*

&#x20;\* Polls the client's Gmail mailbox strictly for:

&#x20;\* 1. Tracked draftIds that have been manually sent by the user (using GmailApp.getDrafts() lookup).

&#x20;\* 2. Verified DSN/NDR bounce notices (skipping outbound messages in the same thread).

&#x20;\* 3. Replies strictly from tracked campaign recipient addresses.

&#x20;\*/

function handleRelayPollEvents\_(payload, primaryMailbox, ownInfo) {

&#x20; const lookbackDays = Math.min(Math.max(parseInt(payload.lookbackDays, 10) || 7, 1), 30);

&#x20; const trackedDrafts = Array.isArray(payload.trackedDrafts) ? payload.trackedDrafts.slice(0, 50) : \[];

&#x20; const trackedRecipients = Array.isArray(payload.trackedRecipients) ? payload.trackedRecipients : \[];



&#x20; const recipientSet = {};

&#x20; trackedRecipients.forEach(function(email) {

&#x20;   const clean = sanitizeRelaySearchEmail\_(email);

&#x20;   if (clean \&\& clean.indexOf('@') > 0 \&\& !ownInfo.ownMap\[clean]) {

&#x20;     recipientSet\[clean] = true;

&#x20;   }

&#x20; });



&#x20; const promotedDrafts = \[];

&#x20; const bounces = \[];

&#x20; const replies = \[];

&#x20; const seenBounceKeys = {};



&#x20; // 1. Check tracked drafts for manual sends using GmailApp.getDrafts() lookup

&#x20; //    (GmailApp.getDraft(id) throws once a draft is sent, which previously prevented promotion)

&#x20; if (trackedDrafts.length > 0) {

&#x20;   const activeDraftIdMap = {};

&#x20;   let draftLookupOk = false;



&#x20;   try {

&#x20;     const allDrafts = GmailApp.getDrafts() || \[];

&#x20;     for (let d = 0; d < allDrafts.length; d++) {

&#x20;       const id = allDrafts\[d].getId();

&#x20;       if (id) activeDraftIdMap\[id] = true;

&#x20;     }

&#x20;     draftLookupOk = true;

&#x20;   } catch (draftListErr) {

&#x20;     draftLookupOk = false;

&#x20;   }



&#x20;   if (draftLookupOk) {

&#x20;     trackedDrafts.forEach(function(draftItem) {

&#x20;       const draftId = String((draftItem \&\& draftItem.draftId) || '').trim();

&#x20;       const recip = sanitizeRelaySearchEmail\_(draftItem \&\& draftItem.email);

&#x20;       const campId = String((draftItem \&\& draftItem.campaignId) || '').trim();

&#x20;       if (!draftId || !recip || activeDraftIdMap\[draftId]) return;



&#x20;       let minSentMs = 0;

&#x20;       const createdRaw = draftItem \&\& (draftItem.createdAt || draftItem.draftCreatedAt);

&#x20;       if (createdRaw) {

&#x20;         const parsedCreated = new Date(createdRaw).getTime();

&#x20;         if (!isNaN(parsedCreated)) {

&#x20;           minSentMs = parsedCreated - 300000; // 5-minute clock skew allowance

&#x20;         }

&#x20;       }



&#x20;       try {

&#x20;         const sentThreads = GmailApp.search('in:sent to:' + recip + ' newer\_than:' + lookbackDays + 'd', 0, 5);

&#x20;         const sentMsg = findMatchingSentOutboundMessage\_(sentThreads, recip, ownInfo.ownMap, minSentMs);

&#x20;         if (sentMsg) {

&#x20;           let rfcId = '';

&#x20;           let msgId = '';

&#x20;           let sentDate = new Date();



&#x20;           msgId = sentMsg.getId() || '';

&#x20;           sentDate = sentMsg.getDate() || new Date();

&#x20;           try { rfcId = sentMsg.getHeader('Message-ID') || ''; } catch (e) {}



&#x20;           promotedDrafts.push({

&#x20;             draftId: draftId,

&#x20;             campaignId: campId,

&#x20;             recipient: recip,

&#x20;             messageId: rfcId || msgId || draftId,

&#x20;             eventDate: sentDate.toISOString()

&#x20;           });

&#x20;           recipientSet\[recip] = true;

&#x20;         }

&#x20;       } catch (sentSearchErr) {}

&#x20;     });

&#x20;   }

&#x20; }



&#x20; // 2. Scan verified DSN/NDR bounce notices (filtering out outbound emails in the same thread)

&#x20; const bounceThreads = GmailApp.search('from:(mailer-daemon OR postmaster) newer\_than:' + lookbackDays + 'd', 0, 50);

&#x20; bounceThreads.forEach(function(thread) {

&#x20;   const msgs = thread.getMessages();

&#x20;   msgs.forEach(function(msg) {

&#x20;     try {

&#x20;       if (msg.isDraft()) return;

&#x20;     } catch (e) {}



&#x20;     const msgFrom = extractRelayEmailAddress\_(msg.getFrom());

&#x20;     if (msgFrom \&\& ownInfo.ownMap\[msgFrom]) return; // Never parse our own outbound message as a bounce



&#x20;     const subject = String(msg.getSubject() || '');

&#x20;     const plainBody = String(msg.getPlainBody() || '');

&#x20;     let raw = '';

&#x20;     try { raw = msg.getRawContent() || ''; } catch (e) { raw = plainBody; }



&#x20;     if (!isRelayDsnBounceMessage\_(msgFrom, subject, plainBody, raw)) return;



&#x20;     const parsed = parseRelaySmtpBounceDetails\_(subject, plainBody, raw, ownInfo.ownMap);

&#x20;     if (!parsed.recipient) return;



&#x20;     const hasDiagnosticSignal = parsed.smtpCode !== 'UNKNOWN' ||

&#x20;       /Final-Recipient:|X-Failed-Recipients:|Diagnostic-Code:|report-type=delivery-status/i.test(raw);

&#x20;     if (!hasDiagnosticSignal) return;



&#x20;     const msgId = msg.getId() || '';

&#x20;     const dedupeKey = (msgId || parsed.originalMessageId || '') + '|' + parsed.recipient + '|' + parsed.smtpCode;

&#x20;     if (seenBounceKeys\[dedupeKey]) return;

&#x20;     seenBounceKeys\[dedupeKey] = true;



&#x20;     bounces.push({

&#x20;       messageId: msgId,

&#x20;       originalMessageId: parsed.originalMessageId,

&#x20;       recipient: parsed.recipient,

&#x20;       smtpCode: parsed.smtpCode,

&#x20;       bounceType: parsed.bounceType,

&#x20;       reason: parsed.reason,

&#x20;       eventDate: (msg.getDate() || new Date()).toISOString()

&#x20;     });

&#x20;   });

&#x20; });



&#x20; // 3. Scan replies strictly matching tracked campaign recipients

&#x20; const recipList = Object.keys(recipientSet);

&#x20; for (let idx = 0; idx < recipList.length; idx += 15) {

&#x20;   const chunk = recipList.slice(idx, idx + 15);

&#x20;   const query = 'in:anywhere newer\_than:' + lookbackDays + 'd -from:mailer-daemon -from:postmaster (' +

&#x20;     chunk.map(function(addr) { return 'from:' + addr; }).join(' OR ') + ')';

&#x20;   const replyThreads = GmailApp.search(query, 0, 50);



&#x20;   replyThreads.forEach(function(thread) {

&#x20;     const msgs = thread.getMessages();

&#x20;     msgs.forEach(function(msg) {

&#x20;       try {

&#x20;         if (msg.isDraft()) return;

&#x20;       } catch (e) {}



&#x20;       const sender = extractRelayEmailAddress\_(msg.getFrom());

&#x20;       if (!sender || ownInfo.ownMap\[sender] || !recipientSet\[sender]) return;



&#x20;       const subj = String(msg.getSubject() || '');

&#x20;       const plainBody = String(msg.getPlainBody() || '');

&#x20;       if (isRelayDsnBounceMessage\_(sender, subj, plainBody, '')) return;



&#x20;       let rfcId = '';

&#x20;       try { rfcId = msg.getHeader('Message-ID') || ''; } catch (e) {}

&#x20;       const cleanSubj = subj.slice(0, 120);



&#x20;       replies.push({

&#x20;         messageId: rfcId || msg.getId() || '',

&#x20;         recipient: sender,

&#x20;         subject: cleanSubj,

&#x20;         reason: 'Reply received: ' + cleanSubj,

&#x20;         eventDate: (msg.getDate() || new Date()).toISOString()

&#x20;       });

&#x20;     });

&#x20;   });

&#x20; }



&#x20; return {

&#x20;   success: true,

&#x20;   action: 'pollEvents',

&#x20;   mailbox: primaryMailbox,

&#x20;   promotedDrafts: promotedDrafts,

&#x20;   bounces: bounces,

&#x20;   replies: replies

&#x20; };

}



/\*\*

&#x20;\* Finds the newest non-draft outbound message sent to `recip`, ensuring that

&#x20;\* prospect replies in the same thread are never mistaken for the sent message.

&#x20;\*/

function findMatchingSentOutboundMessage\_(sentThreads, recip, ownMap, minSentMs) {

&#x20; if (!sentThreads || sentThreads.length === 0) return null;



&#x20; for (let t = 0; t < sentThreads.length; t++) {

&#x20;   const msgs = sentThreads\[t].getMessages() || \[];

&#x20;   for (let m = msgs.length - 1; m >= 0; m--) {

&#x20;     const msg = msgs\[m];

&#x20;     try {

&#x20;       if (msg.isDraft()) continue;

&#x20;     } catch (e) {}



&#x20;     const fromAddr = extractRelayEmailAddress\_(msg.getFrom());

&#x20;     if (fromAddr === recip) continue;

&#x20;     if (Object.keys(ownMap).length > 0 \&\& fromAddr \&\& !ownMap\[fromAddr]) continue;



&#x20;     const toField = (String(msg.getTo() || '') + ',' + String(msg.getCc() || '')).toLowerCase();

&#x20;     if (toField.indexOf(recip) < 0) continue;



&#x20;     if (minSentMs > 0) {

&#x20;       const msgTime = msg.getDate() ? msg.getDate().getTime() : 0;

&#x20;       if (msgTime \&\& msgTime < minSentMs) continue;

&#x20;     }



&#x20;     return msg;

&#x20;   }

&#x20; }



&#x20; return null;

}



/\*\*

&#x20;\* Verifies a message is a genuine DSN/NDR delivery failure notice.

&#x20;\*/

function isRelayDsnBounceMessage\_(fromAddr, subject, plainBody, rawContent) {

&#x20; const fromLow = String(fromAddr || '').toLowerCase();

&#x20; const subjLow = String(subject || '').toLowerCase();

&#x20; const rawStr = String(rawContent || '');



&#x20; const isDsnSender = /(?:mailer-daemon|postmaster|mail-daemon|bounces?@|return@)/i.test(fromLow);

&#x20; const hasDsnSubject = /(?:delivery status notification|mail delivery subsystem|undeliverable|returned mail|failure notice|delivery has failed|couldn't be delivered|could not be delivered|message not delivered|address not found|undelivered mail)/i.test(subjLow);

&#x20; const hasDsnHeaders = /(?:Final-Recipient:|X-Failed-Recipients:|report-type=delivery-status)/i.test(rawStr);



&#x20; if (isDsnSender \&\& (hasDsnSubject || hasDsnHeaders || /5\\d{2}|4\\d{2}|Diagnostic-Code:/i.test(plainBody || rawStr))) {

&#x20;   return true;

&#x20; }

&#x20; if (hasDsnHeaders \&\& hasDsnSubject) {

&#x20;   return true;

&#x20; }

&#x20; return false;

}



function parseRelaySmtpBounceDetails\_(subject, plainBody, rawContent, ownMap) {

&#x20; const plainStr = String(plainBody || '');

&#x20; const rawStr = String(rawContent || '');

&#x20; const combined = plainStr + '\\n' + rawStr;

&#x20; const selfMap = ownMap || {};



&#x20; // Isolate embedded original message section (message/rfc822 or "Original message")

&#x20; // so outer DSN headers (like the DSN's own Message-ID and To:) are not misread.

&#x20; let embeddedRfc822 = '';

&#x20; const rfcSplit = rawStr.split(/Content-Type:\\s\*message\\/rfc822\[^\\r\\n]\*|---+\\s\*Original message\\s\*---+/i);

&#x20; if (rfcSplit \&\& rfcSplit.length > 1) {

&#x20;   embeddedRfc822 = rfcSplit.slice(1).join('\\n');

&#x20; }



&#x20; // 1. Extract bounced recipient (never matching mailer-daemon, postmaster, or own addresses)

&#x20; let recipient = '';

&#x20; const recipPatterns = \[

&#x20;   /Final-Recipient:\\s\*(?:rfc822;)?\\s\*(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i,

&#x20;   /Original-Recipient:\\s\*(?:rfc822;)?\\s\*(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i,

&#x20;   /X-Failed-Recipients:\\s\*(\[^\\s<>;"',\\r\\n]+@\[^\\s<>;"',\\r\\n]+)/i,

&#x20;   /(?:delivery to the following recipient failed permanently|wasn't delivered to|undeliverable to|could not be delivered to|address not found)\[:\\s]+(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i

&#x20; ];



&#x20; for (let i = 0; i < recipPatterns.length; i++) {

&#x20;   const m = combined.match(recipPatterns\[i]);

&#x20;   if (m \&\& m\[1]) {

&#x20;     const candidate = m\[1].trim().toLowerCase().replace(/\[.,;>]+$/, '');

&#x20;     if (candidate.indexOf('mailer-daemon') < 0 \&\& candidate.indexOf('postmaster') < 0 \&\& !selfMap\[candidate]) {

&#x20;       recipient = candidate;

&#x20;       break;

&#x20;     }

&#x20;   }

&#x20; }



&#x20; if (!recipient) {

&#x20;   const fallbackSearchText = (embeddedRfc822 || plainStr);

&#x20;   const toMatch = fallbackSearchText.match(/^To:\\s\*(?:\[^<\\r\\n]\*<)?(\[^\\s<>;"',\\r\\n]+@\[^\\s<>;"',\\r\\n]+)/im);

&#x20;   if (toMatch \&\& toMatch\[1]) {

&#x20;     const candidate = toMatch\[1].trim().toLowerCase().replace(/\[.,;>]+$/, '');

&#x20;     if (candidate.indexOf('mailer-daemon') < 0 \&\& candidate.indexOf('postmaster') < 0 \&\& !selfMap\[candidate]) {

&#x20;       recipient = candidate;

&#x20;     }

&#x20;   }

&#x20; }



&#x20; // 2. Prioritize Original-Message-ID / In-Reply-To / embedded message/rfc822 Message-ID

&#x20; //    (Never grab the DSN wrapper's own top-level Message-ID header)

&#x20; let originalMessageId = '';

&#x20; const priorityIdMatch =

&#x20;   combined.match(/(?:In-Reply-To|Original-Message-ID|X-Original-Message-ID):\\s\*(<\[^>]+>)/i) ||

&#x20;   (embeddedRfc822 ? embeddedRfc822.match(/^Message-ID:\\s\*(<\[^>]+>)/im) : null) ||

&#x20;   plainStr.match(/Message-ID:\\s\*(<\[^>]+>)/i);



&#x20; if (priorityIdMatch \&\& priorityIdMatch\[1]) {

&#x20;   originalMessageId = priorityIdMatch\[1].trim();

&#x20; }



&#x20; // 3. Extract SMTP status code strictly from diagnostic/status sections or plainBody

&#x20; const diagnosticSection = (plainStr + '\\n' + (rawStr.match(/Diagnostic-Code:\[\\s\\S]{0,600}/i) || \[''])\[0]);

&#x20; let smtpCode = 'UNKNOWN';



&#x20; const smtpFullMatch =

&#x20;   diagnosticSection.match(/\\b(\[45]\\d{2})\[\\s-]+(\[45]\\.\\d{1,3}\\.\\d{1,3})\\b/) ||

&#x20;   combined.match(/\\b(\[45]\\d{2})\[\\s-]+(\[45]\\.\\d{1,3}\\.\\d{1,3})\\b/);



&#x20; if (smtpFullMatch) {

&#x20;   smtpCode = smtpFullMatch\[1] + ' ' + smtpFullMatch\[2];

&#x20; } else {

&#x20;   const statusMatch = combined.match(/^Status:\\s\*(\[45]\\.\\d{1,3}\\.\\d{1,3})/im);

&#x20;   const explicitCodeMatch =

&#x20;     diagnosticSection.match(/(?:Diagnostic-Code:\[^\\r\\n]\*|Remote server returned\[^\\r\\n]\*|SMTP error\[^\\r\\n]\*|The response(?:\\s+from the remote server)?\\s+was:\[\\s\\S]{0,120}?)\\b(\[45]\\d{2})\\b/i) ||

&#x20;     plainStr.match(/\\b(\[45]\[0-5]\\d)\\b/);



&#x20;   if (explicitCodeMatch \&\& statusMatch) {

&#x20;     smtpCode = explicitCodeMatch\[1] + ' ' + statusMatch\[1];

&#x20;   } else if (explicitCodeMatch) {

&#x20;     smtpCode = explicitCodeMatch\[1];

&#x20;   } else if (statusMatch) {

&#x20;     smtpCode = statusMatch\[1];

&#x20;   }

&#x20; }



&#x20; // 4. Extract reason

&#x20; let reason = '';

&#x20; const diagMatch = combined.match(/Diagnostic-Code:\\s\*(?:smtp;)?\\s\*(\[^\\r\\n]+(?:\\r?\\n\[ \\t]+\[^\\r\\n]+)?)/i) ||

&#x20;   combined.match(/Remote server returned\\s\*'(\[^']+)'/i);

&#x20; if (diagMatch \&\& diagMatch\[1]) {

&#x20;   reason = diagMatch\[1].replace(/\\s+/g, ' ').trim();

&#x20; } else {

&#x20;   reason = String(subject || '').trim() || plainStr.replace(/\\s+/g, ' ').trim().slice(0, 220);

&#x20; }



&#x20; // 5. Classify Hard vs. Soft bounce

&#x20; const codeLead = smtpCode.charAt(0);

&#x20; const reasonLow = (reason + ' ' + String(subject || '')).toLowerCase();

&#x20; let bounceType = 'Soft';



&#x20; if (codeLead === '4' ||

&#x20;     smtpCode.indexOf('552') === 0 ||

&#x20;     smtpCode.indexOf('5.2.2') >= 0 ||

&#x20;     smtpCode.indexOf('5.4.7') >= 0 ||

&#x20;     reasonLow.indexOf('mailbox full') >= 0 ||

&#x20;     reasonLow.indexOf('over quota') >= 0 ||

&#x20;     reasonLow.indexOf('temporarily') >= 0 ||

&#x20;     reasonLow.indexOf('try again') >= 0 ||

&#x20;     reasonLow.indexOf('rate limit') >= 0 ||

&#x20;     reasonLow.indexOf('throttl') >= 0 ||

&#x20;     reasonLow.indexOf('greylist') >= 0) {

&#x20;   bounceType = 'Soft';

&#x20; } else if (codeLead === '5' ||

&#x20;     reasonLow.indexOf('user unknown') >= 0 ||

&#x20;     reasonLow.indexOf('does not exist') >= 0 ||

&#x20;     reasonLow.indexOf('recipientnotfound') >= 0 ||

&#x20;     reasonLow.indexOf('bad-mailbox') >= 0 ||

&#x20;     reasonLow.indexOf('no such user') >= 0 ||

&#x20;     reasonLow.indexOf('domain not found') >= 0 ||

&#x20;     reasonLow.indexOf('address not found') >= 0 ||

&#x20;     reasonLow.indexOf('permanent') >= 0) {

&#x20;   bounceType = 'Hard';

&#x20; }



&#x20; return {

&#x20;   recipient: recipient,

&#x20;   originalMessageId: originalMessageId,

&#x20;   smtpCode: smtpCode,

&#x20;   bounceType: bounceType,

&#x20;   reason: reason.slice(0, 350)

&#x20; };

}



function sanitizeRelaySearchEmail\_(email) {

&#x20; return String(email || '')

&#x20;   .trim()

&#x20;   .toLowerCase()

&#x20;   .replace(/\["'<>\\\\]/g, '');

}



function extractRelayEmailAddress\_(headerStr) {

&#x20; const s = String(headerStr || '').trim();

&#x20; const angle = s.match(/<(\[^>]+@\[^>]+)>/);

&#x20; if (angle \&\& angle\[1]) return angle\[1].trim().toLowerCase();

&#x20; const plain = s.match(/(\[^\\s<>;"']+@\[^\\s<>;"']+)/);

&#x20; return plain \&\& plain\[1] ? plain\[1].trim().toLowerCase() : '';

}



function authorizeRelayPermissions() {

&#x20; GmailApp.getDrafts();

&#x20; GmailApp.getAliases();

&#x20; GmailApp.search('in:sent', 0, 1);

&#x20; Logger.log('Gmail Draft, Alias, Send \& Deliverability Read permissions authorized for: ' + getRelayMailboxEmail\_());

}



function getRelayMailboxEmail\_() {

&#x20; try {

&#x20;   return Session.getActiveUser().getEmail() || 'Client Workspace';

&#x20; } catch (e) {

&#x20;   return 'Client Workspace';

&#x20; }

}



function jsonResponse\_(obj) {

&#x20; return ContentService

&#x20;   .createTextOutput(JSON.stringify(obj))

&#x20;   .setMimeType(ContentService.MimeType.JSON);

}

