How to Apply (For Each External Google Workspace Client Using Gmail Relay)

1. Open the client's standalone Relay project at https://script.google.com (inside the client's Google Workspace account).

2\. Replace all code in Code.gs with the complete code below and click Save.

3\. Run authorizeRelayPermissions once from the top toolbar if Apps Script asks to verify scopes.

4\. Important: Click Deploy → Manage deployments → Edit (pencil icon) → Version: New version → Deploy. 
(Choosing "New version" updates the live relay while keeping the exact same /exec Web App URL in your master Script Properties!)

5\. (Optional) Save this updated code in your ClientWorkspaceGmailRelay.md documentation file as well.


/\*\*

&#x20;\* ============================================================================

&#x20;\* CLIENT GOOGLE WORKSPACE GMAIL RELAY (STANDALONE WEB APP v2.2)

&#x20;\* ============================================================================

&#x20;\* Deploy this script INSIDE the client's own Google Workspace account at:

&#x20;\* https://script.google.com

&#x20;\*

&#x20;\* CLIENT: ClientName (example.com)

&#x20;\* DEFAULT OUTREACH: FirstName.LastName@example.com

&#x20;\* DEFAULT REPLY-TO: FirstName.LastName@example.com

&#x20;\*

&#x20;\* WHAT IT DOES:

&#x20;\* Allows your AI Prospecting Studio to stage drafts, dispatch outreach, and

&#x20;\* report sent message IDs, manual draft promotions, bounces (SMTP DSNs), and

&#x20;\* campaign replies directly inside the client's Google Workspace mailbox with:

&#x20;\* - Sender Display Name (e.g., "First Name Last Name")

&#x20;\* - Sender Email / Alias (e.g., "FirstName.LastName@example.com")

&#x20;\* - Reply-To Email Address (e.g., "FirstName.LastName@example.com")

&#x20;\* - Deliverability Event Polling (authenticated strictly via RELAY\_SECRET)

&#x20;\*

&#x20;\* SETUP STEPS (2 MINUTES):

&#x20;\* 1. Log in to the client's Google Workspace account (ideally FirstName.LastName@example.com).

&#x20;\* 2. Create (or open) the standalone project at https://script.google.com

&#x20;\* 3. Paste this entire code into Code.gs and click Save.

&#x20;\* 4. Go to Project Settings (gear icon) > Script Properties > Add script property:

&#x20;\*      Property: RELAY\_SECRET

&#x20;\*      Value:    <any strong random passphrase, e.g. 32+ characters>

&#x20;\* 5. Select "authorizeRelayPermissions" in the top toolbar and click Run once

&#x20;\*    to grant Gmail permissions in the client's Workspace.

&#x20;\* 6. Click Deploy > New deployment (or Manage deployments > New version) > Web app:

&#x20;\*      - Execute as: Me (the client's Google Workspace email)

&#x20;\*      - Who has access: Anyone

&#x20;\* 7. Copy the Web App URL and paste both values into YOUR main

&#x20;\*    AI Prospecting Studio Project Settings > Script Properties:

&#x20;\*      - CLIENT\_{CLIENT\_NAME}\_GMAIL\_RELAY\_URL    = <Web App URL>

&#x20;\*      - CLIENT\_{CLIENT\_NAME}\_GMAIL\_RELAY\_SECRET = <Same RELAY\_SECRET>

&#x20;\* ============================================================================

&#x20;\*/



function doGet() {

&#x20; return ContentService

&#x20;   .createTextOutput(JSON.stringify({

&#x20;     ok: true,

&#x20;     service: 'Client Gmail Relay (v2.2)',

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



&#x20;   // Deliverability Event Polling (v2.2)

&#x20;   if (action === 'pollEvents') {

&#x20;     return jsonResponse\_(handleRelayPollEvents\_(payload, primaryMailbox));

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



&#x20;   const mailOptions = {};



&#x20;   if (senderName) {

&#x20;     mailOptions.name = senderName;

&#x20;   }

&#x20;   if (replyTo \&\& replyTo.indexOf('@') > 0) {

&#x20;     mailOptions.replyTo = replyTo;

&#x20;   }

&#x20;   if (senderEmail \&\& senderEmail.indexOf('@') > 0 \&\& senderEmail.toLowerCase() !== primaryMailbox.toLowerCase()) {

&#x20;     mailOptions.from = senderEmail;

&#x20;   }



&#x20;   if (action === 'createDraft') {

&#x20;     let draft;

&#x20;     try {

&#x20;       draft = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;     } catch (aliasErr) {

&#x20;       // If senderEmail is an unverified alias on this Workspace account, fallback without 'from' so draft creation still succeeds

&#x20;       if (mailOptions.from) {

&#x20;         delete mailOptions.from;

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

&#x20;       senderEmail: senderEmail || primaryMailbox,

&#x20;       replyTo: mailOptions.replyTo || primaryMailbox,

&#x20;       senderName: senderName || 'Default'

&#x20;     });

&#x20;   }



&#x20;   if (action === 'sendEmail') {

&#x20;     let draftToSend;

&#x20;     try {

&#x20;       draftToSend = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;     } catch (aliasErr) {

&#x20;       if (mailOptions.from) {

&#x20;         delete mailOptions.from;

&#x20;         draftToSend = GmailApp.createDraft(to, subject, body, mailOptions);

&#x20;       } else {

&#x20;         throw aliasErr;

&#x20;       }

&#x20;     }



&#x20;     const draftId = draftToSend.getId();

&#x20;     const sentMsg = draftToSend.send();

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



&#x20;     return jsonResponse\_({

&#x20;       success: true,

&#x20;       action: 'sendEmail',

&#x20;       draftId: draftId,

&#x20;       messageId: messageId,

&#x20;       threadId: threadId,

&#x20;       rfcMessageId: rfcMessageId,

&#x20;       mailbox: primaryMailbox,

&#x20;       senderEmail: senderEmail || primaryMailbox,

&#x20;       replyTo: mailOptions.replyTo || primaryMailbox,

&#x20;       senderName: senderName || 'Default'

&#x20;     });

&#x20;   }



&#x20;   return jsonResponse\_({ success: false, error: 'Unsupported action: ' + action });

&#x20; } catch (err) {

&#x20;   return jsonResponse\_({ success: false, error: err \&\& err.message ? err.message : String(err) });

&#x20; }

}



/\*\*

&#x20;\* Polls the client's Gmail mailbox strictly for:

&#x20;\* 1. Tracked draftIds that have been manually sent by the user.

&#x20;\* 2. DSN bounce notices (SMTP codes \& Hard/Soft reasons).

&#x20;\* 3. Replies strictly from tracked campaign recipient addresses.

&#x20;\*/

function handleRelayPollEvents\_(payload, primaryMailbox) {

&#x20; const lookbackDays = Math.min(Math.max(parseInt(payload.lookbackDays, 10) || 7, 1), 30);

&#x20; const trackedDrafts = Array.isArray(payload.trackedDrafts) ? payload.trackedDrafts.slice(0, 50) : \[];

&#x20; const trackedRecipients = Array.isArray(payload.trackedRecipients) ? payload.trackedRecipients : \[];



&#x20; const recipientSet = {};

&#x20; trackedRecipients.forEach(function(email) {

&#x20;   const clean = String(email || '').trim().toLowerCase();

&#x20;   if (clean \&\& clean.indexOf('@') > 0) {

&#x20;     recipientSet\[clean] = true;

&#x20;   }

&#x20; });



&#x20; const promotedDrafts = \[];

&#x20; const bounces = \[];

&#x20; const replies = \[];

&#x20; const primaryLower = String(primaryMailbox || '').trim().toLowerCase();



&#x20; // 1. Check tracked drafts for manual sends

&#x20; trackedDrafts.forEach(function(draftItem) {

&#x20;   const draftId = String((draftItem \&\& draftItem.draftId) || '').trim();

&#x20;   const recip = String((draftItem \&\& draftItem.email) || '').trim().toLowerCase();

&#x20;   const campId = String((draftItem \&\& draftItem.campaignId) || '').trim();

&#x20;   if (!draftId || !recip) return;



&#x20;   try {

&#x20;     const existingDraft = GmailApp.getDraft(draftId);

&#x20;     if (!existingDraft) {

&#x20;       const sentThreads = GmailApp.search('in:sent to:' + recip + ' newer\_than:' + lookbackDays + 'd', 0, 5);

&#x20;       if (sentThreads \&\& sentThreads.length > 0) {

&#x20;         const msgs = sentThreads\[0].getMessages();

&#x20;         const lastMsg = msgs\[msgs.length - 1];

&#x20;         let rfcId = '';

&#x20;         let msgId = '';

&#x20;         let sentDate = new Date();



&#x20;         if (lastMsg) {

&#x20;           msgId = lastMsg.getId() || '';

&#x20;           sentDate = lastMsg.getDate() || new Date();

&#x20;           try { rfcId = lastMsg.getHeader('Message-ID') || ''; } catch (e) {}

&#x20;         }



&#x20;         promotedDrafts.push({

&#x20;           draftId: draftId,

&#x20;           campaignId: campId,

&#x20;           recipient: recip,

&#x20;           messageId: rfcId || msgId || draftId,

&#x20;           eventDate: sentDate.toISOString()

&#x20;         });

&#x20;         recipientSet\[recip] = true;

&#x20;       }

&#x20;     }

&#x20;   } catch (draftErr) {}

&#x20; });



&#x20; // 2. Scan DSN bounce emails (mailer-daemon / postmaster)

&#x20; const bounceThreads = GmailApp.search('from:(mailer-daemon OR postmaster) newer\_than:' + lookbackDays + 'd', 0, 50);

&#x20; bounceThreads.forEach(function(thread) {

&#x20;   const msgs = thread.getMessages();

&#x20;   msgs.forEach(function(msg) {

&#x20;     let raw = '';

&#x20;     try { raw = msg.getRawContent(); } catch (e) { raw = msg.getPlainBody() || ''; }

&#x20;     const parsed = parseRelaySmtpBounceDetails\_(msg.getSubject(), msg.getPlainBody() || '', raw);

&#x20;     if (!parsed.recipient) return;



&#x20;     bounces.push({

&#x20;       messageId: msg.getId() || '',

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

&#x20;       const sender = extractRelayEmailAddress\_(msg.getFrom());

&#x20;       if (!sender || sender === primaryLower || !recipientSet\[sender]) return;



&#x20;       let rfcId = '';

&#x20;       try { rfcId = msg.getHeader('Message-ID') || ''; } catch (e) {}

&#x20;       const subj = String(msg.getSubject() || '').slice(0, 120);



&#x20;       replies.push({

&#x20;         messageId: rfcId || msg.getId() || '',

&#x20;         recipient: sender,

&#x20;         subject: subj,

&#x20;         reason: 'Reply received: ' + subj,

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



function parseRelaySmtpBounceDetails\_(subject, plainBody, rawContent) {

&#x20; const combined = String(plainBody || '') + '\\n' + String(rawContent || '');



&#x20; let recipient = '';

&#x20; const recipPatterns = \[

&#x20;   /Final-Recipient:\\s\*(?:rfc822;)?\\s\*(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i,

&#x20;   /Original-Recipient:\\s\*(?:rfc822;)?\\s\*(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i,

&#x20;   /X-Failed-Recipients:\\s\*(\[^\\s<>;"',\\r\\n]+@\[^\\s<>;"',\\r\\n]+)/i,

&#x20;   /(?:delivery to the following recipient failed permanently|wasn't delivered to|undeliverable to|could not be delivered to)\[:\\s]+(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i,

&#x20;   /To:\\s\*(\[^\\s<>;"'\\r\\n]+@\[^\\s<>;"'\\r\\n]+)/i

&#x20; ];

&#x20; for (let i = 0; i < recipPatterns.length; i++) {

&#x20;   const m = combined.match(recipPatterns\[i]);

&#x20;   if (m \&\& m\[1]) {

&#x20;     const candidate = m\[1].trim().toLowerCase().replace(/\[.,;>]+$/, '');

&#x20;     if (candidate.indexOf('mailer-daemon') < 0 \&\& candidate.indexOf('postmaster') < 0) {

&#x20;       recipient = candidate;

&#x20;       break;

&#x20;     }

&#x20;   }

&#x20; }



&#x20; let originalMessageId = '';

&#x20; const msgIdMatch = combined.match(/(?:In-Reply-To|Original-Message-ID|Message-ID):\\s\*(<\[^>]+>)/i);

&#x20; if (msgIdMatch \&\& msgIdMatch\[1]) {

&#x20;   originalMessageId = msgIdMatch\[1].trim();

&#x20; }



&#x20; let smtpCode = 'UNKNOWN';

&#x20; const smtpFullMatch = combined.match(/\\b(\[45]\\d{2})\[\\s-]+(\[45]\\.\\d{1,3}\\.\\d{1,3})\\b/);

&#x20; if (smtpFullMatch) {

&#x20;   smtpCode = smtpFullMatch\[1] + ' ' + smtpFullMatch\[2];

&#x20; } else {

&#x20;   const statusMatch = combined.match(/Status:\\s\*(\[45]\\.\\d{1,3}\\.\\d{1,3})/i);

&#x20;   const code3Match = combined.match(/(?:Diagnostic-Code:\[^\\r\\n]\*|Remote server returned\[^\\r\\n]\*|SMTP error\[^\\r\\n]\*)\\b(\[45]\\d{2})\\b/i) ||

&#x20;     combined.match(/\\b(\[45]\[0-5]\\d)\\b/);

&#x20;   if (code3Match \&\& statusMatch) {

&#x20;     smtpCode = code3Match\[1] + ' ' + statusMatch\[1];

&#x20;   } else if (code3Match) {

&#x20;     smtpCode = code3Match\[1];

&#x20;   } else if (statusMatch) {

&#x20;     smtpCode = statusMatch\[1];

&#x20;   }

&#x20; }



&#x20; let reason = '';

&#x20; const diagMatch = combined.match(/Diagnostic-Code:\\s\*(?:smtp;)?\\s\*(\[^\\r\\n]+(?:\\r?\\n\[ \\t]+\[^\\r\\n]+)?)/i) ||

&#x20;   combined.match(/Remote server returned\\s\*'(\[^']+)'/i);

&#x20; if (diagMatch \&\& diagMatch\[1]) {

&#x20;   reason = diagMatch\[1].replace(/\\s+/g, ' ').trim();

&#x20; } else {

&#x20;   reason = String(subject || '').trim() || String(plainBody || '').replace(/\\s+/g, ' ').trim().slice(0, 220);

&#x20; }



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

&#x20;     reasonLow.indexOf('throttl') >= 0) {

&#x20;   bounceType = 'Soft';

&#x20; } else if (codeLead === '5' ||

&#x20;     reasonLow.indexOf('user unknown') >= 0 ||

&#x20;     reasonLow.indexOf('does not exist') >= 0 ||

&#x20;     reasonLow.indexOf('recipientnotfound') >= 0 ||

&#x20;     reasonLow.indexOf('bad-mailbox') >= 0 ||

&#x20;     reasonLow.indexOf('no such user') >= 0 ||

&#x20;     reasonLow.indexOf('domain not found') >= 0 ||

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



function extractRelayEmailAddress\_(headerStr) {

&#x20; const s = String(headerStr || '').trim();

&#x20; const angle = s.match(/<(\[^>]+@\[^>]+)>/);

&#x20; if (angle \&\& angle\[1]) return angle\[1].trim().toLowerCase();

&#x20; const plain = s.match(/(\[^\\s<>;"']+@\[^\\s<>;"']+)/);

&#x20; return plain \&\& plain\[1] ? plain\[1].trim().toLowerCase() : '';

}



function authorizeRelayPermissions() {

&#x20; GmailApp.getDrafts();

&#x20; GmailApp.search('in:sent', 0, 1);

&#x20; Logger.log('Gmail Draft, Send \& Deliverability Read permissions authorized for: ' + getRelayMailboxEmail\_());

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

