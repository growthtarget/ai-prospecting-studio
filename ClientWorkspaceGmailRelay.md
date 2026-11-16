/\*\*

&#x20;\* ============================================================================

&#x20;\* CLIENT GOOGLE WORKSPACE GMAIL RELAY (STANDALONE WEB APP v2.1)

&#x20;\* ============================================================================

&#x20;\* Deploy this script INSIDE the client's own Google Workspace account at:

&#x20;\* https://script.google.com

&#x20;\*

&#x20;\* CLIENT: ClientName (example.com)

&#x20;\* DEFAULT OUTREACH: FirstName.LastName@example.com

&#x20;\* DEFAULT REPLY-TO: FirstName.LastName@example.com

&#x20;\*

&#x20;\* WHAT IT DOES:

&#x20;\* Allows your AI Prospecting Studio to stage drafts or dispatch outreach

&#x20;\* directly inside the client's Google Workspace mailbox with support for:

&#x20;\* - Sender Display Name (e.g., "First Name Last Name")

&#x20;\* - Sender Email / Alias (e.g., "FirstName.LastName@example.com")

&#x20;\* - Reply-To Email Address (e.g., "FirstName.LastName@example.com")

&#x20;\*

&#x20;\* SETUP STEPS (2 MINUTES):

&#x20;\* 1. Log in to the client's Google Workspace account (ideally FirstName.LastName@example.com).

&#x20;\* 2. Create a new standalone project at https://script.google.com

&#x20;\* 3. Paste this entire code into Code.gs and click Save.

&#x20;\* 4. Go to Project Settings (gear icon) > Script Properties > Add script property:

&#x20;\*      Property: RELAY\_SECRET

&#x20;\*      Value:    <any strong random passphrase, e.g. 32+ characters>

&#x20;\* 5. Select "authorizeRelayPermissions" in the top toolbar and click Run once

&#x20;\*    to grant Gmail permissions in the client's Workspace.

&#x20;\* 6. Click Deploy > New deployment > Type: Web app:

&#x20;\*      - Execute as: Me (the client's Google Workspace email)

&#x20;\*      - Who has access: Anyone

&#x20;\* 7. Copy the generated Web App URL and paste both values into YOUR main

&#x20;\*    AI Prospecting Studio Project Settings > Script Properties:

&#x20;\*      - CLIENT\_EVALUATER\_GMAIL\_RELAY\_URL    = <Web App URL>

&#x20;\*      - CLIENT\_EVALUATER\_GMAIL\_RELAY\_SECRET = <Same RELAY\_SECRET>

&#x20;\* ============================================================================

&#x20;\*/



function doGet() {

&#x20; return ContentService

&#x20;   .createTextOutput(JSON.stringify({

&#x20;     ok: true,

&#x20;     service: 'Client Gmail Relay (v2.1)',

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



&#x20;   const primaryMailbox = getRelayMailboxEmail\_();

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

&#x20;     try {

&#x20;       GmailApp.sendEmail(to, subject, body, mailOptions);

&#x20;     } catch (aliasErr) {

&#x20;       if (mailOptions.from) {

&#x20;         delete mailOptions.from;

&#x20;         GmailApp.sendEmail(to, subject, body, mailOptions);

&#x20;       } else {

&#x20;         throw aliasErr;

&#x20;       }

&#x20;     }



&#x20;     return jsonResponse\_({

&#x20;       success: true,

&#x20;       action: 'sendEmail',

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



function authorizeRelayPermissions() {

&#x20; GmailApp.getDrafts();

&#x20; Logger.log('Gmail Draft \& Send permissions authorized for: ' + getRelayMailboxEmail\_());

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

