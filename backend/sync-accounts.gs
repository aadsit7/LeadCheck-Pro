// ============================================================
// LeadCheck Pro — Backend: the account-aware Contact Sync actions
// (annotated reference copy — Code.gs in this folder already contains all of
// this; paste Code.gs whole for the normal update path)
// ============================================================
//
// These power the "Sync contacts" button, the auto-sync on open, and the
// near-live change feed for both Partner accounts and Opportunity accounts.
// When the front-end syncs a company it POSTs { action, companyName } to the
// Web App; this code:
//
//   1. SEARCHES the "web API worksheet" (the Partner_Portal_Database spreadsheet)
//      for the partner / opportunity account that is RELEVANT to companyName,
//      using tolerant, legal-suffix-aware matching so the right account is found
//      even when the name is written slightly differently.
//   2. GATHERS *all* of that account's content — every meeting recap, deal
//      note, Opportunity_Descriptions history row, transcript, Meeting_Index
//      summary (attendees / key decisions), Partner_Documents body, and
//      Opportunity_Documents attached-file record — with the HTML stripped to
//      clean text and nothing truncated away.
//   3. Runs an ADVANCED REASONING pass (Claude, extended thinking) over the full
//      notes to extract every real, named person who actually works at the target
//      account — deliberately excluding the vendor/seller (Recast) side and any
//      unrelated third parties — and returns them in the exact JSON shape the
//      front-end's syncOpportunityContacts() already consumes.
//
// It also serves the near-live change feed (action `listSourceAccounts`):
// every partner & opportunity account with an MD5 content signature computed
// over EXACTLY the note text the extraction would receive. The front-end
// stores each signature on the Companies row after a successful analysis
// (columns `sync_signature` / `sync_checked`, auto-created on first write) and
// re-analyzes an account only when its signature changes — that is what makes
// opening the app pick up new accounts and new notes automatically without
// re-running the AI for unchanged accounts.
//
// ------------------------------------------------------------
// HOW TO INSTALL (one-time)
// ------------------------------------------------------------
// 1. Open your existing Apps Script project (the one behind CONFIG.WEB_APP_URL)
//    and paste EVERYTHING in this file into Code.gs.
//    - It reuses the ANTHROPIC_API_KEY constant and the jsonResponse() helper
//      already defined there (see reason-org.gs / the README), so no secret is
//      duplicated here.
//    - If your project ALREADY defines syncPartners / syncOpportunities /
//      listPartnerCompanies / listOpportunityCompanies / listSourceAccounts
//      handlers, REPLACE those older implementations with the ones below.
//      (The private helpers are all prefixed `lcSync_` to avoid clashing with
//      anything else.)
// 2. In doPost(e), route the five actions to these handlers (add or replace the
//    matching branches). `payload` is your parsed JSON body:
//
//        if (payload.action === 'listPartnerCompanies')     return doListPartnerCompanies();
//        if (payload.action === 'listOpportunityCompanies') return doListOpportunityCompanies();
//        if (payload.action === 'syncPartners')             return doSyncPartners(payload.companyName);
//        if (payload.action === 'syncOpportunities')        return doSyncOpportunities(payload.companyName);
//        if (payload.action === 'listSourceAccounts')       return doListSourceAccounts();
//
//    And so the signature columns can be auto-created on the Companies tab,
//    add this right after doPost reads the header row (see Code.gs):
//
//        if (tab === 'Companies') headers = lcEnsureSyncColumns_(sheet, headers, data);
//
// 3. Deploy → Manage deployments → Edit → New version → Deploy.
//
// SECURITY — keep the API key out of the source. Store it in Project Settings →
// Script properties (property name ANTHROPIC_API_KEY) and read it via:
//        const ANTHROPIC_API_KEY =
//          PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
// This file references that same ANTHROPIC_API_KEY constant.
// ============================================================


// ------------------------------------------------------------
// Configuration
// ------------------------------------------------------------

// The "web API worksheet" that holds the partner & opportunity accounts and
// their description notes (the Partner_Portal_Database spreadsheet). The Web App
// itself is bound to the LeadCheck *database* sheet (Contacts/Companies/…), so
// the source data lives in a different spreadsheet and we open it by id.
// Override at runtime with a Script property named PARTNER_PORTAL_SHEET_ID if it
// ever moves.
function lcSync_sourceSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty('PARTNER_PORTAL_SHEET_ID')
    || '18Yhe3Yiq9_eI7kBxtFOzdu6Pb0_VUx730TYjq1xPjzI';
}

// Strongest widely-available reasoning model. Change to a model your key can
// access if needed (e.g. 'claude-sonnet-4-20250514', already used elsewhere in
// this backend).
var SYNC_EXTRACTION_MODEL = 'claude-opus-4-8';

// Generous cap so we effectively send *all* the notes while staying inside the
// model's context window.
var SYNC_MAX_NOTE_CHARS = 120000;

// Version stamp mixed into every account's change signature. Bump this when the
// note-gathering logic changes (new tabs read, different formatting) so every
// account re-analyzes once against the new, richer notes.
var SYNC_SIGNATURE_VERSION = '2';

// Sync-state columns the backend is allowed to auto-create on the Companies
// tab. `sync_signature` stores the source-content signature captured when an
// account's notes were last fully analyzed; `sync_checked` stores when.
var SYNC_STATE_COLUMNS = ['sync_signature', 'sync_checked'];


// ============================================================
// LIST OPPORTUNITY COMPANIES
// ============================================================

function doListOpportunityCompanies() {
  var ss = SpreadsheetApp.openById(lcSync_sourceSpreadsheetId_());
  var oppSheet = ss.getSheetByName('Opportunities');
  if (!oppSheet) throw new Error('Opportunities tab not found in Partner Portal sheet');

  var lastRow = oppSheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ ok: true, companies: [] });

  var headers = oppSheet.getRange(1, 1, 1, oppSheet.getLastColumn()).getValues()[0];
  var custIdx = headers.indexOf('customer_name');
  if (custIdx === -1) throw new Error('customer_name column not found in Opportunities tab');

  var data = oppSheet.getRange(2, 1, lastRow - 1, oppSheet.getLastColumn()).getValues();
  var seen = {};
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var name = (data[i][custIdx] || '').toString().trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    out.push(name);
  }
  out.sort(function(a, b) { return a.localeCompare(b); });

  return jsonResponse({ ok: true, companies: out });
}

// ============================================================
// SYNC OPPORTUNITIES — Extract contacts from the relevant opportunity account.
// Searches the Opportunities tab (customer_name / deal_name), reads ALL of that
// account's notes — the inline description/notes columns AND every
// Opportunity_Descriptions history row for those deals — and runs an advanced
// reasoning pass to pull out only the customer-side people.
// ============================================================

function doSyncOpportunities(companyName) {
  companyName = String(companyName || '').trim();
  if (!companyName) return lcSync_emptyContacts_();

  var ss = SpreadsheetApp.openById(lcSync_sourceSpreadsheetId_());
  var notes = lcSync_oppNotesForAccount_(ss, companyName);
  if (!notes) return lcSync_emptyContacts_();

  return lcSync_extractContacts_(companyName, 'opportunity (prospective customer)', notes);
}

// All opportunity-side notes for one account, exactly as the extraction sees
// them. This is the SINGLE shared path used by both doSyncOpportunities and
// doListSourceAccounts, so the change signature is always computed over the
// same text the analysis would receive — if this function's output changes,
// the signature changes, and the account re-analyzes.
function lcSync_oppNotesForAccount_(ss, companyName) {
  var opps = lcSync_sheetObjects_(ss, 'Opportunities');
  var matched = opps.filter(function (o) {
    return lcSync_accountMatches_(companyName, o.customer_name) ||
           lcSync_accountMatches_(companyName, o.deal_name);
  });
  if (!matched.length) return '';
  return lcSync_gatherOpportunityNotes_(ss, companyName, matched);
}

// ============================================================
// LIST PARTNER COMPANIES
// ============================================================

function doListPartnerCompanies() {
  var ss = SpreadsheetApp.openById(lcSync_sourceSpreadsheetId_());
  var partnerSheet = ss.getSheetByName('Partners');
  if (!partnerSheet) throw new Error('Partners tab not found in Partner Portal sheet');

  var lastRow = partnerSheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ ok: true, companies: [] });

  var headers = partnerSheet.getRange(1, 1, 1, partnerSheet.getLastColumn()).getValues()[0];
  var custIdx = headers.indexOf('display_name');
  if (custIdx === -1) custIdx = headers.indexOf('customer_name');
  if (custIdx === -1) custIdx = headers.indexOf('partner_name');
  if (custIdx === -1) throw new Error('display_name (or customer_name / partner_name) column not found in Partners tab');

  var data = partnerSheet.getRange(2, 1, lastRow - 1, partnerSheet.getLastColumn()).getValues();
  var seen = {};
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var name = (data[i][custIdx] || '').toString().trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    out.push(name);
  }
  out.sort(function(a, b) { return a.localeCompare(b); });

  return jsonResponse({ ok: true, companies: out });
}

// ============================================================
// SYNC PARTNERS — Extract contacts from the relevant partner account.
// Searches the Partners tab (display_name), then reads that partner's notes —
// the Transcripts (transcript_text, keyed by partner_id) plus any deal where the
// partner is itself the customer — and runs an advanced reasoning pass to pull
// out only the partner-side people.
// ============================================================

function doSyncPartners(companyName) {
  companyName = String(companyName || '').trim();
  if (!companyName) return lcSync_emptyContacts_();

  var ss = SpreadsheetApp.openById(lcSync_sourceSpreadsheetId_());
  var notes = lcSync_partnerNotesForAccount_(ss, companyName);
  if (!notes) return lcSync_emptyContacts_();

  return lcSync_extractContacts_(companyName, 'channel partner', notes);
}

// All partner-side notes for one account, exactly as the extraction sees them.
// Shared by doSyncPartners and doListSourceAccounts — see the invariant note on
// lcSync_oppNotesForAccount_.
function lcSync_partnerNotesForAccount_(ss, companyName) {
  var partners = lcSync_sheetObjects_(ss, 'Partners');
  var matchedPartners = partners.filter(function (p) {
    return lcSync_accountMatches_(companyName, p.display_name) ||
           lcSync_accountMatches_(companyName, p.username);
  });
  if (!matchedPartners.length) return '';

  var partnerIds = {};
  matchedPartners.forEach(function (p) {
    var id = String(p.partner_id == null ? '' : p.partner_id).trim();
    if (id) partnerIds[id] = true;
  });

  return lcSync_gatherPartnerNotes_(ss, companyName, partnerIds);
}

// ============================================================
// LIST SOURCE ACCOUNTS — the near-live change feed.
// One cheap call returns every partner & opportunity account in the source
// spreadsheet together with a content signature over ALL of that account's
// notes, documents, transcripts and attached-file records. The front-end
// compares each signature with the one saved on the Companies row after the
// last successful analysis:
//   • name not in Companies            → new account  → add + analyze
//   • signature differs from saved     → new content  → re-analyze
//   • signature matches                → nothing new  → skip (no AI cost)
// Signatures are computed over the EXACT text the extraction would receive
// (same gather functions), so "signature unchanged" always means "an analysis
// run would see identical input".
// ============================================================

function doListSourceAccounts() {
  var ss = SpreadsheetApp.openById(lcSync_sourceSpreadsheetId_());

  // Union of account names, mirroring doListOpportunityCompanies (customer_name)
  // and doListPartnerCompanies (display_name → customer_name → partner_name).
  var byKey = {};
  var order = [];
  function addName(raw, source) {
    var name = String(raw == null ? '' : raw).trim();
    if (!name) return;
    var key = name.toLowerCase();
    if (!byKey[key]) {
      byKey[key] = { name: name, opportunity: false, partner: false };
      order.push(key);
    }
    byKey[key][source] = true;
  }
  lcSync_sheetObjects_(ss, 'Opportunities').forEach(function (o) {
    addName(o.customer_name, 'opportunity');
  });
  lcSync_sheetObjects_(ss, 'Partners').forEach(function (p) {
    var name = String(p.display_name == null ? '' : p.display_name).trim() ||
               String(p.customer_name == null ? '' : p.customer_name).trim() ||
               String(p.partner_name == null ? '' : p.partner_name).trim();
    addName(name, 'partner');
  });

  order.sort(function (a, b) { return byKey[a].name.localeCompare(byKey[b].name); });

  var accounts = order.map(function (key) {
    var a = byKey[key];
    // Both sides are gathered unconditionally because the front-end's contact
    // sync always runs BOTH syncOpportunities and syncPartners for a company —
    // e.g. a partner can also appear as a customer on a deal.
    var oppNotes = lcSync_oppNotesForAccount_(ss, a.name);
    var partnerNotes = lcSync_partnerNotesForAccount_(ss, a.name);
    return {
      name: a.name,
      opportunity: a.opportunity,
      partner: a.partner,
      signature: lcSync_signature_(oppNotes, partnerNotes),
      note_chars: oppNotes.length + partnerNotes.length
    };
  });

  return jsonResponse({
    ok: true,
    signature_version: SYNC_SIGNATURE_VERSION,
    accounts: accounts
  });
}

// Stable MD5 hex signature over an account's full note text (both sides), with
// the gathering-logic version mixed in so improved gathering re-triggers one
// analysis per account.
function lcSync_signature_(oppNotes, partnerNotes) {
  var text = 'v' + SYNC_SIGNATURE_VERSION + '\u0001' +
             String(oppNotes == null ? '' : oppNotes) + '\u0001' +
             String(partnerNotes == null ? '' : partnerNotes);
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

// Auto-create the whitelisted sync-state columns on the Companies tab the
// first time a write carries them, then return the refreshed header row. Only
// SYNC_STATE_COLUMNS can ever be created — arbitrary payload keys still map
// onto existing headers only.
function lcEnsureSyncColumns_(sheet, headers, data) {
  if (!data) return headers;
  var missing = SYNC_STATE_COLUMNS.filter(function (c) {
    return data[c] !== undefined && headers.indexOf(c) === -1;
  });
  if (!missing.length) return headers;
  missing.forEach(function (c) {
    var col = sheet.getLastColumn() + 1;
    if (sheet.getMaxColumns() < col) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), col - sheet.getMaxColumns());
    }
    sheet.getRange(1, col).setValue(c);
  });
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ------------------------------------------------------------
// Contact-sync note gathering
// ------------------------------------------------------------

// All of an opportunity account's description notes. Reads EVERY place the
// portal stores them, so nothing is missed:
//   (a) the inline `description` + `notes` columns on the matched Opportunities
//       rows,
//   (b) every `Opportunity_Descriptions.description_text` row keyed to those
//       opportunity_ids — the full history of meeting recaps / documents (there
//       are typically several per deal), and
//   (c) every `Opportunity_Documents` row for those deals — the files attached
//       to the opportunity (name/type/date metadata; the file bodies live in
//       Drive), so a newly attached file both informs the analysis and changes
//       the account's change signature.
function lcSync_gatherOpportunityNotes_(ss, companyName, rows) {
  var oppIds = {};
  rows.forEach(function (o) {
    var id = String(o.opportunity_id == null ? '' : o.opportunity_id).trim();
    if (id) oppIds[id] = true;
  });

  var parts = [
    lcSync_inlineOppNotes_(rows),
    lcSync_oppDescriptionNotes_(ss, oppIds),
    lcSync_oppDocumentNotes_(ss, companyName, oppIds)
  ].filter(Boolean);

  return lcSync_capNotes_(parts.join('\n\n'));
}

// Attached-file records from the Opportunity_Documents tab, matched by
// opportunity_id (with a customer_name fallback so a document row whose deal
// link is broken is still counted for the right account).
function lcSync_oppDocumentNotes_(ss, companyName, oppIdSet) {
  var parts = [];
  lcSync_sheetObjects_(ss, 'Opportunity_Documents').forEach(function (d) {
    var id = String(d.opportunity_id == null ? '' : d.opportunity_id).trim();
    var byId = id && oppIdSet && oppIdSet[id];
    var byName = lcSync_accountMatches_(companyName, d.customer_name);
    if (!byId && !byName) return;

    var name = String(d.file_name == null ? '' : d.file_name).trim();
    if (!name) return;
    var bits = ['File: ' + name];
    if (d.mime_type)     bits.push('Type: ' + d.mime_type);
    if (d.date_added)    bits.push('Added: ' + d.date_added);
    if (d.customer_name) bits.push('Customer: ' + d.customer_name);
    parts.push('=== OPPORTUNITY ATTACHED FILE ===\n' + bits.join(' | '));
  });
  return parts.join('\n\n');
}

// The `description` + `notes` columns stored directly on the Opportunities rows.
function lcSync_inlineOppNotes_(rows) {
  var parts = [];
  rows.forEach(function (o) {
    var header = [];
    if (o.deal_name)     header.push('Deal: ' + o.deal_name);
    if (o.customer_name) header.push('Customer: ' + o.customer_name);
    if (o.stage)         header.push('Stage: ' + o.stage);
    if (o.status)        header.push('Status: ' + o.status);

    var body = [lcSync_htmlToText_(o.description), lcSync_htmlToText_(o.notes)]
      .filter(Boolean).join('\n\n');
    if (body) {
      parts.push('=== OPPORTUNITY NOTE ===\n' +
        (header.length ? header.join(' | ') + '\n' : '') + body);
    }
  });
  return parts.join('\n\n');
}

// Every Opportunity_Descriptions.description_text row whose opportunity_id is in
// the given set — the deal's full recap/document history.
function lcSync_oppDescriptionNotes_(ss, oppIdSet) {
  if (!oppIdSet) return '';
  var has = false;
  for (var k in oppIdSet) { if (oppIdSet[k]) { has = true; break; } }
  if (!has) return '';

  var parts = [];
  lcSync_sheetObjects_(ss, 'Opportunity_Descriptions').forEach(function (d) {
    var id = String(d.opportunity_id == null ? '' : d.opportunity_id).trim();
    if (!id || !oppIdSet[id]) return;

    var body = lcSync_htmlToText_(d.description_text);
    if (!body) return;
    var head = [];
    if (d.deal_name)        head.push('Deal: ' + d.deal_name);
    if (d.description_date)  head.push('Date: ' + d.description_date);
    parts.push('=== OPPORTUNITY DESCRIPTION ===\n' +
      (head.length ? head.join(' | ') + '\n' : '') + body);
  });
  return parts.join('\n\n');
}

function lcSync_gatherPartnerNotes_(ss, companyName, partnerIds) {
  var parts = [];

  // 1. Transcripts keyed to this partner — the primary partner-notes source.
  lcSync_sheetObjects_(ss, 'Transcripts').forEach(function (t) {
    var id = String(t.partner_id == null ? '' : t.partner_id).trim();
    var byId = id && partnerIds[id];
    var byName = lcSync_accountMatches_(companyName, t.partner_name);
    if (!byId && !byName) return;

    var body = lcSync_htmlToText_(t.transcript_text);
    if (!body) return;
    var head = [];
    if (t.partner_name)      head.push('Partner: ' + t.partner_name);
    if (t.conversation_date) head.push('Date: ' + t.conversation_date);
    parts.push('=== PARTNER TRANSCRIPT ===\n' +
      (head.length ? head.join(' | ') + '\n' : '') + body);
  });

  // 1b. Meeting_Index rows for this partner — structured recaps of those
  //     conversations with the strongest contact signal of all: an explicit
  //     attendees list, plus summary / key decisions / topics.
  lcSync_sheetObjects_(ss, 'Meeting_Index').forEach(function (m) {
    var id = String(m.partner_id == null ? '' : m.partner_id).trim();
    var byId = id && partnerIds[id];
    var byName = lcSync_accountMatches_(companyName, m.partner_name);
    if (!byId && !byName) return;

    var head = [];
    if (m.meeting_title) head.push('Meeting: ' + m.meeting_title);
    if (m.partner_name)  head.push('Partner: ' + m.partner_name);
    if (m.meeting_date)  head.push('Date: ' + m.meeting_date);
    var lines = [];
    var attendees = lcSync_htmlToText_(m.attendees);
    var summary = lcSync_htmlToText_(m.summary);
    var decisions = lcSync_htmlToText_(m.key_decisions);
    var topics = lcSync_htmlToText_(m.topics_discussed);
    if (attendees) lines.push('Attendees: ' + attendees);
    if (summary)   lines.push('Summary: ' + summary);
    if (decisions) lines.push('Key decisions: ' + decisions);
    if (topics)    lines.push('Topics: ' + topics);
    if (!lines.length) return;
    parts.push('=== PARTNER MEETING SUMMARY ===\n' +
      (head.length ? head.join(' | ') + '\n' : '') + lines.join('\n'));
  });

  // 1c. Partner_Documents for this partner — enablement plans, joint business
  //     plans etc. whose full HTML body is stored right in the sheet, so the
  //     document text itself is analyzed (and signature-tracked), not just its
  //     title.
  lcSync_sheetObjects_(ss, 'Partner_Documents').forEach(function (d) {
    var id = String(d.partner_id == null ? '' : d.partner_id).trim();
    var byId = id && partnerIds[id];
    var byName = lcSync_accountMatches_(companyName, d.partner_name);
    if (!byId && !byName) return;

    var body = lcSync_htmlToText_(d.html_content);
    var head = [];
    if (d.title)      head.push('Document: ' + d.title);
    if (d.doc_type)   head.push('Type: ' + d.doc_type);
    if (d.partner_name) head.push('Partner: ' + d.partner_name);
    if (d.updated_at) head.push('Updated: ' + d.updated_at);
    else if (d.created_at) head.push('Created: ' + d.created_at);
    if (!body && !d.title) return;
    parts.push('=== PARTNER DOCUMENT ===\n' +
      (head.length ? head.join(' | ') + '\n' : '') + (body || '[Document has no text content.]'));
  });

  // 2. Deals where this partner is ALSO the customer (i.e. the partner buying
  //    for itself) — those notes describe the partner's own people, and we pull
  //    both the inline columns and the full Opportunity_Descriptions history.
  //    Deals the partner registered for OTHER end customers are intentionally
  //    left out here: those people belong to the opportunity account and are
  //    picked up by doSyncOpportunities(thatCustomer) instead, so the partner
  //    roster never gets contaminated with an unrelated customer's staff.
  var selfDeals = lcSync_sheetObjects_(ss, 'Opportunities').filter(function (o) {
    var id = String(o.partner_id == null ? '' : o.partner_id).trim();
    var underPartner = id && partnerIds[id];
    var aboutPartner = lcSync_accountMatches_(companyName, o.customer_name) ||
                       lcSync_accountMatches_(companyName, o.deal_name);
    return underPartner && aboutPartner;
  });
  if (selfDeals.length) {
    var inline = lcSync_inlineOppNotes_(selfDeals);
    if (inline) parts.push(inline);
    var ids = {};
    selfDeals.forEach(function (o) {
      var id = String(o.opportunity_id == null ? '' : o.opportunity_id).trim();
      if (id) ids[id] = true;
    });
    var hist = lcSync_oppDescriptionNotes_(ss, ids);
    if (hist) parts.push(hist);
    var docs = lcSync_oppDocumentNotes_(ss, companyName, ids);
    if (docs) parts.push(docs);
  }

  return lcSync_capNotes_(parts.join('\n\n'));
}

function lcSync_capNotes_(text) {
  text = String(text || '').trim();
  if (text.length > SYNC_MAX_NOTE_CHARS) {
    text = text.slice(0, SYNC_MAX_NOTE_CHARS) + '\n\n[Notes truncated for length.]';
  }
  return text;
}

// ------------------------------------------------------------
// Contact-sync advanced-reasoning extraction (Claude)
// ------------------------------------------------------------

function lcSync_extractContacts_(companyName, sideLabel, notesText) {
  var prompt = lcSync_buildPrompt_(companyName, sideLabel, notesText);

  // First pass: extended thinking for the most accurate read of the notes.
  var first = lcSync_callAnthropic_(prompt, true);
  if (lcSync_hasTextBlock_(first)) {
    return ContentService.createTextOutput(first).setMimeType(ContentService.MimeType.JSON);
  }

  // Fallback: if thinking consumed the whole budget and left no text block,
  // retry once WITHOUT thinking so the front-end always receives the JSON array.
  var second = lcSync_callAnthropic_(prompt, false);
  return ContentService.createTextOutput(second).setMimeType(ContentService.MimeType.JSON);
}

function lcSync_callAnthropic_(prompt, useThinking) {
  var body = {
    model: SYNC_EXTRACTION_MODEL,
    max_tokens: 12000,
    messages: [{ role: 'user', content: prompt }]
  };
  // Adaptive thinking improves reasoning quality and keeps the visible response
  // as clean JSON. Remove this line if your key rejects it.
  if (useThinking) body.thinking = { type: 'adaptive' };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return response.getContentText();
}

function lcSync_hasTextBlock_(raw) {
  try {
    var data = JSON.parse(raw);
    if (!data || !Array.isArray(data.content)) return false;
    return data.content.some(function (b) {
      return b && b.type === 'text' && String(b.text || '').trim();
    });
  } catch (e) { return false; }
}

function lcSync_buildPrompt_(companyName, sideLabel, notesText) {
  return [
'You are a meticulous B2B sales-operations analyst. Read ALL of the meeting notes, transcripts, deal descriptions, meeting summaries, partner documents and attached-file records below and extract EVERY real, named person who works at the target account, so they can be added to a CRM contact database.',
'',
'TARGET ACCOUNT: "' + companyName + '"  (this is a ' + sideLabel + ' account).',
'',
'=== WHO TO EXTRACT ===',
'- Include ONLY real, named individuals who are employees or stakeholders AT "' + companyName + '" (the ' + sideLabel + ' / buying organization).',
'- A person qualifies only if the notes give an actual personal name (first name, or first + last). Include them even when only a first name is given.',
'',
'=== WHO TO EXCLUDE (this is critical for accuracy) ===',
'- The VENDOR / selling side and its representatives. That is Recast Software — often labelled "Recast", "vendor", "vendor company", "sales representative", or "Recast representative". NEVER extract vendor/seller people.',
'- Anyone who works at a DIFFERENT organization: distributors (e.g. TD SYNNEX), competitors or third-party products named in passing (e.g. CDW, Citrix), other partners, or an unrelated end customer.',
'- Generic or unnamed roles with no personal name ("the technical lead", "Legal team", "VMO", "Marketing team", "IT Engineer (name not discussed)"). Skip these entirely — do not invent a name for them.',
'- Do NOT invent, guess, or add anyone who is not explicitly named in the notes.',
'',
'=== HOW TO REASON ===',
'- Read every note in full. The same person may appear several times — sometimes by first name only, later by full name. MERGE those into a single entry using the most complete name and the union of everything said about them.',
'- "=== PARTNER MEETING SUMMARY ===" blocks include an explicit Attendees list — a strong signal for who exists — but attendees may belong to EITHER side; use the exclusion rules to keep only people at the target account.',
'- "=== OPPORTUNITY ATTACHED FILE ===" blocks describe files by name/metadata only. You cannot see the file contents, so NEVER invent people from a file name; use these records only as context.',
'- Many notes end with a "People" section that labels each person\'s side (e.g. "James – Manager – Customer company", "Aaron – Recast"). Treat those labels as the strongest signal for whether someone belongs to "' + companyName + '" versus the vendor or another org.',
'- Derive every field ONLY from the notes. If the notes do not support a field, leave it as an empty string (or an empty array for priorities). Never fabricate titles, emails, or reporting lines.',
'',
'=== FIELDS (one object per person) ===',
'- "name": the person\'s full name as best known from the notes.',
'- "title": their job title / role at the company, if stated or clearly implied (else "").',
'- "email": their email address ONLY if it literally appears in the notes (else "").',
'- "linkedin_url": their LinkedIn URL ONLY if it literally appears in the notes (else "").',
'- "role_in_deal": their role in this deal/relationship — e.g. "Decision maker", "Economic buyer", "Champion", "Technical evaluator", "Influencer", "Blocker", "End user" — inferred from what they do in the notes (else "").',
'- "reports_to": the NAME of the person they report to, only if the notes state or clearly imply it (e.g. "Parvati – Director (Gerard\'s manager)" means Gerard reports_to Parvati). Manager name only (else "").',
'- "sentiment": their disposition toward the vendor/solution in one short phrase grounded in the notes — e.g. "Skeptical — reframed the deal back to discovery", "Supportive — impressed by the POC" (else "").',
'- "priorities": an array of this person\'s stated priorities, concerns or objections, each a short phrase drawn from the notes (else []).',
'- "context": one to three sentences summarising what the notes say about THIS person. The summary MUST mention the person by name. Keep it factual and specific.',
'',
'=== OUTPUT ===',
'Respond with ONLY a valid JSON array — no markdown, no backticks, no preamble, no trailing commentary:',
'[{"name":"","title":"","email":"","linkedin_url":"","role_in_deal":"","reports_to":"","sentiment":"","priorities":[],"context":""}]',
'If the notes contain no real, named people who work at "' + companyName + '", return exactly: []',
'',
'=== NOTES FOR "' + companyName + '" ===',
notesText
  ].join('\n');
}

// ------------------------------------------------------------
// Contact-sync account matching + sheet/HTML helpers
// ------------------------------------------------------------

// Read a worksheet into an array of row objects keyed by header name. Blank rows
// are skipped. Returns [] if the tab is missing or has only a header.
//
// Results are memoized for the lifetime of the execution (Apps Script globals
// reset on every request, so this can never go stale across requests). The
// sync/list actions only READ these source tabs, and doListSourceAccounts
// gathers notes for every account in one request — without the memo it would
// re-read each tab once per account.
var LC_TAB_CACHE_ = Object.create(null);
function lcSync_sheetObjects_(ss, name) {
  var cacheKey = ss.getId() + '::' + name;
  if (LC_TAB_CACHE_[cacheKey]) return LC_TAB_CACHE_[cacheKey];

  var sh = ss.getSheetByName(name);
  if (!sh) return (LC_TAB_CACHE_[cacheKey] = []);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return (LC_TAB_CACHE_[cacheKey] = []);

  var headers = values[0].map(function (h) { return String(h == null ? '' : h).trim(); });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r], obj = {}, any = false;
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      if (!key) continue;
      var val = row[c];
      obj[key] = (val == null) ? '' : val;
      if (String(val == null ? '' : val).trim()) any = true;
    }
    if (any) out.push(obj);
  }
  LC_TAB_CACHE_[cacheKey] = out;
  return out;
}

// Normalise a company/account name for tolerant comparison: lower-cased, with
// punctuation and common legal-entity suffixes removed so "Insight",
// "Insight Enterprises" and "Insight Enterprises, Inc." all reduce to "insight".
// Memoized per execution — doListSourceAccounts compares every account against
// every source row, so the same strings normalize thousands of times.
var LC_NORM_CACHE_ = Object.create(null);
function lcSync_normalizeAccount_(s) {
  var raw = String(s == null ? '' : s);
  var hit = LC_NORM_CACHE_[raw];
  if (hit !== undefined) return hit;
  var norm = raw
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(inc|incorporated|llc|ltd|limited|corp|corporation|co|company|plc|gmbh|ag|sa|nv|bv|pty|group|holdings|enterprises|the)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  LC_NORM_CACHE_[raw] = norm;
  return norm;
}

// True when two account names refer to the same company. Layered from strict to
// tolerant so the RELEVANT account is found even with spacing/suffix/format
// differences, without matching genuinely different companies.
function lcSync_accountMatches_(query, candidate) {
  var q = lcSync_normalizeAccount_(query);
  var c = lcSync_normalizeAccount_(candidate);
  if (!q || !c) return false;
  if (q === c) return true;

  var qs = q.replace(/\s+/g, '');
  var cs = c.replace(/\s+/g, '');
  if (qs === cs) return true;                              // "green shield" == "greenshield"

  // Whole-phrase, token-aligned containment: "insight" within "insight enterprises".
  if ((' ' + c + ' ').indexOf(' ' + q + ' ') >= 0) return true;
  if ((' ' + q + ' ').indexOf(' ' + c + ' ') >= 0) return true;

  // Spaceless prefix, for concatenations/typos of the shorter within the longer.
  if (qs.length >= 4 && cs.length >= 4 && (cs.indexOf(qs) === 0 || qs.indexOf(cs) === 0)) return true;

  // Token-subset: every token of the shorter name appears in the longer one.
  var qt = q.split(' ').filter(Boolean);
  var ct = c.split(' ').filter(Boolean);
  var shorter = qt.length <= ct.length ? qt : ct;
  var longer = qt.length <= ct.length ? ct : qt;
  var longerSet = {};
  longer.forEach(function (t) { longerSet[t] = true; });
  if (shorter.length && shorter.every(function (t) { return longerSet[t]; })) return true;

  return false;
}

// Convert the rich-text HTML stored in description / transcript cells into clean,
// readable plain text: block closings and list items become line breaks, tags are
// dropped, and the common HTML entities are decoded. Keeps the notes fully
// legible to the model without any markup noise.
function lcSync_htmlToText_(html) {
  if (html == null) return '';
  var s = String(html);
  if (s.indexOf('<') < 0 && s.indexOf('&') < 0) return s.replace(/\s+/g, ' ').trim();

  s = s.replace(/<li[^>]*>/gi, '\n• ');
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<\/(p|div|li|h[1-6]|tr|ul|ol|table)>/gi, '\n');
  s = s.replace(/<[^>]+>/g, ' ');

  s = s.replace(/&nbsp;/gi, ' ')
       .replace(/&amp;/gi, '&')
       .replace(/&lt;/gi, '<')
       .replace(/&gt;/gi, '>')
       .replace(/&quot;/gi, '"')
       .replace(/&#39;|&apos;|&rsquo;/gi, "'")
       .replace(/&ldquo;|&rdquo;/gi, '"')
       .replace(/&mdash;/gi, '—')
       .replace(/&ndash;/gi, '–')
       .replace(/&hellip;/gi, '…');

  s = s.replace(/[ \t\f\v]+/g, ' ')
       .replace(/ *\n */g, '\n')
       .replace(/\n{3,}/g, '\n\n')
       .trim();
  return s;
}

// Empty-result response in the same Anthropic content shape the front-end unwraps
// (an empty JSON array of contacts).
function lcSync_emptyContacts_() {
  return jsonResponse({ ok: true, content: [{ type: 'text', text: '[]' }] });
}
