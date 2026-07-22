// ============================================================
// LeadCheck Pro — Apps Script Backend
// Sheet IDs live here; the Anthropic API key is read from
// Project Settings → Script properties (property name: ANTHROPIC_API_KEY),
// so no secret is stored in this file or on GitHub.
// ============================================================

const SHEET_ID = '11JvRXvlZAPoJTf5SwVRUD4EP32Q74Z7yd1ikZQ6NY2g';
const OPP_SHEET_ID = '18Yhe3Yiq9_eI7kBxtFOzdu6Pb0_VUx730TYjq1xPjzI';

// API key — set this in Project Settings → Script properties (do NOT paste the
// key here). Property name must be exactly: ANTHROPIC_API_KEY
const ANTHROPIC_API_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');

// Model for the "Deep infer" org-structure reasoning. If your key doesn't have
// Opus access (Deep infer's bar says "proposed" instead of "Claude reasoned"),
// change this to 'claude-sonnet-4-20250514'.
const ORG_REASONING_MODEL = 'claude-opus-4-8';

// Model for the contact-sync extraction (advanced reasoning over the account's
// notes). If your key doesn't have Opus access, change this to
// 'claude-sonnet-4-20250514'.
const SYNC_EXTRACTION_MODEL = 'claude-opus-4-8';

// Generous cap so we effectively send ALL of an account's notes while staying
// inside the model's context window.
const SYNC_MAX_NOTE_CHARS = 120000;

// Version stamp mixed into every account's change signature. Bump this when the
// note-gathering logic changes (new tabs read, different formatting) so every
// account re-analyzes once against the new, richer notes.
const SYNC_SIGNATURE_VERSION = '2';

// Sync-state columns the backend is allowed to auto-create on the Companies
// tab. `sync_signature` stores the source-content signature captured when an
// account's notes were last fully analyzed; `sync_checked` stores when. The
// front-end compares the stored signature against `listSourceAccounts` output
// to decide — with no AI cost — whether anything new needs analyzing.
const SYNC_STATE_COLUMNS = ['sync_signature', 'sync_checked'];

// ============================================================
// MAIN ENTRY POINT — handles all POST requests
// ============================================================

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.action === 'research') {
      return doResearch(payload.name, payload.company,
        payload.research_mode, payload.partner_research_instructions);
    }

    if (payload.action === 'reasonOrg') {
      return doReasonOrg(payload.company, payload.roster);
    }

    if (payload.action === 'syncOpportunities') {
      return doSyncOpportunities(payload.companyName);
    }

    if (payload.action === 'listOpportunityCompanies') {
      return doListOpportunityCompanies();
    }

    if (payload.action === 'syncPartners') {
      return doSyncPartners(payload.companyName);
    }

    if (payload.action === 'listPartnerCompanies') {
      return doListPartnerCompanies();
    }

    if (payload.action === 'listSourceAccounts') {
      return doListSourceAccounts();
    }

    var tab = payload.tab;
    var data = payload.data;
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(tab);
    if (!sheet) throw new Error('Tab not found: ' + tab);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // The sync-state columns are created on demand the first time the front-end
    // persists a signature, so no manual sheet edit is needed after deploying.
    if (tab === 'Companies') {
      headers = lcEnsureSyncColumns_(sheet, headers, data);
    }

    if (payload.action === 'append') {
      // ---- Server-side duplicate guards -----------------------------------
      // The sheet is the single source of truth, so duplicates are blocked
      // HERE regardless of which browser tab / client attempts the write.
      //
      // Companies: never create a second row for the same company name
      // (case- and whitespace-insensitive). Returns the existing row's id so
      // the client can re-point at it instead.
      if (tab === 'Companies' && data && data.name) {
        var dup = lcDedup_findByName_(sheet, headers, 'name', data.name, 'company_id');
        if (dup) {
          return jsonResponse({ ok: true, action: 'append', tab: tab,
            deduped: true, existing_company_id: dup.id });
        }
      }
      // Contacts: never create a second row for the same full_name at the
      // same company_id. (Different companies may share a person's name.)
      if (tab === 'Contacts' && data && data.full_name && data.company_id) {
        var dupC = lcDedup_findContact_(sheet, headers, data.full_name, data.company_id);
        if (dupC) {
          return jsonResponse({ ok: true, action: 'append', tab: tab,
            deduped: true, existing_contact_id: dupC.id });
        }
      }
      // ---------------------------------------------------------------------
      var row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
      sheet.appendRow(row);
      return jsonResponse({ ok: true, action: 'append', tab: tab });
    }

    if (payload.action === 'update') {
      var idColumn = payload.idColumn;
      var idValue = payload.idValue;
      var idColIndex = headers.indexOf(idColumn);
      if (idColIndex === -1) throw new Error('ID column not found: ' + idColumn);
      var allValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var rowIndex = -1;
      for (var i = 0; i < allValues.length; i++) {
        if (allValues[i][idColIndex] == idValue) { rowIndex = i; break; }
      }
      if (rowIndex === -1) throw new Error('Record not found: ' + idValue);
      var updatedRow = headers.map(function(h, i) { return data[h] !== undefined ? data[h] : allValues[rowIndex][i]; });
      sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([updatedRow]);
      return jsonResponse({ ok: true, action: 'update', tab: tab });
    }

    if (payload.action === 'delete') {
      var idColumn = payload.idColumn;
      var idValue = payload.idValue;
      var idColIndex = headers.indexOf(idColumn);
      if (idColIndex === -1) throw new Error('ID column not found: ' + idColumn);
      var allValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var rowIndex = -1;
      for (var i = 0; i < allValues.length; i++) {
        if (allValues[i][idColIndex] == idValue) { rowIndex = i; break; }
      }
      if (rowIndex === -1) throw new Error('Record not found: ' + idValue);
      sheet.deleteRow(rowIndex + 2);
      return jsonResponse({ ok: true, action: 'delete', tab: tab });
    }

    throw new Error('Unknown action: ' + payload.action);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ------------------------------------------------------------
// Duplicate-guard helpers (used by the append action)
// ------------------------------------------------------------

function lcDedup_norm_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
}

// Find a row whose `nameColumn` matches `name` (normalized). Returns
// { id: <idColumn value> } or null.
function lcDedup_findByName_(sheet, headers, nameColumn, name, idColumn) {
  var nameIdx = headers.indexOf(nameColumn);
  if (nameIdx === -1 || sheet.getLastRow() < 2) return null;
  var idIdx = headers.indexOf(idColumn);
  var want = lcDedup_norm_(name);
  if (!want) return null;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (lcDedup_norm_(values[i][nameIdx]) === want) {
      return { id: idIdx > -1 ? String(values[i][idIdx]) : '' };
    }
  }
  return null;
}

// Find a Contacts row with the same full_name (normalized) AND company_id.
// Rows already hard-merged away (merged_into set) are skipped so a merged
// duplicate can't block re-use of the surviving record.
function lcDedup_findContact_(sheet, headers, fullName, companyId) {
  var nameIdx = headers.indexOf('full_name');
  var coIdx = headers.indexOf('company_id');
  if (nameIdx === -1 || coIdx === -1 || sheet.getLastRow() < 2) return null;
  var idIdx = headers.indexOf('contact_id');
  var mergedIdx = headers.indexOf('merged_into');
  var wantName = lcDedup_norm_(fullName);
  var wantCo = String(companyId == null ? '' : companyId).trim();
  if (!wantName || !wantCo) return null;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (mergedIdx > -1 && String(values[i][mergedIdx] || '').trim()) continue;
    if (String(values[i][coIdx]).trim() !== wantCo) continue;
    if (lcDedup_norm_(values[i][nameIdx]) === wantName) {
      return { id: idIdx > -1 ? String(values[i][idIdx]) : '' };
    }
  }
  return null;
}

// ============================================================
// LIST OPPORTUNITY COMPANIES
// ============================================================

function doListOpportunityCompanies() {
  var ss = SpreadsheetApp.openById(OPP_SHEET_ID);
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

  var ss = SpreadsheetApp.openById(OPP_SHEET_ID);
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
  var ss = SpreadsheetApp.openById(OPP_SHEET_ID);
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

  var ss = SpreadsheetApp.openById(OPP_SHEET_ID);
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
  var ss = SpreadsheetApp.openById(OPP_SHEET_ID);

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

// ============================================================
// AI RESEARCH
// ============================================================

function doResearch(name, company, researchMode, partnerInstructions) {
  var prompt = buildResearchPrompt(name, company);

  // Partner qualification mode — the Partners workspace sends the qualification
  // rubric with the request; prepend it and require the partner_qualification
  // block in the output so the front-end gets a real, model-driven verdict
  // instead of falling back to keyword heuristics.
  researchMode = String(researchMode || '');
  partnerInstructions = String(partnerInstructions || '');
  if (researchMode === 'partner_qualification' && partnerInstructions) {
    prompt = partnerInstructions + '\n\n' + prompt +
      '\n\nADDITIONAL OUTPUT REQUIREMENT: inside "company_info", also include a ' +
      '"partner_qualification" object with fields: status ("QUALIFIED" | "NOT QUALIFIED"), ' +
      'score (0-100), partner_type, evidence, joint_value, overlap_risk, disqualifiers, sources.';
  }

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 5000,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 12 }],
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });
  return ContentService.createTextOutput(response.getContentText())
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AI ORG REASONING — "Deep infer" reporting structure
// Given the researched roster + company profile, work out who reports to whom.
// Truth first (public sources via web search + "reports to" notes), then
// comparable-company reasoning where this company's structure isn't public.
// ============================================================

function doReasonOrg(company, roster) {
  company = company || {};
  roster = Array.isArray(roster) ? roster : [];

  // Nothing to reason about with fewer than two people.
  if (roster.length < 2) {
    return jsonResponse({ ok: true, content: [{ type: 'text', text: '{"edges":[]}' }] });
  }

  var prompt = buildOrgReasoningPrompt(company, roster);

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: ORG_REASONING_MODEL,
      max_tokens: 6000,
      // Adaptive thinking improves reasoning and keeps the visible response as
      // clean JSON. Remove this line if your key rejects it.
      thinking: { type: 'adaptive' },
      // Let the model verify the ACTUAL reporting lines from public sources
      // before falling back to comparable-company reasoning.
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 8 }],
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  return ContentService.createTextOutput(response.getContentText())
    .setMimeType(ContentService.MimeType.JSON);
}

function buildOrgReasoningPrompt(company, roster) {
  var lines = [];
  for (var i = 0; i < roster.length; i++) {
    var p = roster[i] || {};
    var parts = [p.name || '(unknown)'];
    if (p.title) parts.push('— ' + p.title);
    var meta = [];
    if (p.functionArea) meta.push('function: ' + p.functionArea);
    if (p.seniority) meta.push('seniority: ' + p.seniority);
    if (p.note) meta.push('notes: ' + p.note);
    if (meta.length) parts.push('(' + meta.join('; ') + ')');
    lines.push('- ' + parts.join(' '));
  }

  var sizeBits = [];
  if (company.size) sizeBits.push('size: ' + company.size);
  if (company.revenue) sizeBits.push('revenue: ' + company.revenue);
  if (company.industry) sizeBits.push('industry: ' + company.industry);
  var companyProfile = (company.name || 'the company') +
    (sizeBits.length ? ' (' + sizeBits.join(', ') + ')' : '');

  return 'You are a world-class organizational-structure analyst with web search. ' +
    'Determine the most likely internal reporting hierarchy — who reports to whom — among the people on the roster below, all of whom work at ' + companyProfile + '.\n\n' +

    '=== PRIORITIES ===\n' +
    '1. ABSOLUTE TRUTH FIRST. Prefer verifiable reporting lines. Use the roster notes (an explicit "reports to X" is the strongest signal) and web search of public sources (LinkedIn, TheOrg, the company site, press releases) to establish the ACTUAL structure wherever you can.\n' +
    '2. COMMON SENSE SECOND. Where this specific company\'s structure is not publicly available, reason from how COMPARABLE companies — similar size, similar industry, similar role mix — are typically organized, and make an educated guess. Label these clearly and give them lower confidence.\n\n' +

    '=== RULES ===\n' +
    '- A manager MUST be one of the people on the roster. Do NOT invent people or reference anyone not listed. Return only reporting lines between roster members.\n' +
    '- Each person reports to AT MOST ONE manager. The single most senior person tops the tree and has no manager (do not emit an edge for them).\n' +
    '- No cycles: a person cannot report (directly or indirectly) to someone who reports to them.\n' +
    '- Attach each person up their own function where possible (a Finance analyst reports up Finance), to the nearest senior above them, unless a public source shows otherwise.\n' +
    '- "basis" is one of: "confirmed" (a stated "reports to" or a public source), "inferred" (this company\'s own titles/functions, no explicit source), or "comparable" (an educated guess from similar-size / similar-role companies because this company\'s own structure was not found).\n' +
    '- "confidence" is 0.0-1.0: confirmed 0.85-0.97, inferred 0.5-0.8, comparable 0.30-0.60. Never let a comparable guess read as high-confidence.\n' +
    '- "rationale" is ONE sentence: the specific evidence (cite the source) or the comparable-company logic.\n\n' +

    '=== ROSTER (' + roster.length + ' people) ===\n' +
    lines.join('\n') + '\n\n' +

    '=== OUTPUT ===\n' +
    'Respond with ONLY a valid JSON object — no markdown, no backticks, no preamble:\n' +
    '{"edges":[{"report":"<exact name from roster>","manager":"<exact name from roster>","confidence":0.0,"basis":"confirmed|inferred|comparable","rationale":"one sentence"}]}\n' +
    'Use the EXACT names as they appear in the roster for both "report" and "manager". If you cannot place someone, omit them. Return {"edges":[]} if you cannot determine any lines.';
}

// ============================================================
// RESEARCH PROMPT — World-Class Multi-Source Methodology
// ============================================================

function buildResearchPrompt(name, company) {
  var companyLine = company ? ' at ' + company : '';
  var companyInstruction = company
    ? 'The target company is "' + company + '". Verify this is correct during your research. If the person actually works at a different company, use the correct one.'
    : 'The company was not provided. Your first task is to identify which company this person currently works at.';

  return 'You are a world-class lead research agent with access to web search. Conduct a comprehensive, multi-source research operation on "' + name + '"' + companyLine + '.\n\n' +
    companyInstruction + '\n\n' +

    '=== RESEARCH METHODOLOGY — MULTI-SOURCE, 9-PHASE APPROACH ===\n\n' +

    '=== SOURCE PRIORITY & CROSS-REFERENCING STRATEGY ===\n' +
    'Use these sources in this priority order. Cross-reference across at least 3 sources before marking anything as CONFIRMED.\n\n' +
    'TIER 1 — PRIMARY SOURCES (start here):\n' +
    '- LinkedIn (linkedin.com) — Professional history, current role, education, connections. The single best starting point. Always search LinkedIn first.\n' +
    '- Company website — Leadership page, team page, author/bio pages, press releases. The highest-trust source for current title.\n' +
    '- The Org (theorg.com) — Pre-built org charts for thousands of companies. Search here for instant org structure.\n\n' +

    'TIER 2 — VERIFICATION SOURCES (cross-reference):\n' +
    '- OpenCorporates (opencorporates.com) — Corporate officer records and company affiliations worldwide. Confirms executive titles and board memberships.\n' +
    '- Crunchbase (crunchbase.com) — Leadership team, key personnel, funding history, company metrics.\n' +
    '- Owler (owler.com) — Leadership profiles, company structure, competitive landscape.\n' +
    '- ZoomInfo (zoominfo.com) — Org structure, department mapping, title verification.\n\n' +

    'TIER 3 — EMAIL & CONTACT DISCOVERY:\n' +
    '- Hunter.io (hunter.io) — Email format patterns and verified addresses by domain.\n' +
    '- RocketReach (rocketreach.co) — Email and phone lookups with confidence scoring.\n' +
    '- Apollo.io (apollo.io) — Contact database with email verification.\n' +
    '- EmailRep.io (emailrep.io) — Email reputation and validity checking.\n' +
    '- That\'s Them (thatsthem.com) — Free aggregator for email, phone, address associations.\n\n' +

    'TIER 4 — PUBLIC RECORDS & IDENTITY VERIFICATION:\n' +
    '- FastPeopleSearch (fastpeoplesearch.com) — Public records for addresses, phone numbers, relatives, associates.\n' +
    '- That\'s Them (thatsthem.com) — Aggregates public records from multiple data sources.\n\n' +

    'CROSS-REFERENCING RULE: Start with LinkedIn for the professional baseline. Validate corporate roles on OpenCorporates or The Org. Use Crunchbase/Owler for company intel. Use Hunter/RocketReach for email patterns. Use public records aggregators to confirm location. Every fact should appear in at least 2 sources to be marked CONFIRMED.\n\n' +

    'PHASE 1 — IDENTITY VERIFICATION\n' +
    '- Search LinkedIn: "' + name + '"' + (company ? ' "' + company + '"' : '') + '\n' +
    '- Pull their LinkedIn URL, exact current title, location, education, certifications\n' +
    '- Search The Org (theorg.com) for the company — check if this person appears in their org chart\n' +
    '- Search the company website for an author/bio/team page\n' +
    '- Search OpenCorporates if the person holds an executive or board title\n' +
    '- Cross-reference: title must match on at least 2 sources to mark as VERIFIED\n' +
    '- Search FastPeopleSearch or That\'s Them to verify location/city\n' +
    '- If you cannot confirm the person exists at this company, say so clearly but still fill in whatever you CAN find.\n\n' +

    'PHASE 2 — COMPANY INTELLIGENCE\n' +
    '- Search Crunchbase for company profile: funding, revenue, employee count, key personnel\n' +
    '- Search Owler for leadership profiles and competitive landscape\n' +
    '- Search for the current CEO and any leadership transitions in the last 12 months\n' +
    '- Check OpenCorporates for corporate officers and recent filings\n' +
    '- Find most recent revenue, employee count, and key financial metrics\n' +
    '- Search for recent acquisitions, restructuring, funding rounds, or strategic pivots\n' +
    '- Check the company events/conferences page for upcoming presence\n' +
    '- Identify key partners and technology alliances\n\n' +

    'PHASE 3 — CONTENT FOOTPRINT\n' +
    '- Find authored content: articles, blog posts, videos, podcasts, webinars, conference talks\n' +
    '- Search Google: "' + name + '" ' + (company ? '"' + company + '"' : '') + ' article OR blog OR podcast OR webinar OR conference\n' +
    '- Cap at the 3-5 most recent and relevant pieces\n' +
    '- Extract 2-3 theme keywords from the pattern\n' +
    '- Note recency — content older than 2 years gets lower weight\n' +
    '- Include the source URL for every piece of content found\n\n' +

    'PHASE 4 — EMAIL DISCOVERY & VALIDATION\n' +
    '- Search Hunter.io for the company domain email pattern and confidence percentage\n' +
    '- Search RocketReach for the specific person\'s verified email\n' +
    '- Search Apollo.io for the person\'s contact info\n' +
    '- Check That\'s Them for any publicly associated email addresses\n' +
    '- If you find an email, check it against the Hunter.io domain pattern\n' +
    '- Assign confidence: HIGH (found on 2+ sources or verified), MEDIUM (matches format pattern from Hunter/RocketReach but only 1 source), LOW (guessed or format unknown)\n\n' +

    'PHASE 5 — ORG CHART CONSTRUCTION\n' +
    '- FIRST: Search The Org (theorg.com) for the company — they may already have a complete org chart\n' +
    '- Search LinkedIn by company, filter by department and title to map the hierarchy\n' +
    '- Search Crunchbase for the leadership team\n' +
    '- Search Owler for leadership profiles\n' +
    '- Search ZoomInfo for department structure\n' +
    '- Check the company leadership/about page\n' +
    '- Identify the target\'s functional department based on their title\n' +
    '- Find their direct manager by searching for the title one level up\n' +
    '- Search for peers — other people with similar titles. Cap at 3-5. Each must have a source URL.\n' +
    '- Map upward 2-3 levels to reach C-suite\n' +
    '- Mark every node as CONFIRMED (found on company site, The Org, press release, or 2+ sources) or INFERRED (title pattern match only)\n' +
    '- ONLY include nodes where you found a REAL person\'s name. NEVER create placeholder nodes.\n' +
    '- If you cannot find a real name for a role, OMIT that node entirely — set it to null\n\n' +

    'PHASE 6 — CONFIDENCE ASSESSMENT\n' +
    '- Score each of these 12 checks as true (passed) or false (failed):\n' +
    '  1. Name + title confirmed on LinkedIn\n' +
    '  2. Title confirmed on company site or The Org\n' +
    '  3. Email format matches dominant pattern (Hunter.io/RocketReach)\n' +
    '  4. Published content verified\n' +
    '  5. Certifications confirmed\n' +
    '  6. Education confirmed\n' +
    '  7. Function head identified (real name found)\n' +
    '  8. Peer directors identified (at least 1 real name)\n' +
    '  9. Direct manager name sourced (real name found)\n' +
    '  10. X/Twitter presence found\n' +
    '  11. Personal website/blog found\n' +
    '  12. Conference speaking engagement confirmed\n' +
    '- Overall confidence = percentage of checks passed\n\n' +

    'PHASE 7 — STRATEGIC ASSESSMENT\n' +
    '- Budget holder: Who owns spending authority for this function? Name them with title and source.\n' +
    '- Decision distance: How many levels to budget holder? Is the target an evaluator, champion, or approver?\n' +
    '- Alternative entry points: Which confirmed peers are legitimate lateral entries? Why?\n' +
    '- Gatekeeper vs. authority: Screener or champion? Base on public behavior.\n' +
    '- Timing window: Active CEO transition, acquisition, reorg, or product launch?\n' +
    '- Write these into decision_role and outreach_angle fields.\n\n' +

    'PHASE 8 — OUTREACH PLAYBOOK\n' +
    '- Build a 7-8 step sequenced outreach plan\n' +
    '- Write 4 message templates (LinkedIn Connection, Warm Email, Direct Cold Email, Follow-Up)\n' +
    '- Build a priority matrix: target = PRIMARY, confirmed peers = SECONDARY\n\n' +

    'PHASE 9 — SOURCE TRACKING\n' +
    '- Compile every source URL used\n' +
    '- Each source must have a descriptive label and working URL\n' +
    '- Include which tier the source came from (LinkedIn, The Org, Hunter.io, etc.)\n\n' +

    '=== OUTPUT FORMAT ===\n\n' +
    'Respond with ONLY a valid JSON object. No markdown, no backticks, no preamble.\n\n' +
    '{\n' +
    '  "title": "exact current job title or null",\n' +
    '  "location": "City, State/Country or null",\n' +
    '  "email": "their email address or null",\n' +
    '  "email_confidence": "HIGH or MEDIUM or LOW",\n' +
    '  "linkedin_url": "full LinkedIn profile URL or null",\n' +
    '  "education": "degree, school or null",\n' +
    '  "certifications": ["cert1", "cert2"],\n' +
    '  "verified": true,\n' +
    '  "overall_confidence": 83,\n' +
    '  "background": "2-4 sentence professional background based on verified information.",\n' +
    '  "decision_role": "Detailed assessment of decision-making authority. Budget authority, reporting line, cross-functional influence. Strategic read: budget holder, decision distance, gatekeeper vs. authority.",\n' +
    '  "outreach_angle": "Specific outreach recommendation referencing recent content. What to lead with, what to avoid, timing windows, alternative entry points.",\n' +
    '  "themes": ["Theme1", "Theme2", "Theme3"],\n' +
    '  "dual_role": "secondary role or null",\n' +
    '  "recent_activities": [{"type": "article", "text": "description", "date": "YYYY-MM-DD", "source_url": "url"}],\n' +
    '  "upcoming_events": [{"event_name": "name", "event_date": "YYYY-MM-DD", "status": "Speaker", "location": "city", "source_url": "url", "insight_callout": "tip"}],\n' +
    '  "affiliations": [{"name": "org name", "source_url": "url"}],\n' +
    '  "org_chart": {\n' +
    '    "topExec": {"name": "real name ONLY or null", "title": "CEO", "verified": "CONFIRMED or INFERRED", "source_url": "url", "note": "context"},\n' +
    '    "functionHead": {"name": "real name ONLY or null", "title": "title", "verified": "CONFIRMED or INFERRED", "source_url": "url", "note": "context"},\n' +
    '    "manager": {"name": "real name ONLY or null", "title": "title", "verified": "CONFIRMED or INFERRED", "source_url": "url", "note": "context"},\n' +
    '    "peers": [{"name": "real name", "title": "title", "verified": "CONFIRMED or INFERRED", "source_url": "url", "note": "context"}]\n' +
    '  },\n' +
    '  "playbook_steps": [{"step_order": 1, "title": "step", "detail": "instruction"}],\n' +
    '  "templates": [\n' +
    '    {"label": "LinkedIn Connection", "color": "primary", "subject": "", "text_body": "message"},\n' +
    '    {"label": "Warm Email", "color": "amber", "subject": "subject", "text_body": "body"},\n' +
    '    {"label": "Direct Cold Email", "color": "green", "subject": "subject", "text_body": "body"},\n' +
    '    {"label": "Follow-Up", "color": "gray", "subject": "", "text_body": "body"}\n' +
    '  ],\n' +
    '  "priority_matrix": [{"contact_name": "' + name + '", "channel": "Email", "priority": "PRIMARY", "email_confidence": "HIGH", "notes": "reason"}],\n' +
    '  "confidence_checks": [\n' +
    '    {"label": "Name + title confirmed on LinkedIn", "pass": true},\n' +
    '    {"label": "Title confirmed on company site or The Org", "pass": false},\n' +
    '    {"label": "Email format matches pattern (Hunter/RocketReach)", "pass": true},\n' +
    '    {"label": "Published content verified", "pass": true},\n' +
    '    {"label": "Certifications confirmed", "pass": false},\n' +
    '    {"label": "Education confirmed", "pass": true},\n' +
    '    {"label": "Function head identified (real name)", "pass": true},\n' +
    '    {"label": "Peer directors identified", "pass": true},\n' +
    '    {"label": "Direct manager name sourced", "pass": false},\n' +
    '    {"label": "X/Twitter presence found", "pass": false},\n' +
    '    {"label": "Personal website/blog found", "pass": false},\n' +
    '    {"label": "Conference speaking confirmed", "pass": true}\n' +
    '  ],\n' +
    '  "sources": [{"label": "Source description (Tier: LinkedIn/TheOrg/Hunter/etc)", "url": "url"}],\n' +
    '  "company_info": {\n' +
    '    "name": "' + (company || 'determine from research') + '",\n' +
    '    "ticker": "ticker or null",\n' +
    '    "tagline": "one-line description",\n' +
    '    "industry": "industry",\n' +
    '    "hq": "City, State/Country",\n' +
    '    "size": "employee count",\n' +
    '    "revenue": "revenue or null",\n' +
    '    "founded": "year",\n' +
    '    "ceo": "CEO name and start date",\n' +
    '    "ceo_prior": "Previous CEO or null",\n' +
    '    "core_focus": "primary business focus",\n' +
    '    "partners": ["partner1", "partner2"],\n' +
    '    "website": "url",\n' +
    '    "email_domain": "domain.com",\n' +
    '    "email_format": "first.last@domain.com",\n' +
    '    "email_format_confidence": 85,\n' +
    '    "badges": ["recognition"],\n' +
    '    "achievements": [{"label": "title:", "description": "detail", "source_url": "url", "date": "YYYY-MM-DD"}]\n' +
    '  }\n' +
    '}\n\n' +

    '=== CRITICAL RULES ===\n\n' +
    '1. VERIFICATION FIRST: Every fact must come from a public source. If unconfirmed, use null.\n' +
    '2. SOURCE EVERYTHING: Every claim must have a source_url. No exceptions.\n' +
    '3. NO FABRICATION: Never invent names or relationships.\n' +
    '4. CONFIRMED vs INFERRED: CONFIRMED = found on 2+ sources. INFERRED = single source or title pattern match.\n' +
    '5. COMPLETENESS OVER SPEED: Gaps with null values are better than fabricated data.\n' +
    '6. OUTREACH ANGLE LAST: Write only after all themes and activity are locked.\n' +
    '7. Valid JSON only. No trailing commas, no comments, no markdown fences.\n' +
    '8. ONLY the JSON object. No text before or after.\n' +
    '9. MINIMUM EFFORT: Try at least 3 different search queries before giving up on any field.\n' +
    '10. CONFIDENCE FLOOR: If the person EXISTS at the company, overall_confidence must be at least 30.\n' +
    '11. NO PLACEHOLDER NODES: Never return "Unknown", "Cannot determine" in org chart. Omit the node.\n' +
    '12. SPARSE BUT ACCURATE: Sparse + accurate beats fabricated.\n' +
    '13. MULTI-SOURCE CROSS-REFERENCE: CONFIRMED requires 2+ independent sources.\n' +
    '14. TIER LABELING: Note which tier each source came from.\n\n' +

    '=== PARTNER QUALIFICATION (only when instructed above) ===\n' +
    'If this prompt began with partner-development qualification instructions, follow them exactly and include the "partner_qualification" object inside "company_info" as instructed. Otherwise omit it.';
}

// ============================================================
// UTILITY
// ============================================================

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
