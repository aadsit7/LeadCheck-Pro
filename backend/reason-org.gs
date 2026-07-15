// ============================================================
// LeadCheck Pro — Backend add-on: the `reasonOrg` action
// ============================================================
//
// This powers the Account Map's "🧠 Deep infer" button: after the front-end
// researches each contact online, it sends the researched roster + company
// profile here and asks an AI reasoning model to work out who reports to whom.
// It prioritises verifiable truth (roster notes + public sources via web
// search), and falls back to comparable-company reasoning (similar size /
// similar roles) where this company's own structure isn't public.
//
// ------------------------------------------------------------
// HOW TO INSTALL (one-time)
// ------------------------------------------------------------
// 1. Open your existing Apps Script project (the one behind CONFIG.WEB_APP_URL)
//    and paste the two functions below (`doReasonOrg` and
//    `buildOrgReasoningPrompt`) plus the `ORG_REASONING_MODEL` constant into
//    Code.gs. They reuse the `ANTHROPIC_API_KEY` constant and the
//    `jsonResponse` helper already defined there — no secret is duplicated here.
// 2. In doPost(e), add this branch alongside the other action checks
//    (e.g. right after the `research` branch):
//
//        if (payload.action === 'reasonOrg') {
//          return doReasonOrg(payload.company, payload.roster);
//        }
//
// 3. Deploy → Manage deployments → Edit → New version → Deploy.
//
// The front-end degrades gracefully: if this action isn't deployed yet (or the
// call fails), "🧠 Deep infer" automatically falls back to its built-in local
// reasoning engine, so nothing breaks in the meantime.
// ============================================================

// Strongest widely-available reasoning model. Change this if your API key
// doesn't have access to it (e.g. 'claude-sonnet-4-20250514', which the rest of
// this backend already uses).
var ORG_REASONING_MODEL = 'claude-opus-4-8';

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
      // Adaptive thinking improves reasoning quality and keeps the visible
      // response as clean JSON. Remove this line if your key rejects it.
      thinking: { type: 'adaptive' },
      // Let the model verify the ACTUAL reporting lines from public sources
      // before falling back to comparable-company reasoning.
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 8 }],
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  // Pass the raw Anthropic response through — the front-end unwraps the
  // content blocks and extracts the JSON, exactly like the `research` action.
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
