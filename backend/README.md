# LeadCheck Pro — Backend (Google Apps Script)

The app's front-end (`index.html`) talks to a Google Apps Script Web App
(`CONFIG.WEB_APP_URL`) that proxies to the Anthropic API and reads/writes the
Google Sheets. That script lives in your Apps Script project, **not** in this
repo, so its API key and Sheet IDs are never committed here.

## `reason-org.gs` — the `reasonOrg` action

This add-on powers the Account Map's **🧠 Deep infer** button. After the
front-end researches each contact online, it sends the researched roster plus
the company profile to this action, and an AI reasoning model works out the
likely reporting hierarchy — prioritising verifiable truth (roster notes +
public sources via web search) and falling back to comparable-company reasoning
(similar size / similar roles) where this company's own structure isn't public.

### Install (one-time)

1. Open your existing Apps Script project (the one behind `CONFIG.WEB_APP_URL`).
2. Paste the two functions from [`reason-org.gs`](./reason-org.gs)
   (`doReasonOrg`, `buildOrgReasoningPrompt`) and the `ORG_REASONING_MODEL`
   constant into `Code.gs`. They reuse the `ANTHROPIC_API_KEY` constant and the
   `jsonResponse` helper already defined there.
3. In `doPost(e)`, add this branch next to the other action checks:

   ```javascript
   if (payload.action === 'reasonOrg') {
     return doReasonOrg(payload.company, payload.roster);
   }
   ```

4. **Deploy → Manage deployments → Edit → New version → Deploy.**

## Storing the API key securely (Script properties)

Never hard-code your Anthropic API key in the source. Store it in the project's
**Script properties** and read it at runtime:

1. In the Apps Script editor, open **⚙️ Project Settings** (left sidebar).
2. Scroll to **Script Properties** → **Edit script properties** → **Add script
   property**.
3. Property = `ANTHROPIC_API_KEY`, Value = your key (`sk-ant-…`). **Save.**
4. In `Code.gs`, replace the hard-coded constant with a lookup:

   ```javascript
   // Before:
   // const ANTHROPIC_API_KEY = 'sk-ant-…';   // ← never commit a real key

   // After — reads from Project Settings → Script properties:
   const ANTHROPIC_API_KEY =
     PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
   ```

Every action (`research`, `syncOpportunities`, `syncPartners`, `reasonOrg`)
already references the `ANTHROPIC_API_KEY` constant, so they all pick up the
stored value automatically — no other edits needed. Script properties are
project-scoped, are never exposed to the front-end, and are never committed to
git. If a key was ever committed or shared, rotate it in the Anthropic console
and store the new one this way.

### Graceful fallback

The front-end degrades gracefully. Until this action is deployed (or if the
call fails), **🧠 Deep infer** automatically uses its built-in local reasoning
engine instead — so the button keeps working, just without the live AI pass.
Once `reasonOrg` is deployed, the preview bar shows "reasoned by Claude".

### Notes

- `ORG_REASONING_MODEL` defaults to `claude-opus-4-8` (strongest reasoning). If
  your API key doesn't have access, change it to a model it does — e.g.
  `claude-sonnet-4-20250514`, which the rest of the backend already uses.
- The `thinking: { type: 'adaptive' }` line improves reasoning and keeps the
  response as clean JSON; remove it if your key rejects it.
- `web_search_20250305` matches the version the existing `research` action uses.
  On newer models you can switch to `web_search_20260209` for dynamic filtering.

## Troubleshooting: "Refresh failed — Research completed but returned no text content"

A contact **Refresh** (and the initial lookup) calls the `research` action, which
proxies to Claude with **web search + extended thinking**. Occasionally the model
spends its whole token budget on thinking and search tool-calls and stops
(`stop_reason` = `max_tokens` or `pause_turn`) **before** it writes a final text
block. The response then has content blocks but no text to parse, which the UI
reported as *"Research completed but returned no text content."*

This is fixed on **two** levels:

1. **Front-end (already shipped in `index.html`).** `callResearchAPI` now retries
   automatically (up to 3 attempts with backoff) when a response comes back with
   no text / unparseable JSON, and logs the `stop_reason` to the console. Because
   the empty-text outcome is transient, a retry almost always returns the JSON —
   so the refresh succeeds without any backend change.

2. **Backend (recommended, for a permanent fix).** In your `research` action:
   - **Give the model room to answer.** Raise `max_tokens` (e.g. `4096`+ for the
     research JSON) so thinking + search don't consume the whole budget before
     the text block.
   - **Handle `pause_turn`.** When `web_search` runs long, Claude can return
     `stop_reason: "pause_turn"`. Continue the turn by sending the returned
     `content` back as the next `assistant` message and calling the API again in
     a loop until `stop_reason` is `end_turn` (or a small max-iterations cap),
     then return the accumulated message. Without this, a paused turn can arrive
     with tool-use blocks but no final text.
   - **Optional belt-and-suspenders.** If a completed turn still has no `text`
     block, do one more call **without tools** that asks the model to emit just
     the JSON, so the front-end always receives text.

   The front-end passes the raw Anthropic message straight through, so it already
   surfaces `stop_reason` in the console — check there to confirm which case
   (`max_tokens` vs `pause_turn`) you're hitting before tuning.
