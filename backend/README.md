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
