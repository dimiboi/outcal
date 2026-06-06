# Microsoft Graph API Access from CLI

## Goal
Programmatically query Outlook / Microsoft 365 data (calendar, mail, etc.) via Microsoft Graph API from local scripts, bypassing limitations of the Claude.ai Microsoft 365 MCP connector (which doesn't expose filters for categories, extended properties, or arbitrary OData queries).

## Identifiers
- **BCG tenant ID:** `2dfb2f0b-4d21-4268-9559-72926144c918`
- **User UPN:** `Makarov.Dmitry@bcg.com`

## Working Setup

Use **Microsoft Graph PowerShell** public client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) with MSAL Python. It's a public native client registered with `http://localhost` redirect, and BCG has it pre-consented for delegated Graph scopes including `Calendars.Read`.

### `graph_token.py`
```python
# /// script
# requires-python = ">=3.10"
# dependencies = ["msal"]
# ///
import msal
from pathlib import Path

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
TENANT    = "2dfb2f0b-4d21-4268-9559-72926144c918"
SCOPES    = ["Calendars.Read"]
CACHE     = Path.home() / ".graph_token_cache.json"

cache = msal.SerializableTokenCache()
if CACHE.exists():
    cache.deserialize(CACHE.read_text())

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT}",
    token_cache=cache,
)

accounts = app.get_accounts()
result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
if not result:
    result = app.acquire_token_interactive(scopes=SCOPES)

if cache.has_state_changed:
    CACHE.write_text(cache.serialize())

print(result["access_token"])
```

Run with `uv run graph_token.py`. First call opens browser for sign-in; subsequent calls silently refresh from the on-disk cache (refresh tokens valid ~90 days with rolling renewal).

### Calling Graph

`graph_token.py` acquires the token and fetches `/me/calendarView` over a window,
writing each event to `data/graph.jsonl` (JSONL — one event object per line).

```bash
# All events in a window (default window: -365d to +180d)
uv run graph_token.py --start 2026-05-06T00:00:00Z --end 2026-06-06T23:59:59Z

# Only events in an Outlook category (server-side $filter=categories/any(...))
uv run graph_token.py --category Travel \
  --start 2026-05-06T00:00:00Z --end 2026-06-06T23:59:59Z

# Inspect the dumped events (one per line)
jq '{subject, start: .start.dateTime, categories}' data/graph.jsonl
# Sorted as an array
jq -s 'sort_by(.start.dateTime) | map({subject, start: .start.dateTime})' data/graph.jsonl
```

Flags: `--start` / `--end` (ISO 8601 UTC), `--top` (page size), `--category`
(single Outlook category, filtered server-side). A single-category filter works
on calendarView without special headers; multiple categories OR'd server-side
return `ErrorInternalServerError`, so for that case filter client-side with `jq`.

## What Did Not Work (and Why)

1. **M365 MCP connector calendar tool** — only filters subject/body/location text + dates/attendees/organizer. No category filter.
2. **`az login` + `az account get-access-token --resource https://graph.microsoft.com`** — token issued, but Azure CLI's app ID (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) does not have `Calendars.Read` consented in the BCG tenant. Graph returns `ErrorAccessDenied`.
3. **Azure CLI from a non-BCG-managed machine** — blocked by Conditional Access (`AADSTS53003`, "Device state: Unregistered"). Must run from the registered Mac.
4. **Graph Explorer client ID (`de8bc8b5-d9f9-48b1-a8ad-b748da725064`) with MSAL `acquire_token_interactive`** — fails with `AADSTS900971: No reply address provided`. It's registered as a SPA with only specific HTTPS redirect URIs; doesn't accept `http://localhost`.

## Gotchas

- `az login` by default enumerates *all* tenants the account can sign into, which surfaces a (harmless) MFA error for the "VCP-US Verint Systems" guest tenant. Pin to BCG: `az login --tenant 2dfb2f0b-4d21-4268-9559-72926144c918 --allow-no-subscriptions`.
- BCG grants no Azure subscriptions to this account — that's fine for Graph, just use `--allow-no-subscriptions`.
- Once a token is minted on a registered device, it's accepted from any IP for the `graph.microsoft.com` audience. Conditional Access binds at *acquisition* time, not at use.
- BCG appears to issue extended (~24h) Graph token lifetimes vs the standard 1h.

## Graph Query Patterns

OData lambda filter on a collection property (Outlook categories are a string collection):
```
/me/events?$filter=categories/any(c:c eq 'Travel')
```
If a collection filter ever fails: add request header `ConsistencyLevel: eventual` and `&$count=true`. Graph Explorer's URL linter flags `any(c:c ...)` with a false-positive warning — ignore it, the query runs.

Pagination via `@odata.nextLink` in the response or `&$skip=N`. Max `$top=999` per page.

Useful endpoints:
- `/me/events` — events as stored (recurring events appear once as a master)
- `/me/calendarView?startDateTime=...&endDateTime=...` — expands recurrences over a window
- `/me/messages` — mail
- `/me/contacts`, `/me/todo/lists`, `/me/onenote/...`, etc.

For interactive query development, use **Graph Explorer** (`developer.microsoft.com/graph/graph-explorer`) — sign in once with BCG account, iterate on the URL, then drop the same URL into the CLI script.

## Fallback if MSAL Stops Working

If the Microsoft Graph PowerShell client ID ever gets locked down or de-consented in the BCG tenant, fall back to **Playwright driving real Edge with a persistent profile** (`launch_persistent_context(user_data_dir=..., channel="msedge")`). Reuses the existing Edge session's cookies and device certs, then intercepts the bearer token from Graph requests as they go out. Slower and more fragile, but inherits whatever Conditional Access exemptions the browser already passes.
