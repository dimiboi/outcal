# Microsoft Graph API Access from CLI

## Goal
Query Outlook calendar data via Microsoft Graph from local scripts. The MSAL auth
setup generalizes to other Graph data (mail, contacts, etc.) by adding the matching
delegated scope, but today only calendar is wired up — the token is scoped to
`Calendars.Read`.

## Working Setup
Auth uses the **Microsoft Graph PowerShell** public client ID
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`) via MSAL Python — a public native client
with an `http://localhost` redirect that BCG has pre-consented for delegated Graph
scopes including `Calendars.Read`. This is the one client ID known to work here; see
"What Did Not Work" before trying another.

`fetch_calendar.py` queries `/me/calendarView`, which expands recurring meetings into
one row per occurrence over the requested window.

**IMPORTANT: run from the registered BCG Mac.** Conditional Access binds at token
*acquisition*, so acquiring elsewhere fails (`AADSTS53003`).

```bash
# Default window is -365d .. +180d; --top default 100 (calendarView 504s on large pages over wide windows)
uv run fetch_calendar.py --start 2026-05-06T00:00:00Z --end 2026-06-06T23:59:59Z
uv run fetch_calendar.py --category Travel   # single category, filtered server-side
jq '{subject, start: .start.dateTime, categories}' data/graph.jsonl
```

A **single** `--category` filters server-side. Multiple categories OR'd server-side
return `ErrorInternalServerError` — for that, fetch unfiltered and filter with `jq`.

## Gotchas
- Always run Python via `uv run` (e.g. `uv run python -m py_compile fetch_calendar.py`).
- Once minted on a registered device, a Graph token is accepted from any IP (the
  audience is `graph.microsoft.com`; CA binds at acquisition, not use).

## What Did Not Work
- **M365 MCP connector** (incl. the calendar tools loaded in this session) — no category filter, which is the whole reason this script exists. Don't reach for it for category-scoped queries.
- **Other token sources** — stick with the Graph PowerShell client ID: the Azure CLI app ID has no `Calendars.Read` consent (`ErrorAccessDenied`), and the Graph Explorer client ID is SPA-only and rejects `http://localhost` (`AADSTS900971`).
