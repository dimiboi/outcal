# graph — Outlook calendar via Microsoft Graph from the CLI

`fetch_calendar.py` pulls your Outlook calendar from the Microsoft Graph API and
writes one event per line to `data/graph.jsonl`. It queries
[`/me/calendarView`](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview),
so recurring meetings are expanded into one row per occurrence over the requested
window.

Auth is delegated MSAL against the **Microsoft Graph PowerShell** public client, so
there's no app registration or secret to manage — first run opens a browser, after
that the token is cached and refreshed silently.

## Prerequisites

- **[uv](https://docs.astral.sh/uv/)** — runs the script and resolves its inline
  dependencies (`msal`, `httpx`, `tenacity`). No virtualenv or `pip install` needed.
- **The registered BCG Mac.** Conditional Access binds at token *acquisition*, so
  the first interactive sign-in must happen on the enrolled device. Acquiring from
  anywhere else fails with `AADSTS53003`. (Once minted, the token is accepted from
  any IP — the audience is `graph.microsoft.com` and CA only gates acquisition.)

## Usage

Always invoke via `uv run`:

```bash
# Default window: 365 days back .. 180 days ahead, 100 events per page
uv run fetch_calendar.py

# Explicit window
uv run fetch_calendar.py --start 2026-05-06T00:00:00Z --end 2026-06-06T23:59:59Z

# Filter to a single Outlook category (server-side)
uv run fetch_calendar.py --category Travel

# Look up specific events by id (prints to stdout; does NOT touch data/graph.jsonl)
uv run fetch_calendar.py --id AAMkAGU2... | jq '{subject, start: .start.dateTime}'
uv run fetch_calendar.py --id AAMk...1 AAMk...2   # several at once
```

The first run opens a browser for sign-in; the token is cached at
`~/.graph_token_cache.json` and refreshed silently on later runs.

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--id` | none | Look up one or more events by id via [`/me/events/{id}`](https://learn.microsoft.com/en-us/graph/api/event-get) and print each to stdout. A lookup mode, not a dump — see below. Can't combine with the calendarView flags (`--start`/`--end`/`--top`/`--category`). |
| `--start` | 365 days ago (UTC, midnight) | ISO 8601 `startDateTime` |
| `--end` | 180 days ahead (UTC, midnight) | ISO 8601 `endDateTime` |
| `--top` | `100` | Page size. Graph allows up to 999, but `calendarView` returns `504` on large pages over wide windows — keep it modest. |
| `--category` | none | Restrict to one Outlook category, filtered server-side. |
| `--max-retries` | `5` | Retry a page on transient errors (`429`/`503`/`504`, timeouts) with exponential backoff that honors `Retry-After`. `0` disables. |

A **single** `--category` is filtered server-side. OR'ing multiple categories
server-side returns `ErrorInternalServerError` — instead, fetch unfiltered and
filter locally:

```bash
uv run fetch_calendar.py
jq 'select(.categories | index("Travel") or index("Client"))' data/graph.jsonl
```

## Output

Each line of `data/graph.jsonl` is a full Graph
[event](https://learn.microsoft.com/en-us/graph/api/resources/event) object
(`subject`, `start`, `end`, `categories`, `attendees`, `location`, `organizer`,
`webLink`, …). The write is atomic — output goes to a `.tmp` file and is swapped in
only on success, so a failed or interrupted run never clobbers good data.

Inspect with `jq`:

```bash
jq '{subject, start: .start.dateTime, categories}' data/graph.jsonl
```

`data/` is gitignored — calendar contents stay local.

### Fetching by id (`--id`)

`--id` is a non-destructive lookup, not a dump: it prints each requested event as a
JSON line to **stdout** (same shape as a `graph.jsonl` line) and never writes
`data/graph.jsonl`, so a one-off lookup can't clobber a full calendar pull. Progress
and per-id errors go to **stderr**, keeping stdout clean for `jq`. Pull an id straight
out of the dump:

```bash
uv run fetch_calendar.py --id "$(jq -r '.id' data/graph.jsonl | head -1)" | jq .
```

Pass several ids to fetch them in order. Missing or malformed ids (`ErrorItemNotFound`,
`ErrorInvalidIdMalformed`) are reported on stderr and skipped; the rest still print, and
the process exits non-zero if any id failed. Note that ids from `calendarView` are tied
to a mailbox — they aren't portable across accounts.

One id shape can't be fetched this way: ids containing a `/` (typically recurrence
exceptions or externally-synced events). Microsoft Graph routes the decoded slash as a
path separator in `/me/events/{id}` no matter how it's encoded — neither `%2F` nor
double-encoding works ([MS Q&A](https://learn.microsoft.com/en-us/answers/questions/5516419/how-to-safely-include-slashes-in-graph-api-request)).
Such ids are skipped with an explanatory message; their full event object is already in
the `calendarView` dump, so read them from there (e.g. `jq 'select(.id == "…")'`).

`--id` is a standalone lookup mode. The calendarView flags — `--start`, `--end`,
`--top`, `--category` — only shape the dump query, so passing any of them with `--id`
is **rejected** with a usage error (exit 2) rather than silently ignored. `--max-retries`
works in both modes and still governs the per-id retry/backoff.

## How auth works

The script uses the Microsoft Graph PowerShell public client ID
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`) against the BCG tenant via MSAL Python. It's
a public native client with an `http://localhost` redirect that BCG has pre-consented
for delegated Graph scopes including `Calendars.Read`. The token is scoped to
`Calendars.Read` only.

The same setup generalizes to other Graph data (mail, contacts, etc.) by adding the
matching delegated scope to `SCOPES` — but today only calendar is wired up.

## Troubleshooting

- **`AADSTS53003` on sign-in** — you're not on the registered BCG Mac, or the device
  fell out of Conditional Access compliance. Acquire the token on the enrolled device.
- **`504` / timeouts** — transient ones are retried automatically with exponential
  backoff (honoring `Retry-After`); tune with `--max-retries`. If they *persist* after
  retries, the window is too wide for the page size — lower `--top` or narrow
  `--start`/`--end`.
- **`ErrorInternalServerError` with categories** — you passed a server-side filter
  that OR's multiple categories. Fetch unfiltered and filter with `jq` instead.
- **Re-authenticate from scratch** — delete `~/.graph_token_cache.json`.

## Alternatives considered (and why they're not used)

- **M365 MCP connector** — has no category filter, which is the whole reason this
  script exists. Fine for general calendar reads, not for category-scoped queries.
- **Azure CLI app ID** — no `Calendars.Read` consent in this tenant
  (`ErrorAccessDenied`).
- **Graph Explorer client ID** — SPA-only; rejects the `http://localhost` redirect
  (`AADSTS900971`).

Stick with the Graph PowerShell client ID — it's the one known to work here.
