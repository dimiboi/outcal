# Microsoft Graph calendar CLI

`fetch_calendar.py` fetches your Outlook calendar from Microsoft Graph into
`data/graph.jsonl`. The whole tool is one file.

**README.md is the source of truth** — read it before changing auth or query
behavior. It covers setup, the registered-Mac / client-ID auth requirements,
flags, output format, troubleshooting, and which alternatives were rejected.

## Notes for Claude
- Always go through `uv run`. After editing, syntax-check with
  `uv run python -m py_compile fetch_calendar.py` (the project has no test suite).
- For category-scoped calendar queries, use this script — not the M365 MCP
  calendar tools loaded in this session, which can't filter by category.
