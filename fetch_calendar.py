# /// script
# requires-python = ">=3.10"
# dependencies = ["msal", "httpx", "tenacity"]
# ///
import argparse
import asyncio
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote, urlencode

import httpx
import msal
from tenacity import (
    AsyncRetrying,
    RetryCallState,
    retry_if_exception,
    stop_after_attempt,
    wait_exponential,
)

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
TENANT = "2dfb2f0b-4d21-4268-9559-72926144c918"  # BCG
SCOPES = ["Calendars.Read"]
CACHE = Path.home() / ".graph_token_cache.json"
GRAPH = "https://graph.microsoft.com/v1.0"
OUT = Path("data") / "graph.jsonl"
RETRYABLE_STATUS = {429, 503, 504}  # rate limit, service unavailable, gateway timeout


def get_token() -> str:
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
    if not result or "access_token" not in result:  # cache miss or failed refresh
        result = app.acquire_token_interactive(scopes=SCOPES)

    if cache.has_state_changed:
        CACHE.write_text(cache.serialize())

    if "access_token" not in result:
        raise SystemExit(f"auth failed: {result.get('error_description') or result.get('error') or result}")

    return result["access_token"]


def _should_retry(exc: BaseException) -> bool:
    """Retry transient failures only: rate limits, gateway/service errors, and network blips."""
    if isinstance(exc, httpx.HTTPStatusError):
        return exc.response.status_code in RETRYABLE_STATUS
    return isinstance(exc, httpx.TransportError)  # timeouts, connection resets, etc.


MAX_BACKOFF = 60  # seconds; ceiling for both exponential backoff and Retry-After
_BACKOFF = wait_exponential(multiplier=1, min=1, max=MAX_BACKOFF)


def _wait(retry_state: RetryCallState) -> float:
    """Honor a numeric Retry-After header (Graph sends it on 429/503); else exponential backoff. Both capped at MAX_BACKOFF."""
    exc = retry_state.outcome.exception() if retry_state.outcome else None
    if isinstance(exc, httpx.HTTPStatusError):
        retry_after = exc.response.headers.get("Retry-After", "")
        if retry_after.isdigit():  # numeric seconds; ignore the rarer HTTP-date form
            return min(float(retry_after), MAX_BACKOFF)
    return _BACKOFF(retry_state)


def _log_retry(retry_state: RetryCallState) -> None:
    exc = retry_state.outcome.exception() if retry_state.outcome else None
    reason = f"HTTP {exc.response.status_code}" if isinstance(exc, httpx.HTTPStatusError) else type(exc).__name__
    sleep = retry_state.next_action.sleep if retry_state.next_action else 0.0
    # stderr, not stdout: --id mode pipes stdout to jq, and retry logs are diagnostics in both modes
    print(f"  transient failure ({reason}); retrying in {sleep:.1f}s (attempt {retry_state.attempt_number} failed)", file=sys.stderr)


async def fetch_page(client: httpx.AsyncClient, url: str, max_retries: int) -> dict:
    """GET one page, retrying transient errors with backoff. Non-transient errors propagate at once."""
    async for attempt in AsyncRetrying(
        stop=stop_after_attempt(max_retries + 1),
        wait=_wait,
        retry=retry_if_exception(_should_retry),
        before_sleep=_log_retry,
        reraise=True,  # surface the real httpx error, not tenacity's RetryError
    ):
        with attempt:
            resp = await client.get(url)
            resp.raise_for_status()
            return resp.json()
    raise RuntimeError("unreachable: AsyncRetrying exits via return or raise")


async def fetch_all(token: str, start_url: str, max_retries: int) -> None:
    OUT.parent.mkdir(exist_ok=True)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    tmp = OUT.with_name(OUT.name + ".tmp")  # atomic swap: don't clobber good data on failure

    url: str | None = start_url
    pages = 0
    events = 0
    try:
        async with httpx.AsyncClient(timeout=60, headers=headers) as client:
            with tmp.open("w") as f:
                while url:
                    payload = await fetch_page(client, url, max_retries)
                    page_events = payload.get("value", [])
                    for event in page_events:
                        f.write(json.dumps(event) + "\n")
                    events += len(page_events)
                    pages += 1
                    print(f"page {pages}: {len(page_events)} items")
                    url = payload.get("@odata.nextLink")
        tmp.replace(OUT)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise

    print(f"wrote {events} event(s) across {pages} page(s) to {OUT}")


def _graph_error_code(exc: httpx.HTTPStatusError) -> str:
    """Pull Graph's machine-readable error code (e.g. ErrorItemNotFound) out of the body, if any."""
    try:
        return exc.response.json().get("error", {}).get("code", "")
    except (ValueError, AttributeError):  # non-JSON or unexpected shape
        return ""


async def fetch_by_ids(token: str, ids: list[str], max_retries: int) -> None:
    """Look up specific events via /me/events/{id} and print each as one JSON line to stdout.

    Unlike the calendarView dump this is a non-destructive lookup: it never touches
    data/graph.jsonl, so a single-event fetch can't clobber a full calendar pull.
    Diagnostics go to stderr so stdout stays clean JSONL for piping into jq.
    """
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    failures = 0
    async with httpx.AsyncClient(timeout=60, headers=headers) as client:
        for event_id in ids:
            if "/" in event_id:
                # Graph routes a decoded slash as a path separator in /me/events/{id}; neither %2F
                # nor double-encoding gets through (per MS guidance), so a '/' id can't be fetched
                # this way. Such ids are usually recurrence exceptions or externally-synced events,
                # whose full object is already in the calendarView dump. Fail fast with a real reason
                # rather than a misleading ErrorItemNotFound.
                failures += 1
                print(f"  skipped {event_id}: id contains '/', which Graph can't route in /me/events/{{id}} — read this event from the calendarView dump instead", file=sys.stderr)
                continue
            url = f"{GRAPH}/me/events/{quote(event_id, safe='')}"  # + and = round-trip fine through single-encoding; / is rejected above
            print(f"GET {url}", file=sys.stderr)
            try:
                event = await fetch_page(client, url, max_retries)
            except httpx.HTTPStatusError as exc:
                failures += 1
                code = _graph_error_code(exc) or "error"
                print(f"  failed {event_id}: HTTP {exc.response.status_code} {code}", file=sys.stderr)
                continue
            print(json.dumps(event))

    fetched = len(ids) - failures
    print(f"fetched {fetched} of {len(ids)} id(s)", file=sys.stderr)
    if failures:
        raise SystemExit(1)


def category_filter(category: str | None) -> str | None:
    if not category:
        return None
    safe = category.replace("'", "''")  # OData escapes a single quote by doubling
    return f"categories/any(c:c eq '{safe}')"


def build_url(start: str, end: str, top: int, odata_filter: str | None = None) -> str:
    params = {"startDateTime": start, "endDateTime": end, "$top": top}
    if odata_filter:
        params["$filter"] = odata_filter
    return f"{GRAPH}/me/calendarView?{urlencode(params)}"


def _non_negative_int(value: str) -> int:
    n = int(value)
    if n < 0:
        raise argparse.ArgumentTypeError("must be >= 0")
    return n


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Fetch /me/calendarView events into data/graph.jsonl")
    p.add_argument("--id", dest="ids", nargs="+", metavar="ID", help="Look up specific events by id via /me/events/{id} and print them to stdout (does not write data/graph.jsonl)")
    # The calendarView-query flags default to None, not their real values, so we can tell
    # "passed" from "omitted" and reject them in --id mode. Real defaults are filled in below.
    p.add_argument("--start", help="ISO 8601 startDateTime (UTC); default: 365 days ago")
    p.add_argument("--end", help="ISO 8601 endDateTime (UTC); default: 180 days ahead")
    p.add_argument("--top", type=int, help="Page size, default 100 (Graph max 999, but calendarView 504s on large pages over wide windows)")
    p.add_argument("--category", help="Filter to events in this Outlook category (server-side), e.g. 'Travel'")
    p.add_argument("--max-retries", type=_non_negative_int, default=5, help="Retry a page up to N times on transient errors (429/503/504, timeouts); 0 disables")
    args = p.parse_args()

    if args.ids:
        # --id is a standalone lookup; the calendarView flags don't apply, so reject them
        # outright rather than silently ignore. --max-retries works in both modes.
        rejected = {"--start": args.start, "--end": args.end, "--top": args.top, "--category": args.category}
        passed = [flag for flag, value in rejected.items() if value is not None]
        if passed:
            p.error(f"these flags apply to the calendarView query, not --id lookups: {', '.join(passed)}")
    else:
        today = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        if args.start is None:
            args.start = (today - timedelta(days=365)).isoformat().replace("+00:00", "Z")
        if args.end is None:
            args.end = (today + timedelta(days=180)).isoformat().replace("+00:00", "Z")
        if args.top is None:
            args.top = 100

    return args


async def main() -> None:
    args = parse_args()
    token = get_token()
    if args.ids:
        await fetch_by_ids(token, args.ids, args.max_retries)
        return
    url = build_url(args.start, args.end, args.top, category_filter(args.category))
    print(f"GET {url}")
    await fetch_all(token, url, args.max_retries)


if __name__ == "__main__":
    asyncio.run(main())
