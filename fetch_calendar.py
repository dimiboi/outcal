# /// script
# requires-python = ">=3.10"
# dependencies = ["msal", "httpx"]
# ///
import argparse
import asyncio
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlencode

import httpx
import msal

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
TENANT = "2dfb2f0b-4d21-4268-9559-72926144c918"  # BCG
SCOPES = ["Calendars.Read"]
CACHE = Path.home() / ".graph_token_cache.json"
GRAPH = "https://graph.microsoft.com/v1.0"
OUT = Path("data") / "graph.jsonl"


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
    if not result:
        result = app.acquire_token_interactive(scopes=SCOPES)

    if cache.has_state_changed:
        CACHE.write_text(cache.serialize())

    return result["access_token"]


async def fetch_all(token: str, start_url: str) -> None:
    OUT.parent.mkdir(exist_ok=True)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    url: str | None = start_url
    pages = 0
    events = 0
    async with httpx.AsyncClient(timeout=60) as client:
        with OUT.open("w") as f:
            while url:
                resp = await client.get(url, headers=headers)
                resp.raise_for_status()
                payload = resp.json()
                page_events = payload.get("value", [])
                for event in page_events:
                    f.write(json.dumps(event) + "\n")
                events += len(page_events)
                pages += 1
                print(f"page {pages}: {len(page_events)} items")
                url = payload.get("@odata.nextLink")

    print(f"wrote {events} event(s) across {pages} page(s) to {OUT}")


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


def parse_args() -> argparse.Namespace:
    today = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    default_start = (today - timedelta(days=365)).isoformat().replace("+00:00", "Z")
    default_end = (today + timedelta(days=180)).isoformat().replace("+00:00", "Z")

    p = argparse.ArgumentParser(description="Fetch /me/calendarView events into data/graph.jsonl")
    p.add_argument("--start", default=default_start, help="ISO 8601 startDateTime (UTC)")
    p.add_argument("--end", default=default_end, help="ISO 8601 endDateTime (UTC)")
    p.add_argument("--top", type=int, default=100, help="Page size (Graph max 999, but calendarView 504s on large pages over wide windows)")
    p.add_argument("--category", default=None, help="Filter to events in this Outlook category (server-side), e.g. 'Travel'")
    return p.parse_args()


async def main() -> None:
    args = parse_args()
    token = get_token()
    url = build_url(args.start, args.end, args.top, category_filter(args.category))
    print(f"GET {url}")
    await fetch_all(token, url)


if __name__ == "__main__":
    asyncio.run(main())
