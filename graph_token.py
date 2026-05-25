# /// script
# requires-python = ">=3.10"
# dependencies = ["msal"]
# ///
from pathlib import Path

import msal

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Graph Explorer
TENANT = "2dfb2f0b-4d21-4268-9559-72926144c918"  # BCG
SCOPES = ["Calendars.Read"]
CACHE = Path.home() / ".graph_token_cache.json"

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
