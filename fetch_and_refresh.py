#!/usr/bin/env python3
"""
Pull the current Fangst 2026 v1.xlsx from OneDrive via Microsoft Graph
using a cached MSAL refresh token, save it locally, and print the path.

Requires env vars:
  AZURE_CLIENT_ID   - the Azure app registration's Application (client) ID
  MSAL_TOKEN_CACHE  - a serialized MSAL token cache (from the one-time
                       device-code login) containing a refreshable account

On success, if the token cache changed (refresh token rotated), the new
serialized cache is written to msal_cache_out.txt so the workflow can
update the GitHub secret and keep itself working indefinitely.
"""
import os
import sys
import urllib.parse

import msal
import requests

CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AUTHORITY = "https://login.microsoftonline.com/common"  # works for personal + work/school accounts
SCOPES = ["Files.Read"]
GRAPH_PATH = "/me/drive/root:/Documents/07 - Privat/Fangstdagbok/Fangst 2026 v1.xlsx:/content"
OUT_FILE = "Fangst_2026_v1.xlsx"


def get_access_token():
    cache = msal.SerializableTokenCache()
    cache.deserialize(os.environ["MSAL_TOKEN_CACHE"])
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    accounts = app.get_accounts()
    if not accounts:
        sys.exit("No cached account in MSAL_TOKEN_CACHE -- re-run the one-time device-code login.")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        sys.exit("Silent token refresh failed -- the refresh token may have expired; "
                  "re-run the one-time device-code login and update the MSAL_TOKEN_CACHE secret.")

    if cache.has_state_changed:
        with open("msal_cache_out.txt", "w") as f:
            f.write(cache.serialize())

    return result["access_token"]


def download_workbook(access_token, out_path):
    url = "https://graph.microsoft.com/v1.0" + urllib.parse.quote(GRAPH_PATH, safe="/:")
    resp = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})
    resp.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(resp.content)


if __name__ == "__main__":
    token = get_access_token()
    download_workbook(token, OUT_FILE)
    print(OUT_FILE)
