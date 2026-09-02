import os
import sys

import msal

CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.Read"]

cache = msal.SerializableTokenCache()
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    print("::error::Failed to create device flow: %s" % flow)
    sys.exit(1)

# Safe to print -- this is a short-lived public sign-in code, not a secret.
print(flow["message"])
sys.stdout.flush()

result = app.acquire_token_by_device_flow(flow)  # blocks until completed or expired (~15 min)

if "access_token" not in result:
    print("::error::Device flow failed: %s" % result.get("error_description", result))
    sys.exit(1)

with open("msal_cache_secret.txt", "w") as f:
    f.write(cache.serialize())

print("Login succeeded. Token cache will be stored as a secret (not printed).")
