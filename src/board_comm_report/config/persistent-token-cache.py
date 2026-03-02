import os
import msal
from msal_extensions import (build_encrypted_persistence, 
                             PersistedTokenCache)

# 1. Setup Persistence (Encrypted on Mac/Windows)
cache_location = "token_cache.bin"
persistence = build_encrypted_persistence(cache_location)
cache = PersistedTokenCache(persistence)

client_id = os.environ.get("ONENOTE_CLIENT_ID")
tenant_id = os.environ.get("AZURE_TENANT_ID")
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/Notes.ReadWrite"]

app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

def get_token():
    # 2. Try to get token from cache silently
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result: return result['access_token']

    # 3. If no cache/expired, trigger Device Code (only needed once)
    print("MFA/Login required. Please follow the prompt:")
    flow = app.initiate_device_flow(scopes=scopes)
    print(flow['message'])
    result = app.acquire_token_by_device_flow(flow)
    print(f"Authenticated as: {result.get('id_token_claims').get('preferred_username')}")
    return result.get('access_token')

token = get_token()
print(f"Service is authorized! Token starts with: {token[:10]}...")