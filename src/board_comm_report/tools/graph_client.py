import os
import msal
import time
from typing import Optional
from azure.core.credentials import AccessToken
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient


class MsalCachedTokenCredential:
    """Custom Azure identity credential using MSAL persistent cache."""
    def __init__(self, client_id: str):
        self.client_id = client_id

    def get_token(self, *scopes, **kwargs) -> AccessToken:
        # 1. Setup the persistent cache
        cache = msal.SerializableTokenCache()
        if os.path.exists("token_cache.bin"):
            with open("token_cache.bin", "r") as f:
                cache.deserialize(f.read())

        app = msal.PublicClientApplication(
            self.client_id,
            authority="https://login.microsoftonline.com/common",
            token_cache=cache
        )

        # 2. This call is 100% silent and requires no user input
        accounts = app.get_accounts()
        if accounts:
            msal_scopes = list(scopes) if scopes else ["Notes.Read.All"]
            result = app.acquire_token_silent(msal_scopes, account=accounts[0])
            
            if result and 'access_token' in result:
                expires_in = result.get('expires_in', 3600)
                # Success! Use result['access_token'] for your Graph calls
                print("Token acquired silently.")
                return AccessToken(result["access_token"], int(time.time()) + expires_in)
            else:
                raise Exception(f"Silent auth failed. Please re-authenticate manually. Result: {result}")
        else:
            # This only happens if your cache is deleted or password changed
            print("Manual re-auth required once to re-prime the cache.")
            raise Exception("Manual re-auth required once to re-prime the cache. No accounts found in token_cache.bin.")

# Module-level cache for the Microsoft Graph client to avoid redundant authentication
# and to prevent Pydantic serialization issues within CrewAI tools.
_cached_sharepoint_client: Optional[GraphServiceClient] = None
_cached_onenote_client: Optional[GraphServiceClient] = None

def get_graph_client(client_type: str = "sharepoint") -> Optional[GraphServiceClient]:
    global _cached_sharepoint_client, _cached_onenote_client
    
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    
    if client_type == "sharepoint":
        if _cached_sharepoint_client is None:
            client_id = os.environ.get("SHAREPOINT_CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
            client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET") or os.environ.get("AZURE_CLIENT_SECRET")
            if tenant_id and client_id and client_secret:
                credential = ClientSecretCredential(tenant_id, client_id, client_secret)
                _cached_sharepoint_client = GraphServiceClient(
                    credentials=credential, 
                    scopes=['https://graph.microsoft.com/.default']
                )
        return _cached_sharepoint_client
        
    elif client_type == "onenote":
        if _cached_onenote_client is None:
            client_id = os.environ.get("ONENOTE_CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
            if client_id:
                credential = MsalCachedTokenCredential(client_id)
                _cached_onenote_client = GraphServiceClient(
                    credentials=credential, 
                    scopes=['Notes.Read.All']
                )
        return _cached_onenote_client
        
    return None
