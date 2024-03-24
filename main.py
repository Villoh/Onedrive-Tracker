# The idea is to build a change tracker for Sharepoint/OneDrive folders

# from office365.onedrive.admin import
import atexit
import os
import tempfile
import msal
from office365.graph_client import GraphClient
from config import CLIENT_ID
from onedrive import download_files

# Define the cache file path
MSAL_CACHE_FILE = ".msal_cache.bin"
AUTHORITY_URL = f'https://login.microsoftonline.com/consumers'
SCOPES = ["https://graph.microsoft.com/.default"]

def acquire_token_interactive() -> dict:
    """Generate Microsoft Token interactively. It will launch a web page to login. 
    If you've alredy logged in you won't need to login again.
    Returns:
        dict: Result token.
    """
    # Manage cache storage
    cache=msal.SerializableTokenCache()
    if os.path.exists(MSAL_CACHE_FILE):
        cache.deserialize(open(MSAL_CACHE_FILE, "r").read())

    # Register lambda to save file upon program termination
    atexit.register(lambda:
        open(MSAL_CACHE_FILE, "w").write(cache.serialize())
        # Hint: The following optional line persists only when state changed
        if cache.has_state_changed else None
        )
    
    app = msal.PublicClientApplication(
        authority=AUTHORITY_URL,
        client_id=CLIENT_ID,
        token_cache=cache
    )
    accounts = app.get_accounts()
    if accounts:
        # If so, you could then somehow display these accounts and let end user choose
        # Assuming the end user chose this one
        chosen = accounts[0]
        # Now let's try to find a token in cache for this account
        result = app.acquire_token_silent(SCOPES, account=chosen)
        if not result:
            # If silent acquisition fails, prompt the user to sign in interactively
            result = app.acquire_token_interactive(scopes=SCOPES, account=accounts[0], domain_hint="consumers")
    else:
        result = app.acquire_token_interactive(  # It automatically provides PKCE protection
                scopes=SCOPES # This refers to the permissions setted in the registered app.
            )
    
    # Cache the account for future use
    if "access_token" in result and 'refreh_token' in result:
        cache.add({
            "client_id": CLIENT_ID,
            "home_account_id": result.get("home_account_id"),
            "environment": AUTHORITY_URL,
            "realm": result.get("realm"),
            "local_account_id": result.get("local_account_id"),
            "username": result.get("id_token_claims").get("preferred_username"),
            "authority_type": "MSSTS",
            "access_token": result.get("access_token"),
            "id_token": result.get("id_token"),
            "refresh_token": result.get("refresh_token"),
            "expires_in": result.get("expires_in"),
            "extended_expires_in": result.get("extended_expires_in"),
            "token_type": "Bearer"
        })
    return result

client = GraphClient(acquire_token_interactive)

drive = client.me.drive.get().execute_query()

# download files from OneDrive into local folder 
download_files(drive.root, './tmp_files/')


