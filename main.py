import msal
import sys
import time
import os

# -------------------------------------------------------------------------
# Configuration: Replace these placeholders with your actual Azure AD app info
# -------------------------------------------------------------------------

# The Application (client) ID of your Azure AD application.
# You must register an application in the Azure portal (App Registrations)
# and ensure it is configured as a 'Public client' or 'Mobile and desktop' app
# if using the Device Code Flow.
CLIENT_ID = "486b33ea-7d53-47ef-a6d1-288f6fb606da"  # e.g., "a1b2c3d4-e5f6-7890-a1b2-c3d4e5f67890"

TENTANT_ID = 'ce4606ac-f1df-4637-ab65-3d0e8bcb1706'
# Authority: For multi-tenant applications, use 'common'.
# For a specific tenant, use 'YOUR_TENANT_ID' or 'YOUR_TENANT_NAME.onmicrosoft.com'.
AUTHORITY = f"https://login.microsoftonline.com/{TENTANT_ID}"

# Scopes required for accessing Outlook emails through Microsoft Graph.
# Mail.ReadWrite allows reading, creating, and sending mail.
# Mail.Read allows read-only access.
# The default scope ('user.read') is implicitly added.
SCOPE = ["Mail.Read"]


# -------------------------------------------------------------------------

def acquire_outlook_token(client_id: str, authority: str, scope: list):
    """
    Acquires an access token for Microsoft Graph (Outlook) using the Device Code Flow.
    This method is suitable for console applications where user interaction is required
    to sign in through a browser on a separate device or computer.
    """
    if client_id == "YOUR_CLIENT_ID_HERE":
        print(
            "ERROR: Please update the CLIENT_ID variable with your Azure AD Application ID.",
            file=sys.stderr)
        print("See the comments in the file for setup instructions.", file=sys.stderr)
        return None

    # 1. Initialize the Public Client Application
    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        # Set a basic token cache file to persist the token, avoiding re-auth every time
        token_cache=msal.TokenCache()
    )

    # 2. Try to load the token from the cache first (silent login)
    accounts = app.get_accounts()
    result = None
    if accounts:
        # Assuming we only care about the first account found
        print(f"Attempting silent login for account: {accounts[0]['username']}...")
        result = app.acquire_token_silent(scope, account=accounts[0])

    if result:
        # Token found and renewed silently
        print("Successfully acquired token from cache (silent login).")
        return result

    # 3. If silent login failed, initiate the Device Code Flow
    print("\nNo token found in cache. Starting Device Code Flow for interactive login...")
    flow = app.initiate_device_flow(scopes=scope)

    if "user_code" not in flow:
        print("Could not initiate device flow:", file=sys.stderr)
        print(flow.get("error_description", "Unknown error"), file=sys.stderr)
        return None

    # 4. Display instructions to the user
    print("-" * 50)
    print(f"1. Go to this URL in a web browser: {flow['verification_uri']}")
    print(f"2. Enter the code: {flow['user_code']}")
    print("-" * 50)

    # 5. Wait for the user to authenticate (poll for token)
    result = app.acquire_token_by_device_flow(flow)

    if result and "access_token" in result:
        print("\nAuthentication successful!")
        return result
    else:
        print(result)
        # Authentication failed or user closed the browser
        print("\nAuthentication failed or was canceled.", file=sys.stderr)
        print(result.get("error_description", "Unknown error"), file=sys.stderr)
        return None


if __name__ == "__main__":

    token_response = acquire_outlook_token(CLIENT_ID, AUTHORITY, SCOPE)

    if token_response:
        access_token = token_response.get("access_token")
        expires_in = token_response.get("expires_in")
        token_type = token_response.get("token_type")

        print("\n--- TOKEN DETAILS ---")
        print(f"Token Type: {token_type}")
        print(f"Expires In (seconds): {expires_in}")
        print(f"Access Token (first 50 chars): {access_token[:50]}...")
        print(f"Scopes Granted: {token_response.get('scope').split()}")
        print(
            "\nUse this token in the 'Authorization: Bearer <token>' header to access Outlook mailboxes via Microsoft Graph.")
    else:
        print("\nFailed to acquire access token.")
