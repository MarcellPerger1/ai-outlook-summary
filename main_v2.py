import json
import sys

import msal
import requests

# -------------------------------------------------------------------------
# Configuration: Replace these placeholders with your actual Azure AD app info
# -------------------------------------------------------------------------

# The Application (client) ID of your Azure AD application.
CLIENT_ID = "486b33ea-7d53-47ef-a6d1-288f6fb606da"

TENTANT_ID = 'ce4606ac-f1df-4637-ab65-3d0e8bcb1706'
# Authority: For multi-tenant applications, use 'common'.
# For a specific tenant, use 'YOUR_TENANT_ID' or 'YOUR_TENANT_NAME.onmicrosoft.com'.
AUTHORITY = f"https://login.microsoftonline.com/{TENTANT_ID}"

# Scopes required for accessing Outlook emails through Microsoft Graph.
# Mail.ReadWrite allows reading, creating, and sending mail.
# Mail.Read allows read-only access.
# The default scope ('user.read') is implicitly added.
SCOPE = ["Mail.Read"]

# Redirect URI required for the Authorization Code Grant Flow (with PKCE).
# This URI MUST be registered in your Azure AD App Registration under
# 'Mobile and desktop applications' for this interactive flow to work.
# REDIRECT_URI = "https://tinyurl.com/CS-Bootcamp-Hwk-mp"


# -------------------------------------------------------------------------

def acquire_outlook_token_interactive(client_id: str, authority: str, scope: list):
    """
    Acquires an access token using the Authorization Code Grant Flow with PKCE.
    This method automatically opens the default web browser for sign-in,
    eliminating the need to type a code.
    """
    if client_id == "YOUR_CLIENT_ID_HERE":
        print(
            "ERROR: Please update the CLIENT_ID variable with your Azure AD Application ID.",
            file=sys.stderr)
        return None

    # 1. Initialize the Public Client Application
    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=msal.TokenCache()
    )

    # 2. Try to load the token from the cache first (silent login)
    accounts = app.get_accounts()
    result = None
    if accounts:
        print(f"Attempting silent login for account: {accounts[0]['username']}...")
        result = app.acquire_token_silent(scope, account=accounts[0])

    if result and "access_token" in result:
        print("Successfully acquired token from cache (silent login).")
        return result

    # 3. If silent login failed, initiate the interactive login flow
    print("\nNo valid token found in cache. Opening web browser for interactive sign-in...")

    try:
        # FIX: The previous error occurred because 'redirect_uri' was passed as
        # a keyword argument which was redundant/conflicting with MSAL's internal
        # handling for public client apps. We remove it here.
        result = app.acquire_token_interactive(
            scopes=scope,
            port=12345
        )
    except Exception as e:
        print(f"An error occurred during interactive token acquisition: {e}",
              file=sys.stderr)
        return None

    if result and "access_token" in result:
        print("\nAuthentication successful! Token acquired.")
        return result
    else:
        # Authentication failed or user closed the browser
        print("\nAuthentication failed or was canceled.", file=sys.stderr)
        error_description = result.get("error_description", "Unknown error")
        error_code = result.get("error")
        print(f"Error Code: {error_code}\nDescription: {error_description}", file=sys.stderr)
        return None


def read_latest_email(access_token: str):
    """
    Calls the Microsoft Graph API to fetch the latest email in the user's inbox.
    Requires the Mail.Read scope.
    """
    print("\n--- FETCHING LATEST EMAIL ---")

    # Base URL for Microsoft Graph API
    graph_url = "https://graph.microsoft.com/v1.0/me/messages"

    # Headers containing the access token
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Query parameters to fetch only the top 1 message, ordered by newest first
    params = {
        '$top': 1,
        '$orderby': 'receivedDateTime desc',
        '$select': 'subject,from,receivedDateTime,bodyPreview'  # Only fetch useful fields
    }

    try:
        response = requests.get(graph_url, headers=headers, params=params)
        response.raise_for_status()  # Raise an HTTPError for bad responses (4xx or 5xx)

        email_data = response.json()

        # Check if any messages were returned
        if not email_data.get('value'):
            print("Inbox is empty or no messages were found.")
            return

        latest_message = email_data['value'][0]

        print(f"Subject: {latest_message.get('subject', 'No Subject')}")
        print(
            f"From: {latest_message.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}")
        print(f"Received: {latest_message.get('receivedDateTime')}")
        print(f"Body Preview: {latest_message.get('bodyPreview', 'N/A')}")

    except requests.exceptions.HTTPError as e:
        print(
            f"ERROR: Failed to read email (HTTP {response.status_code}). Check your permissions (Mail.Read).",
            file=sys.stderr)
        try:
            # Try to print the detailed error from the response body
            error_details = response.json().get('error', {})
            print(f"Graph Error: {error_details.get('code')}", file=sys.stderr)
            print(f"Message: {error_details.get('message')}", file=sys.stderr)
        except json.JSONDecodeError:
            print("Could not decode error response.", file=sys.stderr)
    except requests.exceptions.RequestException as e:
        print(f"An error occurred during the API request: {e}", file=sys.stderr)


if __name__ == "__main__":

    # We no longer pass REDIRECT_URI to the main function as the MSAL interactive
    # method handles the local listener automatically.
    token_response = acquire_outlook_token_interactive(CLIENT_ID, AUTHORITY, SCOPE)

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
