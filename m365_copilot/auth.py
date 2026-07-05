"""
M365 Copilot Authentication — MSAL PKCE token acquisition.

Uses the Microsoft first-party client ID to obtain a Sydney JWT token
that authenticates to the substrate.office.com SignalR backend.

Based on reverse engineering from cramt/m365-copilot-proxy.
"""

import json
import os
import time
from pathlib import Path
from typing import Optional

import msal

# Microsoft first-party client ID for Office web Copilot
CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1"
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"

SCOPES = [
    "https://substrate.office.com/sydney/M365Chat.Read",
    "https://substrate.office.com/sydney/sydney.readwrite",
]

CONFIG_DIR = Path.home() / ".config" / "m365-copilot"
CACHE_FILE = os.environ.get("M365_CACHE_FILE", str(CONFIG_DIR / "msal-cache.json"))


def _get_app() -> msal.PublicClientApplication:
    """Get or create the MSAL PublicClientApplication with cached state."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
    )

    # Load cached token cache if exists
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE) as f:
                app.token_cache.deserialize(f.read())
        except Exception:
            pass

    return app


def _save_cache(app: msal.PublicClientApplication):
    """Persist the MSAL token cache to disk."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(CACHE_FILE, "w") as f:
            f.write(app.token_cache.serialize())
    except Exception:
        pass


def decode_jwt(token: str) -> dict:
    """Decode a JWT payload without validation (extracts oid/tid for WS URL)."""
    import base64

    parts = token.split(".")
    if len(parts) != 3:
        raise ValueError("Invalid JWT format")

    payload = parts[1]
    # Add padding
    padding = 4 - (len(payload) % 4)
    if padding != 4:
        payload += "=" * padding

    try:
        decoded = base64.urlsafe_b64decode(payload)
    except Exception:
        decoded = base64.b64decode(payload)

    return json.loads(decoded)


def get_token_interactive() -> Optional[str]:
    """
    Acquire a Sydney token via interactive browser login.
    Opens a browser window for the user to sign in.
    """
    app = _get_app()

    # Try silent acquisition first (from cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(app)
            return result["access_token"]

    # Interactive login — opens a browser
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" in result:
        _save_cache(app)
        return result["access_token"]

    error = result.get("error_description", result.get("error", "Unknown error"))
    print(f"[!] Auth failed: {error}")
    return None


def get_token_device_code() -> Optional[str]:
    """
    Acquire a Sydney token via device code flow.
    NOTE: This app registration may reject device code flow (requires client_secret).
    """
    app = _get_app()

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(app)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"[!] Device flow failed: {flow.get('error_description', 'Unknown error')}")
        return None

    print(f"\n[>] To sign in, open:")
    print(f"    {flow['verification_uri']}")
    print(f"    And enter code: {flow['user_code']}")
    print(f"    (expires in {flow.get('expires_in', 900)} seconds)\n")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _save_cache(app)
        return result["access_token"]

    error = result.get("error_description", result.get("error", "Unknown error"))
    print(f"[!] Auth failed: {error}")
    return None


def get_token_manual() -> Optional[str]:
    """
    Manual PKCE auth flow: prints an auth URL, user visits it and pastes the
    redirect URL back into the terminal. No local server needed.
    """
    import base64
    import hashlib
    import secrets
    import urllib.parse

    app = _get_app()

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(app)
            return result["access_token"]

    # Generate PKCE codes manually
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()
    ).rstrip(b"=").decode()

    redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

    # Build auth URL manually
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "scope": " ".join(SCOPES),
        "response_mode": "query",
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    auth_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?{urllib.parse.urlencode(params)}"

    print(f"\n[>] Open this URL in your browser and log in:")
    print(f"    {auth_url}")
    print(f"\n[>] After logging in, you'll see a warning page saying the URL contains your password.")
    print(f"    This is EXPECTED — it's the auth code in the URL.")
    print(f"\n    ** Don't close the page! **")
    print(f"    Option A: Open DevTools (F12) → Network tab → Preserve log ✓")
    print(f"              Look for a request to 'oauth2/nativeclient?code=...'")
    print(f"              Right-click → Copy URL")
    print(f"    Option B: Copy the URL from the address bar BEFORE it navigates away")
    print(f"              (the 'Don't leave' / 'Stay on page' dialog lets you grab it)")
    print(f"\nPaste the redirect URL below:\n")

    import webbrowser
    try:
        webbrowser.open(auth_url)
    except Exception:
        pass

    redirect_result = input("Paste the redirect URL here:\n> ").strip()

    if not redirect_result:
        print("[!] No URL provided")
        return None

    # Extract auth code from redirect URL
    parsed = urllib.parse.urlparse(redirect_result)
    query_params = urllib.parse.parse_qs(parsed.query)
    auth_code = query_params.get("code", [None])[0]

    if not auth_code:
        fragment_params = urllib.parse.parse_qs(parsed.fragment)
        auth_code = fragment_params.get("code", [None])[0]

    if not auth_code:
        print(f"[!] Could not find 'code' parameter in the URL")
        print(f"    URL parsed: {parsed[:50]}")
        return None

    print(f"[>] Auth code found ({len(auth_code)} chars), exchanging for token...")

    # Exchange auth code for token directly via httpx (handles proxies/env properly)
    import sys
    import httpx
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    token_data = {
        "client_id": CLIENT_ID,
        "code": auth_code,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code_verifier": verifier,
        "scope": " ".join(SCOPES),
    }

    print(f"[>] POST {token_url}")
    sys.stdout.flush()

    try:
        resp = httpx.post(token_url, data=token_data, timeout=30)
        result = resp.json()
        print(f"[>] Response status: {resp.status_code}")
        sys.stdout.flush()
    except Exception as e:
        print(f"[!] HTTP error: {e}")
        return None

    if "access_token" in result:
        _save_cache(app)
        return result["access_token"]

    error = result.get("error_description", result.get("error", "Unknown error"))
    print(f"[!] Token exchange failed: {error}")
    return None


def get_token(force_refresh: bool = False, method: str = "interactive") -> Optional[str]:
    """
    Get a Sydney token. Tries silent refresh first, then falls back to auth flow.

    Args:
        force_refresh: If True, skip cache and force new login
        method: "interactive" (auto browser), "device_code", or "manual" (URL + local server)

    Returns:
        Access token string or None
    """
    app = _get_app()

    if not force_refresh:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])
            if result and "access_token" in result:
                _save_cache(app)
                return result["access_token"]

    if method == "device_code":
        return get_token_device_code()
    elif method == "manual":
        return get_token_manual()
    else:
        return get_token_interactive()


def get_token_info(token: str) -> dict:
    """Return decoded JWT claims for debugging."""
    claims = decode_jwt(token)
    return {
        "oid": claims.get("oid", "?"),
        "tid": claims.get("tid", "?"),
        "aud": claims.get("aud", "?"),
        "upn": claims.get("upn", claims.get("preferred_username", "?")),
        "name": claims.get("name", "?"),
        "expires_in": claims.get("exp", 0) - int(time.time()),
    }


def token_from_browser_js() -> Optional[str]:
    """
    Extract token from M365_TOKEN env var.
    
    In the browser (m365.cloud.microsoft), run in console:
        tokenProviders.sydney().then(t => console.log(t))
    
    Then set the env var and run.
    """
    token = os.environ.get("M365_TOKEN")
    if token:
        try:
            decode_jwt(token)
            return token
        except Exception:
            pass
    return None
