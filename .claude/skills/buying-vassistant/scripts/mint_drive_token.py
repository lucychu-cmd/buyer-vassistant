"""One-time OAuth consent helper: mint a Drive refresh token for uploads.

Run this ONCE, ON YOUR OWN MACHINE -- never in a Claude Code container or
any other shared/remote box. It opens a browser for Google's consent
screen, catches the redirect on localhost, and prints a refresh token that
upload_to_drive.py then uses forever.

Why it can't run in the container: the OAuth redirect comes back to
http://localhost, which only works on the machine running your browser.
A remote container also has no business holding a long-lived credential
to your Drive.

Scopes requested (deliberately narrow -- see --scopes to override):

    drive.file              create/manage only the files this app uploads.
                            It cannot touch anything it didn't create.
    drive.metadata.readonly read file/folder NAMES and permissions, no
                            file contents. Needed so the uploader can find
                            the month folder and check for a name clash.

Together these let the uploader do its job without ever being able to read
the contents of anything in your Drive.

Setup in Google Cloud Console, before running this:

    1. Pick or create a project. If your Workspace admin gives you a
       project inside the REVOLVE organisation, use it -- see step 2.
    2. APIs & Services -> OAuth consent screen. Choose User Type
       "Internal" if it is offered. Internal matters: an "External" app
       left in "Testing" mode has its refresh tokens REVOKED AFTER 7
       DAYS, which would mean re-running this every week. Internal apps
       have no such expiry and need no Google verification.
    3. APIs & Services -> Library -> enable "Google Drive API".
    4. APIs & Services -> Credentials -> Create Credentials -> OAuth
       client ID -> Application type "Desktop app". Desktop clients
       accept any localhost port, so there is no redirect URI to
       register.
    5. Copy the client ID and client secret into the environment (below).

Then:

    export GDRIVE_CLIENT_ID='...apps.googleusercontent.com'
    export GDRIVE_CLIENT_SECRET='...'        # or let this script prompt
    python3 mint_drive_token.py

The client secret is read from the environment or prompted for -- never
passed on the command line, where other processes can read it from argv.

The refresh token it prints is a long-lived credential for your Drive.
Put it in your Claude Code environment config (not in this repo, not in a
file, not in a chat message) and treat it like a password. If it ever
leaks, revoke it at https://myaccount.google.com/permissions.
"""

from __future__ import annotations

import argparse
import base64
import getpass
import hashlib
import http.server
import json
import os
import secrets
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from typing import Dict, List, Optional, Tuple

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
ABOUT_URL = "https://www.googleapis.com/drive/v3/about?fields=user(emailAddress,displayName)"

DEFAULT_SCOPES = (
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
)

SUCCESS_PAGE = b"""<!doctype html>
<html><head><meta charset="utf-8"><title>Authorised</title></head>
<body style="font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;line-height:1.5">
<h2>Authorised</h2>
<p>You can close this tab and return to the terminal.</p>
</body></html>"""

FAILURE_PAGE = b"""<!doctype html>
<html><head><meta charset="utf-8"><title>Failed</title></head>
<body style="font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;line-height:1.5">
<h2>Authorisation failed</h2>
<p>Check the terminal for details.</p>
</body></html>"""


# ---------------------------------------------------------------------------
# PKCE
# ---------------------------------------------------------------------------

def pkce_pair() -> Tuple[str, str]:
    """Return (code_verifier, code_challenge) for PKCE S256."""
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(64)).rstrip(b"=").decode()
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode()
    return verifier, challenge


# ---------------------------------------------------------------------------
# Loopback redirect catcher
# ---------------------------------------------------------------------------

class _Catcher(http.server.BaseHTTPRequestHandler):
    """Single-shot handler that records the query params from the redirect."""

    result: Dict[str, str] = {}

    def do_GET(self) -> None:  # noqa: N802 - stdlib naming
        query = urllib.parse.urlparse(self.path).query
        params = {k: v[0] for k, v in urllib.parse.parse_qs(query).items()}
        _Catcher.result = params

        ok = "code" in params and "error" not in params
        self.send_response(200 if ok else 400)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(SUCCESS_PAGE if ok else FAILURE_PAGE)

    def log_message(self, *_args) -> None:
        """Silence the default stderr access log."""


# ---------------------------------------------------------------------------
# Token exchange
# ---------------------------------------------------------------------------

def post_form(url: str, fields: Dict[str, str]) -> Dict[str, object]:
    body = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace").strip()
        raise RuntimeError(f"HTTP {exc.code} from {url}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"could not reach {url}: {exc.reason}") from exc


def whoami(access_token: str) -> Optional[str]:
    """Confirm which account the new token belongs to."""
    req = urllib.request.Request(
        ABOUT_URL, headers={"Authorization": f"Bearer {access_token}"}
    )
    try:
        with urllib.request.urlopen(req) as resp:
            user = json.loads(resp.read()).get("user", {})
            return user.get("emailAddress")
    except (urllib.error.HTTPError, urllib.error.URLError):
        return None


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Mint a Google Drive refresh token for upload_to_drive.py. "
        "Run once, on your own machine, not in a container.",
    )
    ap.add_argument(
        "--scopes",
        nargs="+",
        default=list(DEFAULT_SCOPES),
        help="OAuth scopes to request (default: drive.file + drive.metadata.readonly)",
    )
    ap.add_argument(
        "--port",
        type=int,
        default=0,
        help="localhost port for the redirect (default: an unused one)",
    )
    ap.add_argument(
        "--timeout",
        type=float,
        default=300.0,
        help="seconds to wait for the browser redirect (default: 300)",
    )
    ap.add_argument(
        "--no-browser",
        action="store_true",
        help="don't try to open a browser; just print the URL",
    )
    args = ap.parse_args(argv)

    client_id = os.environ.get("GDRIVE_CLIENT_ID", "").strip()
    if not client_id:
        print(
            "error: set GDRIVE_CLIENT_ID to your OAuth client ID first "
            "(see this file's docstring for how to create one).",
            file=sys.stderr,
        )
        return 2

    client_secret = os.environ.get("GDRIVE_CLIENT_SECRET", "").strip()
    if not client_secret:
        # Prompted rather than taken from argv, which other processes can read.
        client_secret = getpass.getpass("OAuth client secret (not echoed): ").strip()
    if not client_secret:
        print("error: no client secret given.", file=sys.stderr)
        return 2

    # Bind the loopback listener first, so the redirect_uri we send to Google
    # names the port we are actually listening on.
    try:
        server = http.server.HTTPServer(("127.0.0.1", args.port), _Catcher)
    except OSError as exc:
        print(f"error: cannot listen on port {args.port}: {exc}", file=sys.stderr)
        return 1
    server.timeout = args.timeout
    redirect_uri = f"http://localhost:{server.server_address[1]}/"

    verifier, challenge = pkce_pair()
    state = secrets.token_urlsafe(32)

    auth_url = AUTH_URL + "?" + urllib.parse.urlencode(
        {
            "client_id": client_id,
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(args.scopes),
            # offline + consent together guarantee a refresh token, even if
            # this account has already approved the app before.
            "access_type": "offline",
            "prompt": "consent",
            "include_granted_scopes": "true",
            "state": state,
            "code_challenge": challenge,
            "code_challenge_method": "S256",
        }
    )

    print("Scopes requested:")
    for scope in args.scopes:
        print(f"  - {scope}")
    print(f"\nListening on {redirect_uri}")
    print("\nOpen this URL and approve access:\n")
    print(auth_url)
    print()
    if not args.no_browser:
        try:
            webbrowser.open(auth_url)
        except Exception:  # headless box, no browser configured -- URL is above
            pass

    print(f"Waiting up to {args.timeout:.0f}s for the redirect...")
    _Catcher.result = {}
    thread = threading.Thread(target=server.handle_request, daemon=True)
    thread.start()
    thread.join(args.timeout)
    server.server_close()

    params = _Catcher.result
    if not params:
        print(
            "\nerror: no redirect received before the timeout. If the browser "
            "could not reach localhost, run this on a machine with a browser.",
            file=sys.stderr,
        )
        return 1
    if "error" in params:
        print(f"\nerror: Google returned {params['error']!r}", file=sys.stderr)
        return 1
    if not secrets.compare_digest(params.get("state", ""), state):
        print(
            "\nerror: state mismatch -- discarding this response. "
            "Start over; do not reuse the URL.",
            file=sys.stderr,
        )
        return 1

    try:
        tokens = post_form(
            TOKEN_URL,
            {
                "client_id": client_id,
                "client_secret": client_secret,
                "code": params["code"],
                "code_verifier": verifier,
                "grant_type": "authorization_code",
                "redirect_uri": redirect_uri,
            },
        )
    except RuntimeError as exc:
        print(f"\nerror: token exchange failed: {exc}", file=sys.stderr)
        return 1

    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        print(
            "\nerror: Google issued no refresh token. This usually means the "
            "consent screen was already approved without access_type=offline; "
            "revoke the app at https://myaccount.google.com/permissions and "
            "run this again.",
            file=sys.stderr,
        )
        return 1

    account = whoami(str(tokens.get("access_token", "")))
    granted = str(tokens.get("scope", "")).split()

    print("\n" + "=" * 68)
    print("Refresh token minted.")
    if account:
        print(f"Account: {account}  <- uploads will be owned by this user")
    if granted:
        print("Granted scopes:")
        for scope in granted:
            print(f"  - {scope}")
        missing = [s for s in args.scopes if s not in granted]
        if missing:
            print("\nWARNING: these requested scopes were NOT granted:")
            for scope in missing:
                print(f"  - {scope}")
            print("The uploader will fail until they are.")
    print("=" * 68)
    print(
        "\nPut these three in your Claude Code environment config -- NOT in the "
        "repo, a file, or a chat message:\n"
    )
    print(f"GDRIVE_CLIENT_ID={client_id}")
    print("GDRIVE_CLIENT_SECRET=<the secret you just entered>")
    print(f"GDRIVE_REFRESH_TOKEN={refresh_token}")
    print(
        "\nThis token does not expire on its own, but it is revoked if you "
        "change your Google password, revoke the app's access, or (for an "
        '"External" app still in Testing mode) after 7 days.'
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
