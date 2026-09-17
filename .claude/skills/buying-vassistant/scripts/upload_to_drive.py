"""Upload a finished buy-meeting deck to the team's Google Drive output folder.

Solves the one gap in the lineguide/merch-board workflow: the Drive MCP
connector can only take file content as a base64 *parameter*, which caps
uploads at a few tens of KB. A deck with form photos is far past that, so
this talks to the Drive REST API directly and streams the bytes.

Deliberate constraints (the output folders belong to the buyers, not to us):

- NEVER creates a folder. If the month folder for the meeting date doesn't
  exist yet, the deck goes to the year-level parent folder instead.
- NEVER deletes, trashes, or overwrites anything. There is no code path
  here that removes a file; an existing deck with the same name aborts the
  run unless --allow-duplicate is passed.
- Uploads are owned by the authenticating user, so decks land alongside the
  buyers' own files exactly as a manual upload would.

By default the deck is converted to native Google Slides on upload, matching
the convention in the output folders (the large majority of decks there are
Slides, not raw .pptx). Pass --keep-pptx to upload the file unconverted.

Credentials come from the environment, never from the repo or the command
line (argv is visible to other processes):

    GDRIVE_CLIENT_ID
    GDRIVE_CLIENT_SECRET
    GDRIVE_REFRESH_TOKEN

Mint them once with mint_drive_token.py (run that on your own machine, not
in a container). The token needs BOTH of these scopes:

    drive.file              to create the upload. On its own it is not
                            enough: it grants access only to files this
                            app created, so every folder lookup below
                            would 404.
    drive.metadata.readonly to resolve the month folder, read its
                            canAddChildren capability, and check for a
                            name clash -- all of which read metadata of
                            files this app did not create. Metadata only,
                            so no file contents are ever readable.

Use --dry-run to check argument handling, date/month resolution and the
target folder id without needing credentials or touching Drive.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

TOKEN_URL = "https://oauth2.googleapis.com/token"
API_ROOT = "https://www.googleapis.com/drive/v3"
UPLOAD_ROOT = "https://www.googleapis.com/upload/drive/v3"

FOLDER_MIME = "application/vnd.google-apps.folder"
SLIDES_MIME = "application/vnd.google-apps.presentation"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# The buy-meeting output parent (the "year" folder). The team may move to a
# new Drive link for 2027 and beyond -- pass --parent-folder-id to point at
# it rather than editing this default.
DEFAULT_PARENT_FOLDER_ID = "1Zaw3KSX2tGqIDwPMvAbQl7iZRUmk-9d0"

# The team's month folders are named inconsistently (SEPT, JULY and JUNE sit
# next to AUG, APR, MAR, FEB, JAN), so match any plausible spelling instead
# of assuming one abbreviation.
MONTH_ALIASES: Dict[int, Tuple[str, ...]] = {
    1: ("JAN", "JANUARY"),
    2: ("FEB", "FEBRUARY"),
    3: ("MAR", "MARCH"),
    4: ("APR", "APRIL"),
    5: ("MAY",),
    6: ("JUN", "JUNE"),
    7: ("JUL", "JULY"),
    8: ("AUG", "AUGUST"),
    9: ("SEP", "SEPT", "SEPTEMBER"),
    10: ("OCT", "OCTOBER"),
    11: ("NOV", "NOVEMBER"),
    12: ("DEC", "DECEMBER"),
}


# ---------------------------------------------------------------------------
# Meeting date / title
# ---------------------------------------------------------------------------

def parse_meeting_date(raw: str, year: Optional[int] = None) -> date:
    """Parse a meeting date written the way the team writes it: {M.DD}.

    Accepts 9.17, 09.17, 9.1 and 2026-09-17. Bare M.DD takes its year from
    --year, else the current year.
    """
    raw = raw.strip()
    iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", raw)
    if iso:
        return date(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))

    short = re.fullmatch(r"(\d{1,2})[.\-/](\d{1,2})", raw)
    if short:
        return date(year or date.today().year, int(short.group(1)), int(short.group(2)))

    raise ValueError(
        f"unrecognised meeting date {raw!r} -- use M.DD (e.g. 9.17) or YYYY-MM-DD"
    )


def meeting_date_from_filename(path: Path, year: Optional[int] = None) -> Optional[date]:
    """Pull the {M.DD} prefix off a generated deck filename, if it has one."""
    m = re.match(r"\s*(\d{1,2}\.\d{1,2})\b", path.name)
    if not m:
        return None
    try:
        return parse_meeting_date(m.group(1), year)
    except ValueError:
        return None


def drive_title(path: Path, keep_pptx: bool) -> str:
    """Drive title for the deck.

    The folder convention is a bare title with no extension (the Slides
    conversion drops it anyway); an unconverted .pptx keeps its filename.
    """
    return path.name if keep_pptx else path.stem


# ---------------------------------------------------------------------------
# Drive REST plumbing
# ---------------------------------------------------------------------------

class DriveError(RuntimeError):
    """A Drive API call failed; carries the response body for diagnosis."""


def _request(
    url: str,
    *,
    method: str = "GET",
    headers: Optional[Dict[str, str]] = None,
    body: Optional[bytes] = None,
) -> Tuple[int, Dict[str, str], bytes]:
    req = urllib.request.Request(url, data=body, method=method)
    for key, value in (headers or {}).items():
        req.add_header(key, value)
    try:
        with urllib.request.urlopen(req) as resp:
            return resp.status, dict(resp.headers), resp.read()
    except urllib.error.HTTPError as exc:  # surface the API's own message
        detail = exc.read().decode("utf-8", "replace").strip()
        raise DriveError(f"{method} {url} -> HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise DriveError(f"{method} {url} -> {exc.reason}") from exc


def access_token() -> str:
    """Exchange the stored refresh token for a short-lived access token."""
    missing = [
        name
        for name in ("GDRIVE_CLIENT_ID", "GDRIVE_CLIENT_SECRET", "GDRIVE_REFRESH_TOKEN")
        if not os.environ.get(name)
    ]
    if missing:
        raise DriveError(
            "missing credential environment variable(s): "
            + ", ".join(missing)
            + " -- set them in the Claude Code environment config, not in the repo"
        )

    payload = urllib.parse.urlencode(
        {
            "client_id": os.environ["GDRIVE_CLIENT_ID"],
            "client_secret": os.environ["GDRIVE_CLIENT_SECRET"],
            "refresh_token": os.environ["GDRIVE_REFRESH_TOKEN"],
            "grant_type": "refresh_token",
        }
    ).encode()

    _, _, raw = _request(
        TOKEN_URL,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=payload,
    )
    token = json.loads(raw).get("access_token")
    if not token:
        raise DriveError("token endpoint returned no access_token")
    return token


def _api_get(token: str, path: str, params: Dict[str, str]) -> Any:
    params = {**params, "supportsAllDrives": "true"}
    url = f"{API_ROOT}/{path}?{urllib.parse.urlencode(params)}"
    _, _, raw = _request(url, headers={"Authorization": f"Bearer {token}"})
    return json.loads(raw)


def _escape(value: str) -> str:
    """Escape a literal for a Drive query string."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def folder_info(token: str, folder_id: str) -> Dict[str, Any]:
    """Fetch a folder's name and whether we may add files to it."""
    return _api_get(
        token,
        f"files/{urllib.parse.quote(folder_id)}",
        {"fields": "id,name,mimeType,capabilities(canAddChildren)"},
    )


def list_subfolders(token: str, parent_id: str) -> List[Dict[str, Any]]:
    query = (
        f"'{_escape(parent_id)}' in parents"
        f" and mimeType = '{FOLDER_MIME}'"
        " and trashed = false"
    )
    folders: List[Dict[str, Any]] = []
    page_token: Optional[str] = None
    while True:
        params = {
            "q": query,
            "fields": "nextPageToken,files(id,name,capabilities(canAddChildren))",
            "pageSize": "100",
            "includeItemsFromAllDrives": "true",
        }
        if page_token:
            params["pageToken"] = page_token
        payload = _api_get(token, "files", params)
        folders.extend(payload.get("files", []))
        page_token = payload.get("nextPageToken")
        if not page_token:
            return folders


def find_month_folder(
    folders: List[Dict[str, Any]], month: int
) -> Optional[Dict[str, Any]]:
    """Match a month folder by any of the team's spellings, else None."""
    aliases = MONTH_ALIASES[month]
    for folder in folders:
        name = re.sub(r"\s+", " ", (folder.get("name") or "")).strip().upper()
        if name in aliases:
            return folder
    return None


def existing_file(token: str, folder_id: str, title: str) -> Optional[Dict[str, Any]]:
    """Look for a file already using this title in the target folder."""
    query = (
        f"name = '{_escape(title)}'"
        f" and '{_escape(folder_id)}' in parents"
        " and trashed = false"
    )
    payload = _api_get(
        token,
        "files",
        {
            "q": query,
            "fields": "files(id,name,mimeType,owners(emailAddress))",
            "pageSize": "10",
            "includeItemsFromAllDrives": "true",
        },
    )
    files = payload.get("files", [])
    return files[0] if files else None


def upload_file(
    token: str,
    path: Path,
    *,
    folder_id: str,
    title: str,
    convert_to_slides: bool,
) -> Dict[str, Any]:
    """Resumable upload of `path` into `folder_id`.

    Resumable rather than multipart because merch boards in these folders
    run to hundreds of MB.
    """
    metadata: Dict[str, Any] = {"name": title, "parents": [folder_id]}
    if convert_to_slides:
        # Setting the target mimeType makes Drive convert on ingest.
        metadata["mimeType"] = SLIDES_MIME

    source_mime = mimetypes.guess_type(path.name)[0] or PPTX_MIME
    size = path.stat().st_size

    status, headers, _ = _request(
        f"{UPLOAD_ROOT}/files?{urllib.parse.urlencode({'uploadType': 'resumable', 'supportsAllDrives': 'true'})}",
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=UTF-8",
            "X-Upload-Content-Type": source_mime,
            "X-Upload-Content-Length": str(size),
        },
        body=json.dumps(metadata).encode(),
    )
    session_uri = headers.get("Location")
    if not session_uri:
        raise DriveError(f"no resumable session URI returned (HTTP {status})")

    _, _, raw = _request(
        session_uri,
        method="PUT",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": source_mime,
            "Content-Length": str(size),
        },
        body=path.read_bytes(),
    )
    return json.loads(raw)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Upload a buy-meeting deck to the team's Drive output folder "
        "(additive only: never creates folders, never deletes or overwrites).",
    )
    ap.add_argument("deck", type=Path, help="path to the .pptx to upload")
    ap.add_argument(
        "--meeting-date",
        default=None,
        help="meeting date as M.DD (e.g. 9.17) or YYYY-MM-DD; "
        "defaults to the {M.DD} prefix of the deck filename, else today",
    )
    ap.add_argument(
        "--year",
        type=int,
        default=None,
        help="year for a bare M.DD meeting date (default: current year)",
    )
    ap.add_argument(
        "--parent-folder-id",
        default=DEFAULT_PARENT_FOLDER_ID,
        help="the year-level output folder holding the month subfolders; "
        "pass the new id when the team moves to a new Drive link",
    )
    ap.add_argument("--title", default=None, help="override the Drive title")
    ap.add_argument(
        "--keep-pptx",
        action="store_true",
        help="upload as .pptx instead of converting to Google Slides",
    )
    ap.add_argument(
        "--allow-duplicate",
        action="store_true",
        help="upload even if a file with this title is already in the target "
        "folder (we cannot delete the old one, so this leaves both)",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="resolve the date, target folder and title, then stop without "
        "uploading (skips Drive entirely if credentials are absent)",
    )
    args = ap.parse_args(argv)

    deck: Path = args.deck
    if not deck.is_file():
        print(f"error: no such file: {deck}", file=sys.stderr)
        return 2

    if args.meeting_date:
        try:
            meeting = parse_meeting_date(args.meeting_date, args.year)
        except ValueError as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 2
    else:
        meeting = meeting_date_from_filename(deck, args.year) or date.today()

    title = args.title or drive_title(deck, args.keep_pptx)
    size_mb = deck.stat().st_size / (1024 * 1024)

    print(f"deck        : {deck}  ({size_mb:.1f} MB)")
    print(f"meeting date: {meeting.isoformat()}")
    print(f"title       : {title}")
    print(f"format      : {'.pptx (unconverted)' if args.keep_pptx else 'Google Slides'}")
    print(f"parent      : {args.parent_folder_id}")

    have_creds = all(
        os.environ.get(name)
        for name in ("GDRIVE_CLIENT_ID", "GDRIVE_CLIENT_SECRET", "GDRIVE_REFRESH_TOKEN")
    )
    if args.dry_run and not have_creds:
        print(
            "\ndry run: no credentials in the environment, so the target folder "
            "was not resolved.\nmonth folder would be matched against: "
            + ", ".join(MONTH_ALIASES[meeting.month])
        )
        return 0

    try:
        token = access_token()

        parent = folder_info(token, args.parent_folder_id)
        subfolders = list_subfolders(token, args.parent_folder_id)
        month_folder = find_month_folder(subfolders, meeting.month)

        if month_folder:
            target_id = month_folder["id"]
            target_name = month_folder.get("name", "?")
            print(f"target      : {target_name} (month folder)")
        else:
            # Never create a folder in the buyers' space -- fall back to the
            # year-level parent and say so.
            target_id = args.parent_folder_id
            target_name = parent.get("name", "?")
            print(
                f"target      : {target_name} (year folder -- no "
                f"{MONTH_ALIASES[meeting.month][0]} folder exists; not creating one)"
            )

        target = month_folder or parent
        if not target.get("capabilities", {}).get("canAddChildren", False):
            print(
                f"\nerror: no permission to add files to {target_name!r}. "
                "Ask the folder's owner for edit access.",
                file=sys.stderr,
            )
            return 1

        clash = existing_file(token, target_id, title)
        if clash and not args.allow_duplicate:
            owner = (clash.get("owners") or [{}])[0].get("emailAddress", "unknown")
            print(
                f"\nerror: {title!r} is already in {target_name!r} (owner: {owner}).\n"
                "Nothing was uploaded. This script never deletes or overwrites, so "
                "either rename with --title, or pass --allow-duplicate to add a "
                "second copy and tidy up in Drive yourself.",
                file=sys.stderr,
            )
            return 1

        if args.dry_run:
            print(f"\ndry run: would upload to {target_name!r} ({target_id}). Stopping.")
            return 0

        result = upload_file(
            token,
            deck,
            folder_id=target_id,
            title=title,
            convert_to_slides=not args.keep_pptx,
        )
    except DriveError as exc:
        print(f"\nerror: {exc}", file=sys.stderr)
        return 1

    file_id = result.get("id", "?")
    print(f"\nuploaded    : {title}")
    print(f"file id     : {file_id}")
    if not args.keep_pptx:
        print(f"open        : https://docs.google.com/presentation/d/{file_id}/edit")
    else:
        print(f"open        : https://drive.google.com/file/d/{file_id}/view")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
