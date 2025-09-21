# rejections_core.py

import os
import re
import json
import time
import random
import mimetypes
import base64
from pathlib import Path
from datetime import datetime, timezone
from typing import List, Dict, Tuple, Callable, Any

from jinja2 import Template
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.utils import formatdate, make_msgid

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ----------------- DEFAULT SHEET SETTINGS (you can change in GUI) -----------------
DEFAULT_SPREADSHEET_ID = ""
DEFAULT_TAB_PREFERRED  = "Applicants"
DEFAULT_READ_RANGE     = "A:Z"
# -------------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ---------- Helpers ----------
def _with_backoff(fn: Callable[[], Any], *, retries: int = 5, base: float = 0.8, cap: float = 8.0):
    """Run `fn()` with exponential backoff on exceptions."""
    for i in range(retries):
        try:
            return fn()
        except Exception:
            if i == retries - 1:
                raise
            delay = min(cap, base * (2 ** i)) + random.uniform(0, 0.25)
            time.sleep(delay)

def _strip_html(s: str) -> str:
    # very basic fallback if only HTML provided
    return re.sub(r"<[^>]+>", "", s or "").strip()

# ---------- Auth & services ----------
def load_creds(credentials="credentials.json", token="token.json") -> Credentials:
    """
    Load credentials; if scopes missing or cannot refresh, trigger consent.

    Writes the OAuth token to the requested `token` path. If writing there raises
    PermissionError (e.g., read-only folder), it falls back to a user-writable
    path at `~/.rejections_gui/token.json` and raises a PermissionError with a
    detailed message so the caller (GUI) can surface where the token was saved.
    """
    def _write_token_json(path_str: str, creds_obj: Credentials) -> str:
        path = Path(path_str).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            path.write_text(creds_obj.to_json(), encoding="utf-8")
            try:
                os.chmod(path, 0o600)
            except Exception:
                pass
            return str(path)
        except PermissionError as e:
            # Fall back to a user-writable location
            fallback_dir = Path.home() / ".rejections_gui"
            fallback_dir.mkdir(parents=True, exist_ok=True)
            fb_path = fallback_dir / "token.json"
            fb_path.write_text(creds_obj.to_json(), encoding="utf-8")
            try:
                os.chmod(fb_path, 0o600)
            except Exception:
                pass
            # Raise a controlled error so the caller can show a helpful message
            raise PermissionError(
                f"Cannot write token to '{path}'. Wrote it to '{fb_path}' instead. "
                f"Update your Token JSON path in the GUI to this location."
            ) from e

    creds: Credentials | None = None
    token_path = Path(token).expanduser()
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    def needs_reconsent(c: Credentials | None) -> bool:
        if not c:
            return True
        if c.expired and not c.refresh_token:
            return True
        try:
            return not c.has_scopes(SCOPES)
        except Exception:
            return True

    # Try refresh if we have a refresh token
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    # New consent flow if needed
    if needs_reconsent(creds):
        flow = InstalledAppFlow.from_client_secrets_file(credentials, SCOPES)
        creds = flow.run_local_server(port=0)
        # Try writing where the user asked; may raise PermissionError with fallback info
        _write_token_json(token, creds)

    return creds  # type: ignore[return-value]

def gmail_service(creds):   return build("gmail", "v1", credentials=creds)
def sheets_service(creds):  return build("sheets", "v4", credentials=creds)

# ---------- Sheets helpers ----------
def quote_tab(name: str) -> str:
    """Quote sheet/tab names that contain spaces/special chars; escape single quotes."""
    specials = " []{}():;,'\"!@#$%^&*-+=/\\|?<>`~."
    if any(ch in name for ch in specials):
        return "'" + name.replace("'", "''") + "'"
    return name

def resolve_tab_title(ssvc, spreadsheet_id: str, preferred: str) -> str:
    """Prefer exact match, then case-insensitive, else first sheet title."""
    meta = _with_backoff(lambda: ssvc.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties.title"
    ).execute())
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if not titles:
        raise SystemExit("This spreadsheet has no tabs.")
    if preferred in titles:
        return preferred
    for t in titles:
        if t.lower() == preferred.lower():
            return t
    return titles[0]

def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def get_headers_index(values: List[List[str]]) -> Tuple[List[str], Dict[str,int]]:
    """Return original-case headers and lowercased index map."""
    if not values:
        raise SystemExit("Sheet is empty or range is incorrect.")
    headers = [str(h).strip() for h in values[0]]
    idx = {h.lower(): i for i, h in enumerate(headers)}
    return headers, idx

def to_records(values: List[List[str]]) -> Tuple[List[Dict[str,str]], List[str]]:
    headers, idx = get_headers_index(values)
    headers_lc = [h.lower() for h in headers]
    rows = []
    for r, row in enumerate(values[1:], start=2):  # row numbers in UI are 1-based
        rec = {h: (row[idx[h]].strip() if idx[h] < len(row) else "") for h in headers_lc}
        rec["_row_number"] = r
        rows.append(rec)
    return rows, headers  # records lowercased keys, headers original-case

def ensure_columns(headers: List[str], needed: List[str], ssvc, spreadsheet_id: str, tab_title: str) -> List[str]:
    """Ensure header row contains needed columns; update if missing and return new header list."""
    headers_lc = [h.lower() for h in headers]
    if all(c in headers_lc for c in needed):
        return headers
    new_headers = headers[:]
    for c in needed:
        if c not in [h.lower() for h in new_headers]:
            new_headers.append(c)  # append using lowercased new column name
    range_a1 = f"{quote_tab(tab_title)}!1:1"
    _with_backoff(lambda: ssvc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_a1,
        valueInputOption="RAW",
        body={"values": [new_headers]}
    ).execute())
    return new_headers

def write_status(ssvc, spreadsheet_id: str, tab_title: str, row_number: int,
                 status_col_index: int, time_col_index: int):
    iso = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    rng_status = f"{quote_tab(tab_title)}!{col_letter(status_col_index+1)}{row_number}"
    rng_time   = f"{quote_tab(tab_title)}!{col_letter(time_col_index+1)}{row_number}"
    _with_backoff(lambda: ssvc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "RAW",
            "data": [
                {"range": rng_status, "values": [["sent"]]},
                {"range": rng_time,   "values": [[iso]]},
            ],
        },
    ).execute())

# ---------- Gmail helpers ----------
def _attach(msg, filepath: str):
    ctype, _ = mimetypes.guess_type(filepath)
    if not ctype:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(filepath, "rb") as f:
        part = MIMEBase(maintype, subtype); part.set_payload(f.read())
    encoders.encode_base64(part)
    filename = Path(filepath).name
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    part.add_header("Content-Type", f'{ctype}; name="{filename}"')
    msg.attach(part)

def build_mime(sender, to, subject, text, html=None, cc=None, bcc=None, reply_to=None, attachments=None):
    msg = MIMEMultipart("mixed")
    alt = MIMEMultipart("alternative")

    # Fallbacks
    if not text and html:
        text = _strip_html(html)
    if not (text or html):
        text = "(no content)"

    alt.attach(MIMEText(text or "", "plain"))
    if html:
        alt.attach(MIMEText(html, "html"))
    msg.attach(alt)
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    if cc:
        msg["Cc"] = cc
    if bcc:
        msg["Bcc"] = bcc
    if reply_to:
        msg["Reply-To"] = reply_to
    for a in attachments or []:
        if Path(a).exists():
            _attach(msg, a)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}

def send_gmail(gsvc, body):
    return _with_backoff(lambda: gsvc.users().messages().send(userId="me", body=body).execute())

def render_template(path: str, ctx: dict) -> str | None:
    if not path:
        return None
    return Template(Path(path).read_text(encoding="utf-8")).render(**ctx)

# ---------- Worker ----------
def run_sender(config: dict, logq, stop_event):
    """Reads the sheet and (dry-)sends emails. Thread-safe via logq/stop_event."""
    def log(msg: str):
        logq.put(msg)

    try:
        # Auth
        log("Authorizing with Google… (browser window may open)")
        creds = load_creds(config["credentials"], config["token"])
        ssvc = sheets_service(creds)
        gsvc = None if config["dry_run"] else gmail_service(creds)

        # Resolve sheet & read
        actual_tab = resolve_tab_title(ssvc, config["spreadsheet_id"], config["tab"])
        qtab = quote_tab(actual_tab)
        read_range = f"{qtab}!{config['read_range']}"
        log(f"Reading {config['spreadsheet_id']} · tab '{actual_tab}' · range {config['read_range']}")
        resp = _with_backoff(lambda: ssvc.spreadsheets().values().get(
            spreadsheetId=config["spreadsheet_id"], range=read_range
        ).execute())
        values = resp.get("values", [])
        if not values:
            log(f"No rows found in tab '{actual_tab}' (range {config['read_range']}).")
            return

        records, headers_original_case = to_records(values)

        # Ensure logging columns
        headers_after = ensure_columns(headers_original_case, ["sent_status", "sent_at"], ssvc, config["spreadsheet_id"], actual_tab)
        hdr_index = {h.lower(): i for i, h in enumerate(headers_after)}  # robust lowercased index

        # Validate required columns
        required = {"email", "name", "role", "company"}
        missing = [c for c in required if c not in hdr_index]
        if missing:
            log(f"ERROR: Missing required columns in '{actual_tab}': {', '.join(missing)}")
            return

        # Build eligible list (respect send==yes if present, skip==yes, and skip already sent)
        def is_yes(x: str) -> bool: return (x or "").strip().lower() == "yes"

        eligible = []
        for rec in records:
            if is_yes(rec.get("skip", "")):
                continue
            if "send" in hdr_index and not is_yes(rec.get("send", "")):
                continue
            if (rec.get("sent_status","").strip().lower() == "sent"):
                continue
            if not (rec.get("email") and rec.get("name") and rec.get("role") and rec.get("company")):
                continue
            eligible.append(rec)

        total_eligible = len(eligible)
        if total_eligible == 0:
            log("No eligible rows to process.")
            return

        # Preview limit
        preview_n = int(float(config.get("preview_n") or 0))
        if preview_n > 0:
            eligible = eligible[:preview_n]
            total_eligible = len(eligible)
            log(f"Preview mode: limiting to first {total_eligible} row(s).")

        # Test send mode: send only first eligible row to self, prefix subject
        test_to_self = bool(config.get("test_to_self"))
        if test_to_self:
            eligible = eligible[:1]
            total_eligible = len(eligible)
            log("Test mode: sending first eligible row to Sender address (Bcc/Cc suppressed).")

        # Optional per-domain cooldown
        last_domain_at: Dict[str, float] = {}
        domain_cooldown = float(config.get("domain_throttle") or 0.0)

        sent_ok = 0
        for i, rec in enumerate(eligible, start=1):
            if stop_event.is_set():
                log("Cancelled by user.")
                break

            email = rec.get("email", "").strip()
            name = rec.get("name", "").strip()
            role = rec.get("role", "").strip()
            company = rec.get("company", "").strip()

            ctx = {
                "email": email, "name": name, "role": role, "company": company,
                "stage": rec.get("stage", ""),
                "reason": rec.get("reason", ""),
                "application_date": rec.get("application_date", ""),
                "sender_name": os.environ.get("SENDER_NAME", config.get("sender_name") or "Recruiting Team"),
                "sender_title": os.environ.get("SENDER_TITLE", config.get("sender_title") or "Talent Acquisition"),
            }

            subject_base = Template(config["subject"]).render(**ctx)
            subject = f"[TEST] {subject_base}" if test_to_self else subject_base

            text = render_template(config.get("text_template") or "", ctx)
            html = render_template(config.get("html_template") or "", ctx) if config.get("html_template") else None

            # template guard
            if not ((text and text.strip()) or (html and _strip_html(html))):
                log(f"   Error: rendered templates are empty; skipping {email or '(no email)'}")
                logq.put(f"__PROG__{i}/{total_eligible}")
                continue

            to_addr = config["sender"] if test_to_self else email
            bcc = None if test_to_self else (config.get("bcc") or None)
            cc  = None if test_to_self else (config.get("cc") or None)

            body = build_mime(
                sender=config["sender"], to=to_addr, subject=subject,
                text=text, html=html, cc=cc, bcc=bcc, reply_to=(config.get("reply_to") or None),
                attachments=config.get("attachments") or []
            )

            log(f"[{i}/{total_eligible}] {'DRY' if config['dry_run'] else ('TEST' if test_to_self else 'SEND')} → {to_addr} | CC: {cc or '-'} | BCC: {bcc or '-'} | {subject}")
            logq.put(f"__PROG__{i}/{total_eligible}")

            if config["dry_run"]:
                continue

            # per-domain cooldown
            if domain_cooldown:
                domain = (to_addr.split("@")[-1] if "@" in to_addr else "").lower()
                now = time.time()
                last = last_domain_at.get(domain)
                if last is not None:
                    dt = now - last
                    if dt < domain_cooldown:
                        time.sleep(domain_cooldown - dt)
                last_domain_at[domain] = time.time()

            try:
                send_gmail(gsvc, body)
                if not test_to_self:
                    write_status(
                        ssvc=ssvc, spreadsheet_id=config["spreadsheet_id"], tab_title=actual_tab,
                        row_number=rec["_row_number"],
                        status_col_index=hdr_index["sent_status"],
                        time_col_index=hdr_index["sent_at"]
                    )
                sent_ok += 1
                time.sleep(float(config.get("throttle", 2.0)))
            except Exception as e:
                log(f"   Error: {e}")

        log(f"Done. Processed {len(eligible)} eligible rows in this run. {'Sent ' + str(sent_ok) if not config['dry_run'] else 'No emails sent (dry run).' }")
    except Exception as e:
        # Any PermissionError from token writing or other exceptions surface here
        logq.put(f"FATAL: {e}")
