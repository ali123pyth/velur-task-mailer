#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Velur — Daily Team Task Mailer
==============================

Reads the live Velur task sheet from Google Sheets, builds a per-holder task
update, and sends one email per holder + one consolidated all-team copy to Ali,
all FROM a single Velur address (a.ismail@velurfragrance.com) over SMTP.

This replaces the Zapier-based routine with a free, self-owned script.

Two ways to read the sheet (pick one via environment variables):
  A) Service account  -> set GOOGLE_SA_JSON to the path of a service-account
     JSON key, and share the sheet (Viewer) with that service account's email.
     Recommended for a private sheet.
  B) Public CSV export -> leave GOOGLE_SA_JSON unset. The sheet must be shared
     as "Anyone with the link -> Viewer". No credentials needed.

All email + sheet settings come from environment variables (see .env.example).
Set DRY_RUN=1 to print the emails instead of sending them (test first!).

Author: built for Velur. Python 3.9+.
"""

import os
import csv
import ssl
import sys
import html
import smtplib
import urllib.request
from io import StringIO
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate
from datetime import datetime, timezone, timedelta

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:  # pragma: no cover
    ZoneInfo = None


# ---------------------------------------------------------------------------
# Load settings from a local .env file (no external library needed).
# Real environment variables (e.g. CI secrets) take priority over the file.
# Looks for .env next to this script first, then in the current folder.
# ---------------------------------------------------------------------------
def _load_dotenv():
    here = os.path.dirname(os.path.abspath(__file__))
    for path in (os.path.join(here, ".env"), ".env"):
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
        break  # first file found wins


_load_dotenv()


# ---------------------------------------------------------------------------
# Configuration (from environment / .env)
# ---------------------------------------------------------------------------
SPREADSHEET_ID = os.environ.get(
    "VELUR_SPREADSHEET_ID",
    "1biIY7YBiPQN69CgYWeFdnp9_0lVgN-y9WJweI_W6NVs",
)
WORKSHEET_NAME = os.environ.get("VELUR_WORKSHEET", "Sheet1")
WORKSHEET_GID = os.environ.get("VELUR_WORKSHEET_GID", "0")  # for CSV export mode

SMTP_HOST = os.environ.get("VELUR_SMTP_HOST", "smtp.hostinger.com")
SMTP_PORT = int(os.environ.get("VELUR_SMTP_PORT", "465"))
SMTP_USER = os.environ.get("VELUR_SMTP_USER", "a.ismail@velurfragrance.com")
SMTP_PASS = os.environ.get("VELUR_SMTP_PASS", "")

SENDER_ADDR = os.environ.get("VELUR_SENDER", "a.ismail@velurfragrance.com")
SENDER_NAME = os.environ.get("VELUR_SENDER_NAME", "Velur")

GOOGLE_SA_JSON = os.environ.get("GOOGLE_SA_JSON", "").strip()
DRY_RUN = os.environ.get("DRY_RUN", "0") == "1"

TIMEZONE = os.environ.get("VELUR_TZ", "Asia/Dubai")

# Ali receives the consolidated all-team copy.
ALL_TEAM_ADDR = os.environ.get("VELUR_ALL_TEAM_ADDR", "a.ismail@velurfragrance.com")


# ---------------------------------------------------------------------------
# Holder normalisation, recipients, and ordering
# ---------------------------------------------------------------------------
# Merge spelling variants -> canonical holder name.
HOLDER_ALIASES = {
    "mohd . kawass": "Mohd. Kawass",
    "mohd. kawass": "Mohd. Kawass",
    "mohd kawass": "Mohd. Kawass",
    "moustafa": "Mustafa",
    "mustafa": "Mustafa",
}

# Canonical holder -> recipient address.
HOLDER_EMAILS = {
    "Aghyad": "a.alnajjar@acumenlight.com",
    "Mohd. Kawass": "m.kawass@velurfragrance.com",
    "Asaad": "a.nashed@acumenlight.com",
    "Mustafa": "m.salih@acumenlight.com",
    "Amr": "a.alghazzi@velurfragrance.com",
    "Ali": "a.ismail@velurfragrance.com",
    "Emad": "a.ismail@velurfragrance.com",  # no own address -> goes to Ali
}

# Section order in the consolidated all-team email.
CONSOLIDATED_ORDER = ["Aghyad", "Mohd. Kawass", "Asaad", "Mustafa", "Amr", "Emad", "Ali"]

# Holder values that are codes/placeholders, not people -> ignore the row.
HOLDER_CODE_BLOCKLIST = {"MKTN", "BRAND", "PROD", "SALE", "FIN", "DEV", "MKT"}

# Statuses that count as "still needs work".
OPEN_STATUSES = {"progress", "plan", ""}


# ---------------------------------------------------------------------------
# Sheet reading
# ---------------------------------------------------------------------------
def read_rows_via_service_account():
    """Read all cell values using gspread + a service-account key."""
    import gspread  # pip install gspread google-auth
    gc = gspread.service_account(filename=GOOGLE_SA_JSON)
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except Exception:
        ws = sh.sheet1
    return ws.get_all_values()


def read_rows_via_csv_export():
    """Read the sheet through its public CSV export URL (no credentials).

    Requires the sheet to be shared as 'Anyone with the link -> Viewer'.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        f"/export?format=csv&gid={WORKSHEET_GID}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "velur-task-mailer"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        ctype = resp.headers.get("Content-Type", "")
        data = resp.read().decode("utf-8", errors="replace")
    # If the sheet isn't link-shared, Google returns an HTML sign-in page with
    # HTTP 200 instead of CSV. Detect that and fail with a clear, actionable msg.
    if "text/csv" not in ctype or data.lstrip().lower().startswith("<!doctype", 0) \
            or data.lstrip().lower().startswith("<html"):
        raise RuntimeError(
            "The sheet did not return CSV — it is probably not readable by link. "
            "Fix: open the sheet -> Share -> 'Anyone with the link -> Viewer', "
            "or use a service account by setting GOOGLE_SA_JSON."
        )
    return list(csv.reader(StringIO(data)))


def read_all_values():
    if GOOGLE_SA_JSON:
        return read_rows_via_service_account()
    return read_rows_via_csv_export()


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def _norm(s):
    return (s or "").strip()


def find_header(rows):
    """Locate the main table's header row and return (row_index, col_map).

    col_map maps a logical field -> column index, resolved by header NAME so the
    script keeps working even if columns are reordered or inserted.
    """
    wanted = {
        "task": "task",
        "red flag": "red_flag",
        "status": "status",
        "progress %": "progress",
        "progress": "progress",
        "holder": "holder",
        "main category": "category",
    }
    for i, row in enumerate(rows):
        lowered = [_norm(c).lower() for c in row]
        if "task" in lowered and "holder" in lowered and "status" in lowered:
            col_map = {}
            for idx, name in enumerate(lowered):
                if name in wanted:
                    # don't let a later duplicate header overwrite the first
                    col_map.setdefault(wanted[name], idx)
            return i, col_map
    raise ValueError("Could not find the header row (need Task, Status, Holder).")


def canonical_holder(raw):
    h = _norm(raw)
    if not h:
        return None
    if h.upper() in HOLDER_CODE_BLOCKLIST:
        return None
    # all-caps short token that looks like a code
    if h.isupper() and len(h) <= 6 and " " not in h:
        return None
    return HOLDER_ALIASES.get(h.lower(), h)


def parse_progress(raw):
    """Return an int 0..100 if a number is present, else None."""
    s = _norm(raw).replace("%", "")
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def is_open(status, progress):
    st = _norm(status).lower()
    if st == "done":
        return False
    if progress is not None and progress >= 100:
        return False
    return st in OPEN_STATUSES


def get(row, col_map, key):
    idx = col_map.get(key)
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def build_holder_tasks(rows):
    """Return {holder: {'category': str, 'tasks': [ {name, progress, red} ]}}."""
    header_idx, col_map = find_header(rows)
    holders = {}  # canonical -> dict
    for row in rows[header_idx + 1:]:
        holder = canonical_holder(get(row, col_map, "holder"))
        if not holder:
            continue
        task = _norm(get(row, col_map, "task"))
        if not task:
            continue
        status = get(row, col_map, "status")
        progress = parse_progress(get(row, col_map, "progress"))
        if not is_open(status, progress):
            continue
        red = _norm(get(row, col_map, "red_flag")).lower() == "yes"
        category = _norm(get(row, col_map, "category"))

        bucket = holders.setdefault(holder, {"categories": [], "tasks": []})
        bucket["tasks"].append({"name": task, "progress": progress, "red": red})
        if category and category not in bucket["categories"]:
            bucket["categories"].append(category)
    return holders


# ---------------------------------------------------------------------------
# Email building
# ---------------------------------------------------------------------------
def today_str():
    tz = None
    if ZoneInfo is not None:
        try:
            tz = ZoneInfo(TIMEZONE)
        except Exception:
            # zoneinfo present but the tz database is missing on this host.
            # Install the 'tzdata' package to fix properly; meanwhile fall back
            # to a fixed UTC+4 (Dubai has no daylight saving) so we never crash.
            tz = None
    if tz is not None:
        now = datetime.now(tz)
    else:
        now = datetime.now(timezone.utc) + timedelta(hours=4)
    return now.strftime("%A, %d %B %Y")


def task_li(task):
    prefix = "⚠️ " if task["red"] else ""
    if task["progress"] is not None and task["progress"] > 0:
        tail = f"({task['progress']}%)"
    else:
        tail = "(planned)"
    return f"<li>{prefix}{html.escape(task['name'])} {tail}</li>"


def holder_section(holder, info):
    categories = " &amp; ".join(html.escape(c) for c in info["categories"]) or "Tasks"
    items = "".join(task_li(t) for t in info["tasks"])
    return (
        f'<p style="font-weight:bold;font-size:15px;margin:18px 0 6px;">'
        f"{html.escape(holder)} &mdash; {categories}</p>"
        f'<ul style="margin:0 0 4px 18px;padding:0;">{items}</ul>'
    )


def wrap_html(greeting_name, inner):
    greeting = (
        f"Good morning {html.escape(greeting_name)} &mdash; please reply with a quick "
        f"status on each of your tasks below, even if there's no change since "
        f"yesterday (just reply 'no update'), before end of day."
    )
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
        'color:#2b2b2b;line-height:1.6;">'
        f"<p>{greeting}</p>"
        f"{inner}"
        '<p style="margin-top:18px;">Thank you &mdash; Velur</p>'
        "</div>"
    )


def build_emails(holders, date_label):
    """Return (per_holder_emails, consolidated_email).

    per_holder_emails: list of (to_addr, subject, html_body, holder)
    consolidated_email: (to_addr, subject, html_body) or None
    """
    per = []
    for holder, info in holders.items():
        if not info["tasks"]:
            continue
        to_addr = HOLDER_EMAILS.get(holder)
        if not to_addr:
            # holder not in the recipient map -> only appears in all-team copy
            continue
        subject = f"🌿 Velur — Your Task Update — {date_label}"
        body = wrap_html(holder, holder_section(holder, info))
        per.append((to_addr, subject, body, holder))

    # Consolidated (all holders with open tasks), ordered.
    ordered = [h for h in CONSOLIDATED_ORDER if h in holders]
    ordered += [h for h in holders if h not in CONSOLIDATED_ORDER]
    sections = "".join(holder_section(h, holders[h]) for h in ordered if holders[h]["tasks"])
    consolidated = None
    if sections:
        subject = f"🌿 Velur — Daily Task Update (All Team) — {date_label}"
        body = wrap_html("Ali", sections)
        consolidated = (ALL_TEAM_ADDR, subject, body)
    return per, consolidated


# ---------------------------------------------------------------------------
# Sending
# ---------------------------------------------------------------------------
def make_message(to_addr, subject, html_body):
    msg = MIMEText(html_body, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = formataddr((SENDER_NAME, SENDER_ADDR))
    msg["To"] = to_addr
    msg["Date"] = formatdate(localtime=True)
    return msg


import time as _time

def open_smtp():
    try:
        import certifi
        ctx = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        ctx = ssl.create_default_context()
    last = None
    for attempt in range(4):
        server = None
        try:
            if SMTP_PORT == 465:
                server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx, timeout=45, local_hostname="velurfragrance.com")
            else:
                server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=45, local_hostname="velurfragrance.com")
            _time.sleep(2)  # wait before speaking: Hostinger drops "early talkers"
            server.ehlo()
            if SMTP_PORT != 465:
                server.starttls(context=ctx)
                server.ehlo()
            server.login(SMTP_USER, SMTP_PASS)
            return server
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException, OSError) as e:
            last = e
            if server is not None:
                try:
                    server.close()
                except Exception:
                    pass
            _time.sleep(5)
    raise last


def send_all(per, consolidated):
    outgoing = [(t, s, b, who) for (t, s, b, who) in per]
    if consolidated:
        outgoing.append((consolidated[0], consolidated[1], consolidated[2], "ALL-TEAM"))

    if DRY_RUN:
        print("=== DRY RUN — nothing sent ===")
        for to_addr, subject, _body, who in outgoing:
            print(f"  [{who}] -> {to_addr} | {subject}")
        return

    if not SMTP_PASS:
        sys.exit("ERROR: VELUR_SMTP_PASS is not set. Aborting before sending.")

    _stamp = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".last_sent")
    _today = today_str()
    if os.environ.get("FORCE_SEND", "0") != "1" and os.path.exists(_stamp):
        try:
            if open(_stamp, encoding="utf-8").read().strip() == _today:
                print(f"Already sent for {_today} - skipping (set FORCE_SEND=1 to override).")
                return
        except Exception:
            pass

    server = open_smtp()
    try:
        for to_addr, subject, body, who in outgoing:
            msg = make_message(to_addr, subject, body)
            server.sendmail(SENDER_ADDR, [to_addr], msg.as_string())
            print(f"Sent [{who}] -> {to_addr}")
        try:
            with open(_stamp, "w", encoding="utf-8") as _fh:
                _fh.write(_today)
        except Exception:
            pass
    finally:
        server.quit()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    date_label = today_str()
    print(f"Velur task mailer — {date_label}")

    rows = read_all_values()
    holders = build_holder_tasks(rows)

    if not holders:
        print("No open tasks found for any holder. Nothing to send.")
        return

    per, consolidated = build_emails(holders, date_label)

    print("\nOpen-task summary:")
    for holder, info in holders.items():
        addr = HOLDER_EMAILS.get(holder, "(all-team only)")
        print(f"  {holder}: {len(info['tasks'])} task(s) -> {addr}")

    print()
    send_all(per, consolidated)
    print("\nDone.")


if __name__ == "__main__":
    main()
