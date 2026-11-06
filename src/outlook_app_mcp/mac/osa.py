"""macOS backend: running the AppleScripts, and the small formats they speak.

Each script in ``scripts/read`` or ``scripts/write`` is run with
``scripts/common.applescript`` prepended, fed to ``osascript -`` on stdin.
Caller values travel only as ``argv``; they are never spliced into script text.
"""

from __future__ import annotations

import re
import subprocess
from collections.abc import Callable, Sequence
from datetime import datetime, timedelta
from pathlib import Path

from ..common import OutlookError, iso
from ..common import check_ids as _nonempty_ids

SCRIPTS = Path(__file__).parent / "scripts"
COMMON = SCRIPTS / "common.applescript"

FS = "\x1f"  # between fields
RS = "\x1e"  # between records
GS = "\x1d"  # between list entries inside a field
US = "\x1c"  # between sub-fields of a list entry

SCRIPT_TIMEOUT = 300
# Scripts that can send: Outlook for Mac asks the user to confirm every scripted send,
# so these wait for that answer.
SENDING_SCRIPTS = {"write/compose", "write/send_draft", "write/reply", "write/event", "write/cancel", "write/respond"}
SEND_TIMEOUT = 620

Runner = Callable[[str, Sequence[str]], str]


class FolderChanged(OutlookError):
    """A folder gained or lost items while a script read it in bulk."""


def script_source(name: str) -> str:
    """common.applescript + scripts/<name>.applescript, e.g. name='read/index'."""
    path = SCRIPTS / f"{name}.applescript"
    if not path.is_file():
        raise OutlookError(f"No such script: {name}")
    return COMMON.read_text(encoding="utf-8") + "\n" + path.read_text(encoding="utf-8")


def run_osascript(name: str, args: Sequence[str]) -> str:
    for a in args:
        if "\x00" in a:
            raise OutlookError("Text may not contain NUL characters.")
    try:
        proc = subprocess.run(
            ["osascript", "-", *args],
            check=False,
            input=script_source(name),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=SEND_TIMEOUT if name in SENDING_SCRIPTS else SCRIPT_TIMEOUT,
        )
    except subprocess.TimeoutExpired as exc:
        if name.startswith("write/"):
            # The Apple Event already reached Outlook; killing osascript doesn't undo it.
            raise OutlookError(
                f"Outlook did not answer within {SCRIPT_TIMEOUT} seconds, but the change may still have "
                "happened. Check in Outlook (Sent Items, Drafts, Deleted Items, the calendar) before trying again."
            ) from exc
        raise OutlookError(
            f"Outlook did not answer within {SCRIPT_TIMEOUT} seconds. "
            "Narrow the request (one folder, fewer days) and try again."
        ) from exc
    if proc.returncode != 0:
        stderr = proc.stderr.strip()
        if "FOLDER_CHANGED" in stderr:
            raise FolderChanged(stderr)
        raise OutlookError(explain(stderr))
    out = proc.stdout
    out = out.removesuffix("\n")
    return out


def explain(stderr: str) -> str:
    if "(-1743)" in stderr:
        return (
            "macOS is not allowing this app to control Outlook. Allow it in "
            "System Settings > Privacy & Security > Automation. Details: " + stderr
        )
    if "(-1712)" in stderr:
        return (
            "Outlook did not answer in time. If Outlook is showing 'A script is attempting to send a message', "
            "it is waiting for you: click Send to send or Cancel to keep it unsent; nothing goes out until you "
            "click Send. Details: " + stderr
        )
    if "(-128)" in stderr:
        return "The send was cancelled in Outlook's confirmation prompt; nothing was sent. Details: " + stderr
    if "NOT_SYNCED" in stderr:
        return "Outlook could not save the meeting's attendees. Details: " + stderr
    if "NO_OCCURRENCE" in stderr:
        return "The series has no occurrence at that time. Use the start list_events shows. Details: " + stderr
    if "NOT_EXCHANGE" in stderr:
        return "That needs an Exchange (Microsoft 365) account. Details: " + stderr
    if "NO_SIGNATURE" in stderr:
        return "Outlook has no signature with that name. Use list_signatures. Details: " + stderr
    if "(-600)" in stderr:
        return "Outlook is not running and could not be started. Details: " + stderr
    if "(-1728)" in stderr:
        return "Outlook could not find that item (wrong id, or it was moved or deleted). Details: " + stderr
    if "NO_ACCOUNT" in stderr:
        return "Outlook has no account with that address. Details: " + stderr
    if "ALREADY_DELETED" in stderr:
        return (
            "A message is already in a deleted-items folder; deleting it again would delete it "
            "permanently. Nothing was deleted. Details: " + stderr
        )
    if "HAS_ATTENDEES" in stderr:
        return (
            "This event is a meeting with attendees. It was not deleted. To leave it, use "
            "respond_to_meeting with response='decline'. Details: " + stderr
        )
    return "Outlook AppleScript error: " + stderr


# -- records -------------------------------------------------------------------


def records(raw: str) -> list[list[str]]:
    if raw == "":
        return []
    return [rec.split(FS) for rec in raw.split(RS)]


def entries(field: str) -> list[list[str]]:
    if field == "":
        return []
    return [e.split(US) for e in field.split(GS)]


def addr(entry: str) -> dict:
    name, _, address = entry.partition(US)
    return {"name": name, "address": address}


def addrs(field: str) -> list[dict]:
    return [addr(e) for e in field.split(GS)] if field else []


def as_bool(text: str) -> bool:
    return text.strip().lower() == "true"


# -- enumerations --------------------------------------------------------------

_ENUM_CODES = {
    "FlNF": "not flagged",
    "FlNC": "not completed",
    "FlCo": "completed",
    "eSBu": "busy",
    "eSFr": "free",
    "eSTe": "tentative",
    "eSOO": "out of office",
    "ASac": "accepted",
    "ASte": "tentatively accepted",
    "ASde": "declined",
    "ASnr": "not responded",
    "ATrq": "required attendee type",
    "ATop": "optional attendee type",
    "ATrs": "resource attendee type",
    "PrHi": "priority high",
    "PrNr": "priority normal",
    "PrLo": "priority low",
    "MTrq": "request meeting type",
    "MTre": "response meeting type",
    "MTca": "cancellation meeting type",
    "MTco": "counter proposal meeting type",
    "eNEt": "no end type",
    "eEDt": "end date type",
    "eENt": "end numbered type",
    "eRdp": "daily",
    "eRwp": "weekly",
    "eRrm": "relative monthly",
    "eRam": "absolute monthly",
    "eRry": "relative yearly",
    "eRay": "absolute yearly",
}
_RAW_CONSTANT = re.compile(r"«constant ....(....)»")


def enum(text: str) -> str:
    """Outlook enum text in either form ('busy' or '«constant ****eSBu»') -> its term."""
    m = _RAW_CONSTANT.fullmatch(text.strip())
    if m:
        return _ENUM_CODES.get(m.group(1), text)
    return text.strip()


def flag(text: str) -> str | None:
    return {"not completed": "flagged", "completed": "completed"}.get(enum(text))


def priority(text: str) -> str | None:
    """'priority high' -> 'high'; empty -> None."""
    v = enum(text)
    return v.removeprefix("priority ") if v else None


def to_priority(value: str) -> str:
    """'high' -> the AppleScript term the write scripts expect."""
    return {"low": "priority low", "normal": "priority normal", "high": "priority high"}[value]


def attendee_type(text: str) -> str:
    return enum(text).replace(" attendee type", "")


# -- dates ---------------------------------------------------------------------


def parse_stamp(text: str) -> datetime | None:
    """'Y-M-D-S' local time from AppleScript -> aware datetime."""
    if text == "":
        return None
    y, mo, d, s = (int(x) for x in text.split("-"))
    return (datetime(y, mo, d) + timedelta(seconds=s)).astimezone()


def stamp_iso(text: str) -> str | None:
    return iso(parse_stamp(text))


def to_stamp(dt: datetime) -> str:
    """Aware or naive (= local) datetime -> 'Y-M-D-S' in the Mac's local time."""
    if dt.tzinfo is not None:
        dt = dt.astimezone().replace(tzinfo=None)
    secs = dt.hour * 3600 + dt.minute * 60 + dt.second
    return f"{dt.year}-{dt.month}-{dt.day}-{secs}"


# -- ids ---------------------------------------------------------------------


def check_ids(ids: Sequence[str] | str, what: str = "message_ids") -> list[str]:
    """Outlook for Mac ids are whole numbers; they travel as text."""
    out = _nonempty_ids(ids, what)
    bad = [i for i in out if not i.isdigit()]
    if bad:
        raise OutlookError(f"{what}: {', '.join(repr(b) for b in bad)} is not an Outlook for Mac id (a number).")
    return [str(int(i)) for i in out]


def check_id(value: str | int, what: str) -> str:
    return check_ids([value], what)[0]


def mailbox_key(exchange_id: str) -> str | None:
    """The mailbox GUID inside an Exchange item id, or None.

    The id is base64url of: a format byte, a type byte, the length n of the mailbox
    GUID text, the GUID, then the item part. With format 1 the GUID is compressed:
    a byte below 0x20 repeats the previous character that many times (0 = none).
    Folders and calendars of one account share the GUID, which is how Google
    calendars, which have no account in Outlook's dictionary, are tied to theirs."""
    import base64
    import binascii

    x = (exchange_id or "").strip()
    if not x:
        return None
    try:
        b = base64.urlsafe_b64decode(x + "=" * (-len(x) % 4))
    except (binascii.Error, ValueError):
        return None
    if len(b) < 4 or b[2] == 0:
        return None
    n = b[2]
    if b[0] == 0:
        text = b[3 : 3 + n].decode("latin-1")
    else:
        out: list[str] = []
        for byte in b[3:]:
            if len(out) >= n:
                break
            if byte < 0x20:
                if out:
                    out.extend(out[-1] * byte)
            else:
                out.append(chr(byte))
        text = "".join(out)[:n]
    return text.lower() if len(text) == n else None
