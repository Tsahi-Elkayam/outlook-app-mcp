"""Windows backend: talking to classic Outlook through COM, and the small formats it needs.

Everything here is plain object-model access; nothing runs macros, scripts or
Outlook's "Execute" commands. pywin32 is imported only when a real Outlook is
opened, so the backend can be tested anywhere with a fake object model.
"""

from __future__ import annotations

import hashlib
import re
import time
from datetime import UTC, datetime
from functools import wraps

from ..common import OutlookError, iso

# Outlook enumerations (OlDefaultFolders, OlItemType, ...), from the object model reference.
FOLDER_DELETED = 3
FOLDER_OUTBOX = 4
FOLDER_SENT = 5
FOLDER_INBOX = 6
FOLDER_CALENDAR = 9
FOLDER_CONTACTS = 10
FOLDER_NOTES = 12
FOLDER_TASKS = 13
FOLDER_DRAFTS = 16
FOLDER_JUNK = 23
ROLE_FOLDERS = {
    "inbox": FOLDER_INBOX,
    "sent": FOLDER_SENT,
    "drafts": FOLDER_DRAFTS,
    "deleted": FOLDER_DELETED,
    "junk": FOLDER_JUNK,
    "outbox": FOLDER_OUTBOX,
}

ITEM_MAIL = 0
ITEM_APPOINTMENT = 1
ITEM_CONTACT = 2
ITEM_TASK = 3
ITEM_NOTE = 5
ITEM_DISTLIST = 7

RECIPIENT_TO, RECIPIENT_CC, RECIPIENT_BCC = 1, 2, 3
ATTENDEE_TYPES = {1: "required", 2: "optional", 3: "resource"}
RESPONSE_STATUS = {0: "none", 1: "organizer", 2: "tentatively accepted", 3: "accepted", 4: "declined",
                   5: "not responded"}
BUSY_STATUS = {0: "free", 1: "tentative", 2: "busy", 3: "out of office", 4: "working elsewhere"}
BUSY_CODES = {"free": 0, "tentative": 1, "busy": 2, "oof": 3}
MY_RESPONSE = {0: "none", 1: "organizer", 2: "tentative", 3: "accepted", 4: "declined", 5: "not responded"}
RESPONSE_ORGANIZER = 1
MEETING_NONE, MEETING_ORGANIZED, MEETING_RECEIVED, MEETING_CANCELED = 0, 1, 3, 5  # OlMeetingStatus
IMPORTANCE = {0: "low", 1: "normal", 2: "high"}
IMPORTANCE_CODES = {"low": 0, "normal": 1, "high": 2}
RECURS_DAILY, RECURS_WEEKLY, RECURS_MONTHLY, RECURS_MONTH_NTH, RECURS_YEARLY, RECURS_YEAR_NTH = 0, 1, 2, 3, 5, 6
RECURRENCE_TYPES = {"daily": RECURS_DAILY, "weekly": RECURS_WEEKLY, "monthly": RECURS_MONTHLY,
                    "yearly": RECURS_YEARLY}
DAY_MASK = {"sunday": 1, "monday": 2, "tuesday": 4, "wednesday": 8, "thursday": 16, "friday": 32, "saturday": 64}
# OlCategoryColor, by value.
CATEGORY_COLORS = {
    0: None, 1: "red", 2: "orange", 3: "peach", 4: "yellow", 5: "green", 6: "teal", 7: "olive", 8: "blue",
    9: "purple", 10: "maroon", 11: "steel", 12: "dark steel", 13: "gray", 14: "dark gray", 15: "black",
    16: "dark red", 17: "dark orange", 18: "dark peach", 19: "dark yellow", 20: "dark green", 21: "dark teal",
    22: "dark olive", 23: "dark blue", 24: "dark purple", 25: "dark maroon",
}
ACCOUNT_KINDS = {0: "exchange", 1: "imap", 2: "pop", 3: "other", 4: "other", 5: "other"}
MEETING_RESPONSES = {"tentative": 2, "accept": 3, "decline": 4}
FLAG_NONE, FLAG_COMPLETE, FLAG_MARKED = 0, 1, 2
FORMAT_HTML = 2
SENSITIVITY_PRIVATE = 2
EXCHANGE_NOT_EXCHANGE = 3  # OlExchangeStoreType: PST, IMAP or other non-Exchange store
EXCHANGE_PUBLIC_FOLDERS = 2  # OlExchangeStoreType.olExchangePublicFolder

# MAPI properties read through PropertyAccessor or Table columns.
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_ATTACH_MIME = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
PR_FLAG_STATUS = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"
PR_OOF_STATE = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_CONVERSATION_TOPIC = "http://schemas.microsoft.com/mapi/proptag/0x0070001F"
PR_HASATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"

NOT_INSTALLED_HINT = (
    "Could not start classic Outlook through COM. This server needs classic Outlook for Windows "
    "(Outlook from Microsoft 365 / Office, not the New Outlook app, which has no COM interface). "
    "If classic Outlook is installed, run it once, and make sure Claude and Outlook run at the same "
    "privilege level (both elevated or neither)."
)

# Outlook never stores "no date" as empty; it uses 1 Jan 4501.
_NO_DATE_YEAR = 4501
NO_DATE = datetime(_NO_DATE_YEAR, 1, 1, tzinfo=UTC)  # what to write to clear a date (marked UTC: pywin32 leaves it)


def com_error_type() -> type[BaseException]:
    try:
        import pywintypes  # type: ignore[import-not-found]

        return pywintypes.com_error
    except ImportError:  # not on Windows: tests use fake objects that raise ComError
        return ComError


class ComError(Exception):
    """Stand-in for pywintypes.com_error where pywin32 is not installed (tests)."""

    def __init__(self, hresult: int = -2147352567, text: str = "", excepinfo=None):
        super().__init__(hresult, text, excepinfo, None)
        self.hresult = hresult


_NOT_FOUND = {-2147221233, -2147024809}  # MAPI_E_NOT_FOUND, E_INVALIDARG from GetItemFromID
# Outlook went away (closed, crashed, updated): the cached Application is dead.
_DISCONNECTED = {-2147417848, -2147023174, -2147417849, -2147023170, -2147220995}
# Outlook is showing a dialog and refuses calls for now; the call did not run.
_BUSY = {-2147418111, -2147417846}
BUSY_WAIT_SECONDS = 10


def hresults(exc: BaseException) -> set[int]:
    """The HRESULT of a com_error and, when Outlook wrapped it in DISP_E_EXCEPTION,
    the scode inside its excepinfo."""
    args = getattr(exc, "args", ())
    codes = {args[0]} if args and isinstance(args[0], int) else set()
    info = args[2] if len(args) > 2 else None
    if isinstance(info, tuple) and len(info) > 5 and isinstance(info[5], int):
        codes.add(info[5])
    return codes


def explain_com(exc: BaseException) -> str:
    args = getattr(exc, "args", ())
    hresult = args[0] if args else None
    detail = ""
    if len(args) > 2 and args[2]:
        info = args[2]
        detail = (info[2] or "") if isinstance(info, tuple) and len(info) > 2 else str(info)
    text = detail or (args[1] if len(args) > 1 else str(exc))
    codes = hresults(exc)
    if codes & _NOT_FOUND or "could not be found" in str(text).lower():
        return f"Outlook could not find that item (wrong id, or it was moved or deleted). Details: {text}"
    if codes & _BUSY:
        return "Outlook is busy with an open dialog. Close the dialog in Outlook and try again."
    if codes & _DISCONNECTED:
        return "Outlook was closed or restarted. Open Outlook and try again."
    return f"Outlook COM error: {text} ({hresult})"


def com_errors(fn):
    """Turn COM errors raised inside a backend method into OutlookError.

    While Outlook shows a dialog it rejects calls without running them, so those are
    retried for a few seconds. If Outlook was restarted, the connection is made again;
    a reading method is then retried once, a changing one is not (part of it may have
    run), and the caller is told to check before repeating it."""

    @wraps(fn)
    def wrapper(self, *args, **kwargs):
        deadline = time.monotonic() + BUSY_WAIT_SECONDS
        reconnected = False
        while True:
            try:
                return fn(self, *args, **kwargs)
            except OutlookError:
                raise
            except self._com_error as exc:
                codes = hresults(exc)
                if codes & _BUSY and time.monotonic() < deadline:
                    time.sleep(0.5)
                    continue
                if codes & _DISCONNECTED and not reconnected and self._reconnect():
                    reconnected = True
                    if fn.__name__ in READ_METHODS:
                        continue
                    raise OutlookError(
                        "Outlook was restarted while this ran; the connection is back. Check in Outlook whether "
                        "the change happened before trying again."
                    ) from exc
                raise OutlookError(explain_com(exc)) from exc

    return wrapper


READ_METHODS = {"list_folders", "list_messages", "search", "read_message", "list_calendars", "list_events",
                "read_event", "search_contacts", "list_accounts", "get_conversation", "list_signatures",
                "list_categories", "get_auto_reply", "get_free_busy", "list_tasks", "read_task", "list_notes",
                "read_note", "read_contact", "lookup_directory", "list_contact_groups"}


def open_outlook():
    """The running Outlook.Application, started if needed. Must be called on the thread
    that will use it (COM objects belong to the thread that made them)."""
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise OutlookError("pywin32 is missing; install the package with its Windows dependencies.") from exc
    pythoncom.CoInitialize()
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except pythoncom.com_error as exc:
        raise OutlookError(f"{NOT_INSTALLED_HINT} Details: {explain_com(exc)}") from exc


# -- ids -----------------------------------------------------------------------
# An id is "<store key>.<EntryID>". The store key is a short hash of the store's
# StoreID, so ids stay short and work for items in any open mailbox or data file.

_ID = re.compile(r"^([0-9a-f]{10})\.([0-9A-Fa-f]{8,})$")


def store_key(store_id: str) -> str:
    return hashlib.sha1(store_id.encode("ascii", "replace")).hexdigest()[:10]


def make_id(key: str, entry_id: str) -> str:
    return f"{key}.{entry_id}"


def split_id(value: str, what: str) -> tuple[str, str]:
    m = _ID.match((value or "").strip())
    if not m:
        raise OutlookError(f"{what}: {value!r} is not an id from this Outlook (expected '<store>.<EntryID>').")
    return m.group(1), m.group(2)


# -- dates ---------------------------------------------------------------------


def from_com(value) -> datetime | None:
    """A date from the object model -> aware local datetime, or None.

    Outlook's object model gives local times. pywin32 hands them back marked as UTC
    without converting, so the zone is dropped and the time read as local."""
    if value is None or not isinstance(value, datetime):
        return None
    if value.year >= _NO_DATE_YEAR:
        return None
    naive = datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    return naive.astimezone()


def com_iso(value) -> str | None:
    return iso(from_com(value))


def to_com(dt: datetime) -> datetime:
    """A datetime to hand to the object model, which wants local wall-clock time.

    pywin32 converts every datetime it sends to UTC first (PyWinObject_AsSYSTEMTIME
    calls astimezone(UTC)); a naive one is taken as local and shifted. Labelling the
    local wall-clock time as UTC makes that conversion a no-op, the mirror of how
    pywin32 hands dates back (local fields labelled UTC, see from_com)."""
    return dt.astimezone().replace(tzinfo=UTC, microsecond=0)  # naive input counts as local


def dasl_utc(dt: datetime) -> str:
    """A DASL filter date: UTC, in a form Outlook reads the same in every locale."""
    return dt.astimezone(UTC).strftime("%Y-%m-%d %H:%M")


def dasl_text(value: str) -> str:
    """A text literal for a DASL filter: single quotes doubled."""
    return "'" + value.replace("'", "''") + "'"


def jet_local(dt: datetime) -> str:
    """A Jet filter date ("[Start] < '...'"), which Outlook reads in the user's own
    date format. Uses Windows' own formatting so it parses back in any locale."""
    local = dt.astimezone().replace(tzinfo=None)
    import ctypes  # Windows only; never reached in tests (they patch this function)
    from ctypes import wintypes

    class SYSTEMTIME(ctypes.Structure):
        _fields_ = [(n, wintypes.WORD) for n in ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute",
                                                 "wSecond", "wMilliseconds")]

    st = SYSTEMTIME(local.year, local.month, 0, local.day, local.hour, local.minute, 0, 0)
    buf = ctypes.create_unicode_buffer(128)
    kernel32 = ctypes.windll.kernel32
    date_shortdate, time_noseconds = 0x1, 0x2
    if not kernel32.GetDateFormatEx(None, date_shortdate, ctypes.byref(st), None, buf, 128, None):
        raise OutlookError("Windows could not format a date for Outlook's filter.")
    day = buf.value
    if not kernel32.GetTimeFormatEx(None, time_noseconds, ctypes.byref(st), None, buf, 128):
        raise OutlookError("Windows could not format a time for Outlook's filter.")
    return f"{day} {buf.value}"


# -- small readers -------------------------------------------------------------


def prop(obj, name: str, default=None):
    """An object-model property, or default when this kind of item lacks it."""
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001 - COM raises AttributeError or com_error for missing properties
        return default


def mapi(obj, schema: str, default=None):
    """A MAPI property through PropertyAccessor, or default when it is not set."""
    try:
        return obj.PropertyAccessor.GetProperty(schema)
    except Exception:  # noqa: BLE001 - not set on this item
        return default


def smtp_of_entry(entry, fallback: str = "") -> str:
    """The SMTP address of an AddressEntry. Exchange users carry an X.500 address
    in .Address; their SMTP address is in PR_SMTP_ADDRESS or the Exchange user."""
    if entry is None:
        return fallback
    address = prop(entry, "Address", "") or ""
    if "@" in address and not address.startswith("/"):
        return address
    smtp = mapi(entry, PR_SMTP_ADDRESS)
    if smtp:
        return smtp
    user = None
    try:
        user = entry.GetExchangeUser()
    except Exception:  # noqa: BLE001 - not an Exchange user
        user = None
    if user is not None:
        smtp = prop(user, "PrimarySmtpAddress", "")
        if smtp:
            return smtp
    return fallback or address


def recipient_address(recipient) -> str:
    return smtp_of_entry(prop(recipient, "AddressEntry"), prop(recipient, "Address", "") or "")


def sender_of(item) -> dict:
    name = prop(item, "SenderName", "") or ""
    address = prop(item, "SenderEmailAddress", "") or ""
    if "@" not in address or address.startswith("/"):
        address = mapi(item, PR_SENDER_SMTP) or smtp_of_entry(prop(item, "Sender"), address)
    return {"name": name, "address": address}


def split_categories(text) -> list[str]:
    return [c.strip() for c in re.split(r"[;,]", text or "") if c.strip()]


def categories(item) -> list[str]:
    return split_categories(prop(item, "Categories", ""))


LOCALE_SLIST = 0x0C  # GetLocaleInfoEx: the list separator of the user's regional settings


def list_separator() -> str:
    """The separator Outlook uses in an item's Categories string: the user's Windows list separator
    (',' in most locales, ';' in some, e.g. German). Falls back to ',' where Windows can't be asked."""
    try:
        import ctypes  # Windows only

        buf = ctypes.create_unicode_buffer(16)
        if ctypes.windll.kernel32.GetLocaleInfoEx(None, LOCALE_SLIST, buf, 16) and buf.value.strip():
            return buf.value.strip()
    except (ImportError, AttributeError, OSError):
        pass
    return ","


def write_categories(item, names) -> None:
    """Set an item's categories. Outlook keeps them as one string joined by the list separator."""
    item.Categories = (list_separator() + " ").join(names)


def date_com(value) -> str | None:
    """A date-only property (task due date, birthday) as YYYY-MM-DD, or None for no date."""
    d = from_com(value)
    return d.date().isoformat() if d else None


def day_to_com(dt: datetime) -> datetime:
    """The local calendar day of dt, at midnight, for properties that hold a date only."""
    local = dt.astimezone()
    return to_com(datetime(local.year, local.month, local.day).astimezone())


def flag_name(status) -> str | None:
    return {FLAG_MARKED: "flagged", FLAG_COMPLETE: "completed"}.get(status)


def walk_folders(folder, item_type: int):
    """Folders below `folder` whose default item type is `item_type`, skipping hidden
    system folders (Recipient Cache, GAL Contacts, Conversation History and the like)."""
    for f in each(folder.Folders):
        if mapi(f, PR_HIDDEN, False):
            continue
        if prop(f, "DefaultItemType", None) == item_type:
            yield f
        yield from walk_folders(f, item_type)


def each(collection):
    """Items of a 1-based COM collection."""
    n = collection.Count
    for i in range(1, n + 1):
        yield collection.Item(i)
