"""A small fake of Outlook's COM object model, enough to run the Windows backend
anywhere. It follows the real model's shapes (1-based collections, Items.Add,
GetTable/Columns/GetArray, GetItemFromID with a StoreID) and records the calls
that change things, so tests can check what the backend did.

Its Table understands exactly the DASL filters the backend builds; anything else
fails the test, so a change in filter syntax shows up here.

Where the real thing has traps, the fake has them too:
- Dates behave like pywin32: a datetime written to Outlook is converted with
  astimezone(UTC) and its fields stored as Outlook's local wall-clock time; a date
  read back carries those local fields labelled UTC.
- A Table refuses computed properties (UnRead, Body...) as columns.
- SendUsingAccount ignores plain assignment; only a PROPERTYPUTREF Invoke sets it.
- Accepting a meeting gives the appointment a new EntryID; declining removes it.
- Delete() fails on every item and every folder: the backend must never call it.
- GetRecurrencePattern() on a plain appointment turns it into a recurring one; a series that ends on a
  date and one that ends after a count both report Occurrences and PatternEndDate.
- Appointment.Send() fails without recipients; Categories.Add() fails for a name that exists;
  Recipient.FreeBusy() fails for people with no published free/busy; a NoteItem's Subject is read-only;
  GetGlobalAddressList() fails without an Exchange account; Folder.MoveTo() fails on a name clash.
"""

from __future__ import annotations

import itertools
import re
from datetime import UTC, datetime, timedelta

from outlook_app_mcp.windows.com import (
    EXCHANGE_PUBLIC_FOLDERS,
    ITEM_APPOINTMENT,
    ITEM_CONTACT,
    ITEM_DISTLIST,
    ITEM_MAIL,
    ITEM_NOTE,
    ITEM_TASK,
    PR_CONVERSATION_TOPIC,
    PR_FLAG_STATUS,
    PR_HASATTACH,
    PR_HIDDEN,
    PR_SENDER_SMTP,
    ComError,
)
from outlook_app_mcp.windows.mail import PR_MESSAGE_CLASS, PR_MESSAGE_FLAGS

NOT_FOUND = -2147221233
BLOCKED_PROPERTY = -2147352567
_ids = itertools.count(0xA000)

DATE_PROPS = {"Start", "End", "ReceivedTime", "SentOn", "Birthday", "DueDate", "StartDate", "DateCompleted",
              "ReminderTime", "LastModificationTime", "TaskStartDate", "TaskDueDate"}
NO_DATE = datetime(4501, 1, 1)
# Computed properties a Table refuses as columns (Microsoft: "Unsupported Properties in
# a Table Object or Table Filter"); the subset this backend could be tempted to use.
BLOCKED_COLUMNS = {"UnRead", "Body", "HTMLBody", "RTFBody", "BodyFormat", "Attachments", "Recipients",
                   "IsMarkedAsTask", "TaskCompletedDate", "SendUsingAccount", "Sender", "Parent"}


def to_outlook(value: datetime) -> datetime:
    """What pywin32 stores for a datetime: astimezone(UTC) fields, read by Outlook as local."""
    return value.astimezone(UTC).replace(tzinfo=None)


def new_entry_id() -> str:
    return f"00000000{next(_ids):08X}"


class Coll:
    """A 1-based COM collection."""

    def __init__(self, items=None):
        self._items = list(items or [])

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        if isinstance(i, int):
            return self._items[i - 1]
        raise TypeError(i)

    def __iter__(self):
        return iter(self._items)


class Accessor:
    def __init__(self, props):
        self._props = props

    def GetProperty(self, schema):
        if schema in self._props:
            return self._props[schema]
        raise ComError(-2147221233, "The property is not set")


class Recipient:
    def __init__(self, address, name="", type_=1, status=None, app=None):
        self.Address = address
        self.Name = name or address
        self.Type = type_
        self.MeetingResponseStatus = status
        self.AddressEntry = None
        self.Resolved = False
        self.app = app

    def Resolve(self):
        """Session.CreateRecipient(text).Resolve(): an SMTP address resolves; a name resolves when exactly
        one directory entry has it."""
        if "@" in self.Address:
            self.Resolved = True
            return True
        hits = [e for e in (self.app.gal or []) if e.Name.casefold() == self.Address.casefold()]
        if len(hits) == 1:
            self.AddressEntry, self.Resolved, self.Name = hits[0], True, hits[0].Name
        return self.Resolved

    def FreeBusy(self, start, minutes, complete):
        self.app.log.append(("freebusy", self.Address, start, minutes, complete))
        text = self.app.free_busy.get(self.Address)
        if text is None:
            raise ComError(-2147221233, "no free/busy published")
        return text


class Recipients(Coll):
    def __init__(self, items=None):
        super().__init__(items)
        self.resolved = False

    def Add(self, address):
        r = Recipient(address)
        self._items.append(r)
        return r

    def ResolveAll(self):
        self.resolved = True
        for r in self._items:
            r.Resolved = True
        return True


class Attachment:
    def __init__(self, name, data=b"x", mime=""):
        self.FileName = name
        self.DisplayName = name
        self.Size = len(data)
        self._data = data
        self.PropertyAccessor = Accessor({"http://schemas.microsoft.com/mapi/proptag/0x370E001F": mime} if mime else {})

    def SaveAsFile(self, path):
        from pathlib import Path

        if Path(path).exists():
            raise AssertionError("SaveAsFile would overwrite " + path)
        Path(path).write_bytes(self._data)


class Attachments(Coll):
    def Remove(self, index):
        del self._items[index - 1]

    def Add(self, path):
        from pathlib import Path

        a = Attachment(Path(path).name, Path(path).read_bytes())
        self._items.append(a)
        return a


class Item:
    """Common base: an item lives in a folder (Parent) and can move."""

    MessageClass = "IPM"

    def __setattr__(self, name, value):
        if name in DATE_PROPS and isinstance(value, datetime):
            self.__dict__.setdefault("_raw", {})[name] = to_outlook(value)
            return
        object.__setattr__(self, name, value)

    def __getattr__(self, name):
        raw = self.__dict__.get("_raw", {})
        if name in DATE_PROPS and name in raw:
            return raw[name].replace(tzinfo=UTC)  # local fields, labelled UTC, like pywin32
        raise AttributeError(name)

    def set_local(self, name, value: datetime | None):
        """Put Outlook's own local wall-clock value in place (test data)."""
        if value is not None:
            self.__dict__.setdefault("_raw", {})[name] = value

    def true_utc(self, name) -> datetime:
        return self._raw[name].astimezone(UTC)

    def table_value(self, col):
        if col == PR_MESSAGE_CLASS:
            return self.MessageClass
        if col == PR_MESSAGE_FLAGS:
            return 0 if getattr(self, "UnRead", False) else 1
        if col in BLOCKED_COLUMNS:
            raise AssertionError(f"blocked column {col} reached a row")
        return getattr(self, col, None)

    def Delete(self):
        raise AssertionError("Delete() must never be called; it is permanent inside Deleted Items")

    def __init__(self, **props):
        self.EntryID = new_entry_id()
        self.Parent = None
        self.saved = 0
        self.Categories = ""
        self.Body = ""
        self.Recipients = Recipients()
        self.PropertyAccessor = Accessor({})
        for k, v in props.items():
            setattr(self, k, v)

    def Save(self):
        self.saved += 1
        if self.Parent is None:
            raise AssertionError("saving an item that belongs to no folder")
        if self not in self.Parent._items:
            self.Parent._items.append(self)

    def Move(self, folder):
        self.Parent._items.remove(self)
        if folder.Store is not self.Parent.Store:
            self.EntryID = new_entry_id()
        self.Parent = folder
        folder._items.append(self)
        folder.app.log.append(("move", self.EntryID, folder.Name))
        return self


class Mail(Item):
    MessageClass = "IPM.Note"

    def __init__(self, subject="", sender=("", ""), received=None, unread=False, flag=0, sender_smtp="", **kw):
        super().__init__(**kw)
        self.Subject = subject
        self.SenderName, self.SenderEmailAddress = sender
        self.set_local("ReceivedTime", received or datetime(4501, 1, 1))
        self.set_local("SentOn", received or datetime(4501, 1, 1))
        self.UnRead = unread
        self.FlagStatus = flag
        self.BodyFormat = 2
        self.Importance = 1
        self.ConversationID = ""
        self.ConversationTopic = ""
        self.HTMLBody = "<html><body><p>original</p></body></html>"
        self.Attachments = Attachments()
        self._send_account = None
        self._oleobj_ = OleObj(self)
        self.sent = False
        self._smtp = sender_smtp

    @property
    def SendUsingAccount(self):
        return self._send_account

    @SendUsingAccount.setter
    def SendUsingAccount(self, acc):
        pass  # through pywin32's late binding a plain assignment does not take

    def table_value(self, col):
        if col == PR_SENDER_SMTP:
            return self._smtp
        if col == PR_FLAG_STATUS:
            return self.FlagStatus
        if col == PR_HASATTACH:
            return bool(self.Attachments._items)
        if col == PR_CONVERSATION_TOPIC:
            return self.ConversationTopic
        return super().table_value(col)

    @property
    def Sent(self):
        return self.sent

    def _kind(self, kind):
        return "; ".join(r.Address for r in self.Recipients._items if r.Type == kind)

    def _set_kind(self, kind, text):
        """Assigning To / CC / BCC rebuilds just the recipients of that kind, unresolved."""
        self.Recipients._items = [r for r in self.Recipients._items if r.Type != kind]
        for a in [x.strip() for x in text.split(";") if x.strip()]:
            self.Recipients._items.append(Recipient(a, type_=kind))

    To = property(lambda self: self._kind(1), lambda self, v: self._set_kind(1, v))
    CC = property(lambda self: self._kind(2), lambda self, v: self._set_kind(2, v))
    BCC = property(lambda self: self._kind(3), lambda self, v: self._set_kind(3, v))

    def Send(self):
        self.sent = True
        app = self.app
        app.log.append(("send", self.Subject))
        if self.Parent is not None and self in self.Parent._items:
            self.Parent._items.remove(self)

    def _response(self, prefix):
        r = Mail(subject=prefix + self.Subject)
        r.app = self.app
        r.HTMLBody = "<html><head></head><body lang=EN><p>quoted</p></body></html>"
        # Like Outlook's default: a saved reply lands in the default store's Drafts.
        r.Parent = self.app.default_drafts
        return r

    def Reply(self):
        return self._response("RE: ")

    def ReplyAll(self):
        return self._response("RE: ")

    def Forward(self):
        return self._response("FW: ")


class MeetingRequest(Mail):
    """A meeting request (or update, or cancellation) in the Inbox."""

    def __init__(self, subject="", message_class="IPM.Schedule.Meeting.Request", appointment=None, **kw):
        super().__init__(subject, **kw)
        self.MessageClass = message_class
        self.appointment = appointment

    def GetAssociatedAppointment(self, add_to_calendar):
        if self.appointment is None and add_to_calendar:
            appt = Appointment(self.Subject, meeting=3)
            appt.app = self.app
            appt.Parent = self.app.Session.default_calendar
            appt.Save()
            self.appointment = appt
        return self.appointment


class OleObj:
    """Mail._oleobj_: the raw IDispatch, for the one property pywin32 can't set by name."""

    def __init__(self, item):
        self.item = item
        self.calls = []

    def Invoke(self, dispid, lcid, flags, want_result, *args):
        self.calls.append((dispid, flags))
        if (dispid, flags) == (64209, 8):
            self.item._send_account = args[0]
            return
        raise AssertionError(f"unexpected Invoke {dispid} {flags}")


def _add_months(d: datetime, months: int) -> tuple[int, int]:
    total = d.year * 12 + d.month - 1 + months
    return total // 12, total % 12 + 1


def _clamped(year: int, month: int, day: int, like: datetime) -> datetime:
    last = (datetime(year + (month == 12), month % 12 + 1, 1) - timedelta(days=1)).day
    return like.replace(year=year, month=month, day=min(day, last))


class RecurrencePattern:
    """Outlook's RecurrencePattern, with the occurrence rules the backend relies on."""

    def __init__(self, appt):
        self.appt = appt
        self.RecurrenceType = 0
        self.Interval = 1
        self.DayOfWeekMask = 0
        self.DayOfMonth = 0
        self.MonthOfYear = 0
        self.PatternStartDate = None
        self._noend, self._count, self._until = True, None, None
        self.cache = {}

    NoEndDate = property(lambda self: self._noend, lambda self, v: self._set_noend(v))

    def _set_noend(self, value):
        self._noend = bool(value)
        if value:
            self._count = self._until = None

    def starts(self):
        base = self.appt._raw["Start"]
        out, k = [], 0
        while len(out) < 400 and k < 2000:
            if self.RecurrenceType == 0:
                cands = [base + timedelta(days=k * self.Interval)]
            elif self.RecurrenceType == 1:
                sunday = base - timedelta(days=(base.weekday() + 1) % 7)
                mask = self.DayOfWeekMask or (1 << ((base.weekday() + 1) % 7))
                cands = [sunday + timedelta(days=k * 7 * self.Interval + i) for i in range(7) if mask >> i & 1]
                cands = [c for c in cands if c.date() >= base.date()]
            elif self.RecurrenceType == 2:
                y, m = _add_months(base, k * self.Interval)
                cands = [_clamped(y, m, self.DayOfMonth or base.day, base)]
            else:
                y, m = base.year + k * self.Interval // 12, self.MonthOfYear or base.month
                cands = [_clamped(y, m, self.DayOfMonth or base.day, base)]
            k += 1
            for c in cands:
                if self._until is not None and c.date() > self._until.date():
                    return out
                if self._count is not None and len(out) >= self._count:
                    return out
                out.append(c)
        return out

    @property
    def Occurrences(self):
        return self._count if self._count is not None else len(self.starts())

    @Occurrences.setter
    def Occurrences(self, n):
        self._noend, self._count, self._until = False, n, None

    @property
    def PatternEndDate(self):
        if self._until is not None:
            return self._until.replace(tzinfo=UTC)
        if self._count is not None:
            last = self.starts()[-1]
            return datetime(last.year, last.month, last.day, tzinfo=UTC)
        return NO_DATE.replace(tzinfo=UTC)

    @PatternEndDate.setter
    def PatternEndDate(self, value):
        self._noend, self._count, self._until = False, None, to_outlook(value)

    def GetOccurrence(self, when):
        want = to_outlook(when)
        if want not in self.starts():
            raise ComError(BLOCKED_PROPERTY, "The specified date is not an occurrence of this series.")
        if want not in self.cache:
            self.cache[want] = Occurrence(self.appt, want)
        return self.cache[want]


class Appointment(Item):
    MessageClass = "IPM.Appointment"
    cancel_moves_item = False  # does Send() of a cancellation file the appointment in Deleted Items itself?

    def __init__(self, subject="", start=None, end=None, all_day=False, recurring=False, meeting=0, **kw):
        super().__init__(**kw)
        self.Subject = subject
        self.set_local("Start", start)
        self.set_local("End", end)
        self.AllDayEvent = all_day
        self.IsRecurring = recurring
        self._pattern = None
        self.MeetingStatus = meeting
        self.ResponseStatus = 0
        self.Location = ""
        self.Organizer = ""
        self.BusyStatus = 2
        self.Sensitivity = 0
        self.ReminderSet = False
        self.ReminderMinutesBeforeStart = 15
        self.responses = []
        self.sent_meetings = 0

    def GetRecurrencePattern(self):
        if self._pattern is None:
            self._pattern = RecurrencePattern(self)
            if self.IsRecurring:  # test data flagged as recurring without a pattern: weekly on its own day
                self._pattern.RecurrenceType = 1
                self._pattern.DayOfWeekMask = 1 << ((self._raw["Start"].weekday() + 1) % 7)
            self.IsRecurring = True  # the trap: asking for the pattern makes an appointment recurring
        return self._pattern

    def _kind(self, kind):
        return "; ".join(r.Address for r in self.Recipients._items if r.Type == kind)

    def _set_kind(self, kind, text):
        self.Recipients._items = [r for r in self.Recipients._items if r.Type != kind]
        for a in [x.strip() for x in text.split(";") if x.strip()]:
            self.Recipients._items.append(Recipient(a, type_=kind))
        self.app.log.append(("attendees", kind, text))

    RequiredAttendees = property(lambda self: self._kind(1), lambda self, v: self._set_kind(1, v))
    OptionalAttendees = property(lambda self: self._kind(2), lambda self, v: self._set_kind(2, v))

    def Send(self):
        if not self.Recipients._items:
            raise ComError(BLOCKED_PROPERTY, "This item has no recipients; it cannot be sent.")
        self.sent_meetings += 1
        self.app.log.append(("send_appt", self.Subject, self.MeetingStatus))
        if self.MeetingStatus == 5 and self.cancel_moves_item and self.Parent is not None:
            self.Move(self.app.Session.stores[0].defaults[3])

    def Respond(self, code, no_ui):
        resp = Response(self, code)
        self.responses.append(resp)
        if code in (2, 3):
            # Outlook makes a copy with a new EntryID and removes the original.
            self.EntryID = new_entry_id()
        elif code == 4:
            self.Parent._items.remove(self)
        return resp


class Occurrence(Appointment):
    """One occurrence of a series. Changing and saving it makes an exception; it never becomes an item
    of the calendar folder itself."""

    def __init__(self, master, start):
        duration = master._raw["End"] - master._raw["Start"]
        super().__init__(master.Subject, start, start + duration, master.AllDayEvent, True, master.MeetingStatus)
        self.master, self.app, self.Parent, self.EntryID = master, master.app, master.Parent, master.EntryID
        for name in ("Location", "Organizer", "BusyStatus", "Sensitivity", "ReminderSet", "ResponseStatus",
                     "Categories", "Body", "ReminderMinutesBeforeStart"):
            setattr(self, name, getattr(master, name))
        self.Recipients = Recipients(list(master.Recipients._items))

    def GetRecurrencePattern(self):
        return self.master.GetRecurrencePattern()

    def Save(self):
        self.saved += 1

    def Move(self, folder):
        raise AssertionError("an occurrence cannot be moved")

    def Respond(self, code, no_ui):
        resp = Response(self, code)
        self.responses.append(resp)
        return resp

    def Send(self):
        super().Send()
        self.master.__dict__.setdefault("occurrence_sends", []).append((self._raw["Start"], self.MeetingStatus))


class Response:
    def __init__(self, appt, code):
        self.appt, self.code = appt, code
        self.Body = ""
        self.sent = self.discarded = False

    def Send(self):
        self.sent = True

    def Close(self, mode):
        self.discarded = mode == 1


CONTACT_STRINGS = (
    "FullName", "FirstName", "MiddleName", "LastName", "NickName", "CompanyName", "JobTitle", "Department",
    "OfficeLocation", "MobileTelephoneNumber", "BusinessTelephoneNumber", "HomeTelephoneNumber", "Email1Address",
    "Email2Address", "Email3Address", "HomeAddressStreet", "HomeAddressCity", "HomeAddressState",
    "HomeAddressPostalCode", "HomeAddressCountry", "BusinessAddressStreet", "BusinessAddressCity",
    "BusinessAddressState", "BusinessAddressPostalCode", "BusinessAddressCountry", "WebPage",
)


class Contact(Item):
    MessageClass = "IPM.Contact"

    def __init__(self, **kw):
        for name in CONTACT_STRINGS:
            setattr(self, name, "")
        self.set_local("Birthday", NO_DATE)
        super().__init__(**kw)

    def table_value(self, col):
        return getattr(self, col)


class DistList(Contact):
    MessageClass = "IPM.DistList"

    def __init__(self, DLName=None, **kw):
        super().__init__(**kw)
        self.DLName = DLName if DLName is not None else self.FullName
        self.members = []

    @property
    def MemberCount(self):
        return len(self.members)

    def GetMember(self, i):
        return self.members[i - 1]

    def AddMember(self, recipient):
        if not recipient.Resolved:
            raise ComError(BLOCKED_PROPERTY, "The recipient is not resolved.")
        self.members.append(recipient)
        self.app.log.append(("addmember", self.DLName, recipient.Address))

    def RemoveMember(self, recipient):
        for m in self.members:
            if m.Address.casefold() == recipient.Address.casefold():
                self.members.remove(m)
                self.app.log.append(("removemember", self.DLName, recipient.Address))
                return
        raise ComError(BLOCKED_PROPERTY, "Not a member.")


class Task(Item):
    MessageClass = "IPM.Task"

    def __init__(self, subject="", due=None, start=None, complete=False, importance=1, **kw):
        super().__init__(**kw)
        self.Subject = subject
        self.set_local("DueDate", due or NO_DATE)
        self.set_local("StartDate", start or NO_DATE)
        self.set_local("DateCompleted", NO_DATE)
        self.set_local("ReminderTime", NO_DATE)
        self.Importance = importance
        self.ReminderSet = False
        self._complete = False
        if complete:
            self.Complete = True

    @property
    def Complete(self):
        return self._complete

    @Complete.setter
    def Complete(self, value):
        self._complete = bool(value)
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        self.set_local("DateCompleted", today if value else NO_DATE)

    def table_value(self, col):
        return getattr(self, col)


class Note(Item):
    """A sticky note: its Subject is the first line of its Body and can't be set."""

    MessageClass = "IPM.StickyNote"

    @property
    def Subject(self):
        return re.split(r"\r\n|\n", self.Body, maxsplit=1)[0]

    def Save(self):
        self.set_local("LastModificationTime", datetime.now().replace(microsecond=0))
        super().Save()


class AddressEntry:
    def __init__(self, name, smtp="", title="", department="", office="", phone="", is_list=False):
        self.Name, self.Address = name, "/o=Org/ou=Exchange/cn=" + name
        self._user = None if is_list else ExchangeUser(name, smtp, title, department, office, phone)
        self.PropertyAccessor = Accessor({})

    def GetExchangeUser(self):
        return self._user


class ExchangeUser:
    def __init__(self, name, smtp, title, department, office, phone):
        self.Name, self.PrimarySmtpAddress, self.JobTitle = name, smtp, title
        self.Department, self.OfficeLocation, self.BusinessTelephoneNumber = department, office, phone


class AddressEntries:
    def __init__(self, entries):
        self._entries, self._pos = entries, 0

    def GetFirst(self):
        self._pos = 0
        return self.GetNext()

    def GetNext(self):
        if self._pos >= len(self._entries):
            return None
        self._pos += 1
        return self._entries[self._pos - 1]


class AddressList:
    def __init__(self, entries):
        self.AddressEntries = AddressEntries(entries)


class Category:
    def __init__(self, name, color=0):
        self.Name, self.Color = name, color


class Categories(Coll):
    def Add(self, name, color=0, shortcut=0):
        if any(c.Name.casefold() == name.casefold() for c in self._items):
            raise ComError(BLOCKED_PROPERTY, "A category with this name already exists.")
        c = Category(name, color)
        self._items.append(c)
        return c


_DASL = re.compile(
    r"""^"urn:schemas:httpmail:datereceived" > '(?P<since>\d{4}-\d\d-\d\d \d\d:\d\d)'$"""
    r"""|^"urn:schemas:httpmail:datereceived" < '(?P<until>\d{4}-\d\d-\d\d \d\d:\d\d)'$"""
    rf"""|^"{re.escape(PR_CONVERSATION_TOPIC)}" = '(?P<topic>(?:[^']|'')*)'$"""
    r"""|^"urn:schemas:httpmail:read" = 0$"""
    r"""|^"urn:schemas:httpmail:textdescription" LIKE '%(?P<like>(?:[^']|'')*)%'$"""
    r"""|^\((?P<classes>.*)\)$"""
)
_CLASS_LIKE = re.compile(rf"""^"{re.escape(PR_MESSAGE_CLASS)}" LIKE '(?P<prefix>[A-Za-z.]+)%'$""")


def dasl_predicate(flt: str):
    """The fake's reading of the backend's DASL filters."""
    if not flt:
        return lambda it: True
    assert flt.startswith("@SQL="), flt
    tests = []
    for part in flt[5:].split(" AND "):
        m = _DASL.match(part)
        assert m, f"filter the fake does not know: {part!r}"
        if m.group("since"):
            since = datetime.strptime(m.group("since"), "%Y-%m-%d %H:%M").replace(tzinfo=UTC)
            tests.append(lambda it, since=since: "ReceivedTime" in it.__dict__.get("_raw", {})
                         and it.ReceivedTime.year < 4501 and it.true_utc("ReceivedTime") > since)
        elif m.group("until"):
            until = datetime.strptime(m.group("until"), "%Y-%m-%d %H:%M").replace(tzinfo=UTC)
            tests.append(lambda it, until=until: "ReceivedTime" in it.__dict__.get("_raw", {})
                         and it.ReceivedTime.year < 4501 and it.true_utc("ReceivedTime") < until)
        elif m.group("topic") is not None:
            topic = m.group("topic").replace("''", "'").casefold()
            tests.append(lambda it, t=topic: getattr(it, "ConversationTopic", "").casefold() == t)
        elif m.group("like") is not None:
            needle = m.group("like").replace("''", "'").casefold()
            tests.append(lambda it, n=needle: n in (getattr(it, "Body", "") or "").casefold())
        elif m.group("classes") is not None:
            prefixes = []
            for alt in m.group("classes").split(" OR "):
                cm = _CLASS_LIKE.match(alt)
                assert cm, f"class filter the fake does not know: {alt!r}"
                prefixes.append(cm.group("prefix"))
            tests.append(lambda it, p=tuple(prefixes): it.MessageClass.startswith(p))
        else:
            tests.append(lambda it: it.UnRead)
    return lambda it: all(t(it) for t in tests)


class Table:
    def __init__(self, items, flt):
        self._items = [it for it in items if dasl_predicate(flt)(it)]
        self.filter = flt
        self.Columns = Columns()
        self._pos = 0

    def Sort(self, col, descending):
        self._items.sort(key=lambda it: getattr(it, col), reverse=descending)

    def GetRowCount(self):
        return len(self._items)

    @property
    def EndOfTable(self):
        return self._pos >= len(self._items)

    def GetArray(self, n):
        rows = self._items[self._pos:self._pos + n]
        self._pos += len(rows)
        return tuple(tuple(it.table_value(c) for c in self.Columns.names) for it in rows)


class Columns:
    def __init__(self):
        self.names = ["EntryID", "Subject"]

    def RemoveAll(self):
        self.names = []

    def Add(self, name):
        if name in BLOCKED_COLUMNS:
            raise ComError(BLOCKED_PROPERTY, f"The property \"{name}\" does not support this operation.")
        self.names.append(name)


class Items(Coll):
    def __init__(self, folder):
        super().__init__()
        self._items = folder._items
        self.folder = folder
        self.sorted_by = None
        self.IncludeRecurrences = False
        self.restriction = None

    def Add(self, kind):
        cls = {ITEM_MAIL: Mail, ITEM_APPOINTMENT: Appointment, ITEM_CONTACT: Contact, ITEM_TASK: Task,
               ITEM_NOTE: Note, ITEM_DISTLIST: DistList}[kind]
        it = cls()
        it.app = self.folder.app
        it.Parent = self.folder  # not in the folder until saved
        return it

    def Sort(self, key):
        self.sorted_by = key

    def Restrict(self, flt):
        self.restriction = flt
        self.folder.app.log.append(("restrict", flt))
        items = list(self._items)
        if self.sorted_by == "[Start]":
            items = sorted((it for it in items if isinstance(it, Appointment)), key=lambda it: it.Start)
        return Restricted(items)


class Restricted:
    def __init__(self, items):
        self._items, self._pos = items, 0

    def GetFirst(self):
        self._pos = 0
        return self.GetNext()

    def GetNext(self):
        if self._pos >= len(self._items):
            return None
        self._pos += 1
        return self._items[self._pos - 1]


class Folder:
    def __init__(self, app, store, name, kind=ITEM_MAIL, hidden=False):
        self.app, self.Store, self.Name = app, store, name
        self.EntryID = new_entry_id()
        self.DefaultItemType = kind
        self._items = []
        self._subs = []
        self.Parent = None
        self.PropertyAccessor = Accessor({PR_HIDDEN: True} if hidden else {})

    @property
    def Folders(self):
        folder = self

        class Subs(Coll):
            def Add(self, name):
                return folder.add(name)

        return Subs(self._subs)

    @property
    def Items(self):
        return Items(self)

    def Delete(self):
        raise AssertionError("Folder.Delete() must never be called")

    def MoveTo(self, dest):
        assert dest.Store is self.Store, "the fake moves folders inside one store only"
        if any(f.Name.casefold() == self.Name.casefold() for f in dest._subs):
            raise ComError(BLOCKED_PROPERTY, "A folder with that name already exists.")
        self.Parent._subs.remove(self)
        dest._subs.append(self)
        self.Parent = dest
        self.app.log.append(("moveto", self.Name, dest.Name))

    @property
    def UnReadItemCount(self):
        return sum(1 for it in self._items if getattr(it, "UnRead", False))

    def GetTable(self, flt="", contents=0):
        self.app.log.append(("table", self.Name, flt))
        return Table(self._items, flt)

    def add(self, name, kind=ITEM_MAIL, hidden=False):
        f = Folder(self.app, self.Store, name, kind, hidden)
        f.Parent = self
        self._subs.append(f)
        return f

    def put(self, *items):
        for it in items:
            it.app = self.app
            it.Parent = self
            self._items.append(it)
        return self


class Store:
    def __init__(self, app, name, store_id, exchange_type=0):
        self.app, self.DisplayName, self.StoreID = app, name, store_id
        self.ExchangeStoreType = exchange_type
        self.root = Folder(app, self, "Top of Information Store")
        self.defaults = {}
        self.PropertyAccessor = Accessor({})

    def GetRootFolder(self):
        return self.root

    def GetDefaultFolder(self, const):
        if const not in self.defaults:
            raise ComError(-2147221233, "no such default folder")
        return self.defaults[const]


class Account:
    def __init__(self, smtp, store, account_type=0):
        self.SmtpAddress, self.DisplayName = smtp, smtp
        self.DeliveryStore = store
        self.AccountType = account_type


class Namespace:
    def __init__(self, app):
        self.app = app
        self.stores, self.accounts = [], []
        self.default_calendar = None
        self.default_contacts = None
        self.DefaultStore = None
        self.send_receive = 0
        self.Categories = Categories()

    def SendAndReceive(self, show_progress):
        self.send_receive += 1

    @property
    def Stores(self):
        return Coll(self.stores)

    @property
    def Accounts(self):
        return Coll(self.accounts)

    def GetDefaultFolder(self, const):
        if const == 9 and self.default_calendar:
            return self.default_calendar
        if const == 10 and self.default_contacts:
            return self.default_contacts
        if const in (12, 13) and self.DefaultStore is not None:
            return self.DefaultStore.GetDefaultFolder(const)
        raise ComError(-2147221233, "no default folder")

    def CreateRecipient(self, text):
        return Recipient(text, app=self.app)

    def GetGlobalAddressList(self):
        if self.app.gal is None:
            raise ComError(BLOCKED_PROPERTY, "There is no Exchange address list.")
        return AddressList(self.app.gal)

    def _folders(self, folder):
        for f in folder._subs:
            yield f
            yield from self._folders(f)

    def GetFolderFromID(self, entry, store_id):
        for st in self.stores:
            if st.StoreID == store_id:
                for f in self._folders(st.root):
                    if f.EntryID == entry:
                        return f
        raise ComError(NOT_FOUND, "The operation failed. An object could not be found.")

    def GetItemFromID(self, entry, store_id):
        for st in self.stores:
            if st.StoreID == store_id:
                for f in self._folders(st.root):
                    for it in f._items:
                        if it.EntryID == entry:
                            return it
        raise ComError(NOT_FOUND, "The operation failed. An object could not be found.")


class App:
    def __init__(self):
        self.Session = Namespace(self)
        self.log = []
        self.default_drafts = None
        self.gal = None  # list of AddressEntry; None = no Exchange account
        self.free_busy = {}  # address -> free/busy string from midnight of the day asked for
        self.Explorers = Coll([object()])  # one Outlook window open

    def CreateItem(self, kind):
        it = Items(self.default_drafts).Add(kind)
        return it

    def store(self, name, store_id, smtp=None, exchange_type=0, account_type=None):
        st = Store(self, name, store_id, exchange_type)
        self.Session.stores.append(st)
        if self.Session.DefaultStore is None and exchange_type != EXCHANGE_PUBLIC_FOLDERS:
            self.Session.DefaultStore = st
        if smtp:
            kind = (0 if exchange_type != 3 else 1) if account_type is None else account_type
            self.Session.accounts.append(Account(smtp, st, kind))
        return st
