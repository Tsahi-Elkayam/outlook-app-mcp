"""The Windows backend against a fake Outlook object model (tests/windows/fakecom.py)."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

import pytest
from fakecom import App, Appointment, Contact, DistList, Mail, Recipient, to_outlook

from outlook_app_mcp.common import OutlookError, iso
from outlook_app_mcp.windows import WindowsOutlook
from outlook_app_mcp.windows.com import (
    FOLDER_CALENDAR,
    FOLDER_CONTACTS,
    FOLDER_DELETED,
    FOLDER_DRAFTS,
    FOLDER_INBOX,
    FOLDER_JUNK,
    FOLDER_OUTBOX,
    FOLDER_SENT,
    ITEM_APPOINTMENT,
    ITEM_CONTACT,
    ComError,
    from_com,
    split_id,
    store_key,
    to_com,
)
from outlook_app_mcp.windows.mail import MAIL_ONLY, prepend_html

ME = "me@work.example.com"
GMAIL = "me@gmail.example"


def local(*args) -> datetime:
    """A naive local time, as Outlook's object model hands them out."""
    return datetime(*args)


def build():
    app = App()
    a = app.store("Work", "STORE-A", smtp=ME)
    inbox = a.root.add("Inbox")
    projects = inbox.add("Projects")
    sent = a.root.add("Sent Items")
    drafts = a.root.add("Drafts")
    deleted = a.root.add("Deleted Items")
    old = deleted.add("Old")
    junk = a.root.add("Junk Email")
    outbox = a.root.add("Outbox")
    a.root.add("Sync Issues", hidden=True)
    cal = a.root.add("Calendar", kind=ITEM_APPOINTMENT)
    birthdays = cal.add("Birthdays", kind=ITEM_APPOINTMENT)
    contacts = a.root.add("Contacts", kind=ITEM_CONTACT)
    a.defaults = {FOLDER_INBOX: inbox, FOLDER_SENT: sent, FOLDER_DRAFTS: drafts, FOLDER_DELETED: deleted,
                  FOLDER_JUNK: junk, FOLDER_OUTBOX: outbox}
    app.Session.default_calendar = cal
    app.Session.default_contacts = contacts
    app.default_drafts = drafts

    g = app.store("me@gmail.example", "STORE-B", smtp=GMAIL, exchange_type=3)
    g_inbox = g.root.add("Inbox")
    g_drafts = g.root.add("Drafts")
    g_trash = g.root.add("Trash")
    g.defaults = {FOLDER_INBOX: g_inbox, FOLDER_DRAFTS: g_drafts, FOLDER_DELETED: g_trash}

    # An archive data file with its own Inbox, as AutoArchive makes them: no account, no Deleted Items.
    archive = app.store("Archive", "STORE-C", exchange_type=3)
    old_mail = archive.root.add("Old mail")
    archive.defaults = {FOLDER_INBOX: archive.root.add("Inbox")}

    # Public Folders: never walked (on a real server, thousands of online folders).
    public = app.store("Public Folders - me", "STORE-P", exchange_type=2)
    public.root.add("All Public Folders").add("Team calendar", kind=ITEM_APPOINTMENT)
    public.root.add("Shared contacts", kind=ITEM_CONTACT)

    # Hidden system contact folders (auto-complete cache): not contacts the user keeps.
    cache = a.root.add("Recipient Cache", kind=ITEM_CONTACT, hidden=True)

    now = datetime.now()
    m1 = Mail("older", ("Rina", "rina@x.co"), now - timedelta(days=3))
    m2 = Mail("שלום", ("אבי", "/o=ExchangeLabs/ou=Exchange/cn=Recipients/cn=avi"), now - timedelta(hours=2),
              unread=True, flag=2, sender_smtp="avi-c@x.co")
    m3 = Mail("report", ("Dana", "dana@x.co"), now - timedelta(days=40), Body="the quarterly numbers")
    inbox.put(m1, m2, m3)
    m2.Recipients._items = [Recipient("lee@x.co", "Lee", 1), Recipient("c@x.co", "Cee", 2)]
    m3.Attachments._items = []
    from fakecom import Attachment

    m3.Attachments._items = [Attachment("a.pdf", b"pdf!", "application/pdf"), Attachment("x/../evil:name", b"e")]
    gone = Mail("old trash", ("x", "x@x.co"), now - timedelta(days=1))
    older_gone = Mail("older trash", ("x", "x@x.co"), now - timedelta(days=5))
    deleted.put(gone)
    old.put(older_gone)
    # Deleted Items also holds other kinds of items; mail tools must leave them alone.
    gone_appt = Appointment("deleted appt", local(2026, 11, 3, 9), local(2026, 11, 3, 10))
    gone_contact = Contact(FullName="Gone Contact")
    deleted.put(gone_appt, gone_contact)
    draft = Mail("draft", ("", ""), None)
    draft.Recipients._items = [Recipient("to@x.co", "", 1)]
    empty_draft = Mail("empty", ("", ""), None)
    drafts.put(draft, empty_draft)
    g_mail = Mail("gmail msg", ("Bob", "bob@y.co"), now - timedelta(hours=5))
    g_inbox.put(g_mail)
    old_mail.put(Mail("archived", ("z", "z@z.co"), now - timedelta(days=400)))

    one_off = Appointment("One-off", local(2026, 11, 5, 9), local(2026, 11, 5, 10))
    weekly = Appointment("Weekly", local(2026, 11, 9, 10), local(2026, 11, 9, 11), recurring=True)
    holiday = Appointment("Holiday", local(2026, 11, 10), local(2026, 11, 11), all_day=True)
    before = Appointment("Before", local(2026, 11, 1, 9), local(2026, 11, 1, 10))
    meeting = Appointment("Meeting", local(2026, 11, 6, 9), local(2026, 11, 6, 10), meeting=3)
    meeting.Recipients._items = [Recipient("boss@x.co", "Boss", 1, 1), Recipient(ME, "Me", 1, 5)]
    cal.put(one_off, weekly, holiday, before, meeting)
    birthdays.put(Appointment("Birthday", local(2026, 11, 7), local(2026, 11, 8), all_day=True))

    rina = Contact(FullName="Rina K", FirstName="Rina", LastName="K", Email1Address="rina@x.co",
                    MobileTelephoneNumber="050-123-4567")
    avi = Contact(FullName="אבי", FirstName="אבי", Email1Address="avi@home.co")
    team = DistList(FullName="Team rina")
    contacts.put(rina, avi, team)
    cache.put(Contact(FullName="rina cached", Email1Address="rina@x.co"))

    o = WindowsOutlook(app=app, com_error=ComError)
    refs = SimpleNamespace(**{k: v for k, v in locals().items() if k not in ("app", "o", "refs")})
    return o, app, refs


@pytest.fixture
def world():
    return build()


def wid(store, item) -> str:
    return f"{store_key(store.StoreID)}.{item.EntryID}"


# -- ids and dates ---------------------------------------------------------------


def test_ids_are_store_key_and_entry_id(world):
    o, _, r = world
    f = next(f for f in o.list_folders() if f["role"] == "inbox" and f["account"] == ME)
    assert f["id"] == wid(r.a, r.inbox)
    assert split_id(f["id"], "x") == (store_key("STORE-A"), r.inbox.EntryID)
    for bad in ("55", "abc.def", ""):
        with pytest.raises(OutlookError):
            o.read_message(bad)


def test_unknown_store_and_missing_item(world):
    o, _, r = world
    with pytest.raises(OutlookError, match="not open in Outlook"):
        o.read_message("0123456789." + r.m1.EntryID)
    with pytest.raises(OutlookError, match="could not find"):
        o.read_message(f"{store_key('STORE-A')}.00000000DEADBEEF")


def test_com_dates_are_local_and_no_date_is_none():
    marked_utc = datetime(2026, 11, 5, 9, 0, tzinfo=UTC)  # how pywin32 labels a local 09:00
    assert from_com(marked_utc).hour == 9 and from_com(marked_utc).tzinfo is not None
    assert from_com(datetime(4501, 1, 1)) is None and from_com(None) is None
    aware = datetime(2026, 11, 5, 7, 0, tzinfo=UTC)
    wall = aware.astimezone().replace(tzinfo=None)
    # What Outlook ends up with after pywin32's astimezone(UTC): the local wall-clock time.
    assert to_outlook(to_com(aware)) == wall
    assert to_outlook(to_com(datetime(2026, 11, 5, 14, 0))) == datetime(2026, 11, 5, 14, 0)
    # A naive time handed over as is would be shifted by the UTC offset.
    naive = datetime(2026, 7, 5, 14, 0)
    assert to_outlook(naive) == naive.astimezone(UTC).replace(tzinfo=None)


# -- folders ---------------------------------------------------------------------


def test_folders_roles_paths_accounts(world):
    o, _, _ = world
    fs = {(f["account"], f["path"]): f for f in o.list_folders()}
    assert fs[(ME, "Inbox")]["role"] == "inbox" and fs[(ME, "Inbox/Projects")]["role"] is None
    assert fs[(ME, "Deleted Items")]["role"] == "deleted" and fs[(ME, "Junk Email")]["role"] == "junk"
    assert fs[(ME, "Inbox")]["unread"] == 1 and fs[(ME, "Inbox")]["total"] == 3
    assert fs[(GMAIL, "Trash")]["role"] == "deleted" and fs[(GMAIL, "Trash")]["kind"] == "other"
    assert fs[("Archive", "Old mail")]["account_name"] == "Archive"
    paths = {p for _, p in fs}
    assert "Calendar" not in paths and "Contacts" not in paths and "Sync Issues" not in paths
    assert [f["path"] for f in o.list_folders("gmail")] == ["Inbox", "Drafts", "Trash"]
    with pytest.raises(OutlookError, match="Accounts:"):
        o.list_folders("nobody")


# -- reading ---------------------------------------------------------------------


def test_list_newest_first_with_exact_count(world):
    o, app, r = world
    res = o.list_messages(account="work", limit=2)
    assert [m["subject"] for m in res["messages"]] == ["שלום", "older"]
    assert res["matched"] == 3 and res["returned"] == 2 and res["has_more"] is True
    m = res["messages"][0]
    assert m["from"] == {"name": "אבי", "address": "avi-c@x.co"}  # SMTP, not the X.500 address
    assert m["is_read"] is False and m["flag"] == "flagged" and m["folder"] == "Inbox" and m["account"] == ME
    assert m["id"] == wid(r.a, r.m2) and m["folder_id"] == wid(r.a, r.inbox)
    assert app.log[-1] == ("table", "Inbox", "@SQL=" + MAIL_ONLY)


def test_list_filters_build_dasl(world):
    o, app, _ = world
    res = o.list_messages(account="work", since_days=1)
    assert [m["subject"] for m in res["messages"]] == ["שלום"]
    assert app.log[-1][2].startswith('@SQL="urn:schemas:httpmail:datereceived" > \'')
    res = o.list_messages(account="work", unread_only=True)
    assert res["matched"] == 1 and app.log[-1][2] == '@SQL="urn:schemas:httpmail:read" = 0 AND ' + MAIL_ONLY
    with pytest.raises(OutlookError):
        o.list_messages(since_days=0)
    # Two accounts have an inbox: Outlook's default store wins.
    assert o.list_messages()["folder"]["account"] == ME


def test_search_subject_sender_and_body(world):
    o, app, _r = world
    assert [m["subject"] for m in o.search("AVI-C")["messages"]] == ["שלום"]
    assert o.search("report")["matched"] == 1  # 40 days old: inside the default 90
    assert o.search("report", since_days=30)["matched"] == 0
    res = o.search("quarterly", include_body=True, since_days=None)
    assert [m["subject"] for m in res["messages"]] == ["report"]
    like = [e for e in app.log if e[0] == "table" and "textdescription" in e[2]]
    assert like and "LIKE '%quarterly%' AND (" in like[0][2]
    o.search("it's", include_body=True)
    assert any("LIKE '%it''s%'" in e[2] for e in app.log if e[0] == "table")
    with pytest.raises(OutlookError):
        o.search("  ")


def test_read_message_does_not_mark_read(world):
    o, _, r = world
    m = o.read_message(wid(r.a, r.m2), max_chars=3)
    assert m["subject"] == "שלום" and m["to"] == [{"name": "Lee", "address": "lee@x.co"}]
    assert m["cc"] == [{"name": "Cee", "address": "c@x.co"}] and m["bcc"] == []
    assert m["is_read"] is False and r.m2.UnRead is True and r.m2.saved == 0
    assert m["folder_id"] == wid(r.a, r.inbox) and m["account"] == ME and m["flag"] == "flagged"
    m3 = o.read_message(wid(r.a, r.m3))
    assert m3["attachments"][0] == {"index": 1, "name": "a.pdf", "size": 4, "content_type": "application/pdf"}
    assert m3["body"] == "the quarterly numbers" and m3["body_truncated"] is False


# -- changing ----------------------------------------------------------------------


def test_mark_saves_each_item(world):
    o, _, r = world
    ids = [wid(r.a, r.m1), wid(r.a, r.m2)]
    assert o.mark_messages(ids, "unread") == {"action": "unread", "changed": 2}
    assert r.m1.UnRead is True and r.m1.saved == 1
    o.mark_messages(ids[:1], "complete")
    assert r.m1.FlagStatus == 1
    with pytest.raises(OutlookError, match="action"):
        o.mark_messages(ids, "delete")


def test_move_reports_new_ids_across_stores(world):
    o, _, r = world
    same = o.move_messages([wid(r.a, r.m1)], wid(r.a, r.projects))
    assert same["new_ids"] == {wid(r.a, r.m1): wid(r.a, r.m1)} and r.m1.Parent is r.projects
    old_id = wid(r.a, r.m3)
    other = o.move_messages([old_id], wid(r.archive, r.old_mail))
    assert other["new_ids"][old_id] == wid(r.archive, r.m3) != old_id


def test_delete_moves_to_the_stores_deleted_items(world):
    o, _app, r = world
    res = o.delete_messages([wid(r.a, r.m1), wid(r.g, r.g_mail)])
    assert res["deleted"] == 2 and "permanently" in res["note"]
    assert r.m1.Parent is r.deleted and r.g_mail.Parent is r.g_trash


def test_delete_refuses_items_already_deleted(world):
    o, _, r = world
    for item in (r.gone, r.older_gone):
        with pytest.raises(OutlookError, match="permanently"):
            o.delete_messages([wid(r.a, r.m1), wid(r.a, item)])
    assert r.m1.Parent is r.inbox  # nothing moved when one id is refused


def test_draft_goes_to_the_accounts_drafts(world, tmp_path):
    o, _, r = world
    f = tmp_path / "doc.pdf"
    f.write_bytes(b"x")
    res = o.create_draft(["s@x.co"], "נושא", "שלום\n\nתודה", cc=["c@x.co"], attachments=[str(f)], account="work")
    draft = r.drafts._items[-1]
    assert res == {"id": wid(r.a, draft), "status": "draft", "account": ME}
    assert draft.Subject == "נושא" and 'dir="rtl"' in draft.HTMLBody and draft.SendUsingAccount.SmtpAddress == ME
    assert [(x.Address, x.Type) for x in draft.Recipients] == [("s@x.co", 1), ("c@x.co", 2)]
    assert draft.Recipients.resolved and draft.Attachments.Item(1).FileName == "doc.pdf" and not draft.sent


def test_send_message_sends_without_saving_a_copy(world):
    o, app, r = world
    before = len(r.drafts._items)
    res = o.send_message(["s@x.co"], "hi", "body")
    assert res["status"] == "sent" and res["id"] is None and ("send", "hi") in app.log
    assert len(r.drafts._items) == before


def test_compose_checks(world):
    o, _, _ = world
    with pytest.raises(OutlookError, match="recipient"):
        o.send_message([], "s", "b")
    with pytest.raises(OutlookError, match="not an email"):
        o.create_draft(["bad"])
    with pytest.raises(OutlookError, match="data file"):
        o.create_draft(["a@b.co"], account="Archive")


def test_send_draft_only_from_drafts_with_recipients(world):
    o, _app, r = world
    with pytest.raises(OutlookError, match="not in a Drafts"):
        o.send_draft(wid(r.a, r.m1))
    with pytest.raises(OutlookError, match="no recipients"):
        o.send_draft(wid(r.a, r.empty_draft))
    assert o.send_draft(wid(r.a, r.draft))["status"] == "sent" and r.draft.sent


def test_reply_is_a_draft_in_the_originals_account(world):
    o, _, r = world
    res = o.reply_to_message(wid(r.g, r.g_mail), "תודה")
    reply = r.g_drafts._items[-1]
    assert res == {"id": wid(r.g, reply), "status": "draft", "in_reply_to": wid(r.g, r.g_mail), "mode": "reply"}
    assert reply.HTMLBody.index("תודה") < reply.HTMLBody.index("quoted") and not reply.sent
    res = o.forward_message(wid(r.a, r.m1), ["x@y.co"], send=True)
    assert res["status"] == "sent"
    with pytest.raises(OutlookError, match="at least one"):
        o.forward_message(wid(r.a, r.m1), [], send=True)


def test_prepend_html():
    assert prepend_html('<html><body dir="rtl"><p>q</p></body></html>', "<div>x</div>") == (
        '<html><body dir="rtl"><div>x</div><p>q</p></body></html>'
    )
    assert prepend_html("plain", "<b>") == "<b>plain"


def test_folder_create_and_rename(world):
    o, _, r = world
    res = o.create_folder("2026", parent_id=wid(r.a, r.inbox))
    assert res["id"].startswith(store_key("STORE-A")) and r.inbox._subs[-1].Name == "2026"
    o.create_folder("Top", account="gmail")
    assert r.g.root._subs[-1].Name == "Top"
    o.rename_folder(wid(r.a, r.projects), "Proj")
    assert r.projects.Name == "Proj"
    with pytest.raises(OutlookError, match="inbox folder"):
        o.rename_folder(wid(r.a, r.inbox), "x")
    with pytest.raises(OutlookError):
        o.create_folder("a/b", parent_id=wid(r.a, r.inbox))


def test_save_attachment_never_overwrites(world, tmp_path):
    o, _, r = world
    (tmp_path / "a.pdf").write_bytes(b"old")
    res = o.save_attachment(wid(r.a, r.m3), "a.pdf", str(tmp_path))
    assert res["path"].endswith("a (1).pdf") and res["size"] == 4 and (tmp_path / "a.pdf").read_bytes() == b"old"
    res = o.save_attachment(wid(r.a, r.m3), 2, str(tmp_path))
    assert res["name"] == "x/../evil:name" and Path(res["path"]).parent == tmp_path.resolve()
    assert Path(res["path"]).name == "x_.._evil_name"


# -- calendar ------------------------------------------------------------------------


@pytest.fixture
def jet(monkeypatch):
    monkeypatch.setattr(WindowsOutlook, "_jet", staticmethod(lambda dt: dt.strftime("%d/%m/%Y %H:%M")))


def test_calendars(world, jet):
    o, _, _r = world
    cals = o.list_calendars()
    assert [(c["name"], c["is_default"], c["account"]) for c in cals] == [
        ("Calendar", True, ME), ("Birthdays", False, ME)]


def test_events_in_range_sorted_with_all_day_dates(world, jet):
    o, app, r = world
    res = o.list_events("2026-11-04", "2026-11-11")
    got = [(e["subject"], e["start"][:16]) for e in res["events"]]
    assert got == [("One-off", "2026-11-05T09:00"), ("Meeting", "2026-11-06T09:00"), ("Birthday", "2026-11-07"),
                   ("Weekly", "2026-11-09T10:00"), ("Holiday", "2026-11-10")]
    holiday = res["events"][-1]
    assert holiday["end"] == "2026-11-11" and holiday["all_day"] is True
    assert res["events"][3]["recurring"] is True and res["events"][1]["free_busy"] == "busy"
    restricts = [e[1] for e in app.log if e[0] == "restrict"]
    assert restricts[0] == "[Start] < '11/11/2026 00:00' AND [End] > '04/11/2026 00:00'"
    only = o.list_events("2026-11-04", "2026-11-11", calendar_id=wid(r.a, r.birthdays))
    assert [e["subject"] for e in only["events"]] == ["Birthday"]
    with pytest.raises(OutlookError, match="after start"):
        o.list_events("2026-11-10", "2026-11-09")
    with pytest.raises(OutlookError, match="No calendar"):
        o.list_events(calendar_id=wid(r.a, r.inbox))


def test_read_event_and_attendees(world, jet):
    o, _, r = world
    ev = o.read_event(wid(r.a, r.meeting))
    assert ev["is_meeting"] and ev["calendar_id"] == wid(r.a, r.cal)
    assert ev["attendees"][0] == {"name": "Boss", "address": "boss@x.co", "type": "required", "status": "organizer"}
    assert ev["attendees"][1]["status"] == "not responded"


def test_delete_event_moves_to_deleted_items_and_refuses_meetings(world, jet):
    o, _, r = world
    with pytest.raises(OutlookError, match="respond_to_meeting"):
        o.delete_event(wid(r.a, r.meeting))
    assert r.meeting.Parent is r.cal
    res = o.delete_event(wid(r.a, r.one_off))
    assert res["subject"] == "One-off" and r.one_off.Parent is r.deleted and "permanently" in res["note"]


def test_delete_event_never_deletes_for_good(world, jet):
    """Outlook's delete inside Deleted Items is permanent: those are refused, and ids of
    other kinds of items are not events."""
    o, _, r = world
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_event(wid(r.a, r.gone_appt))
    assert r.gone_appt.Parent is r.deleted
    for not_event in (r.gone, r.empty_draft, r.gone_contact):
        with pytest.raises(OutlookError, match="not a calendar event"):
            o.delete_event(wid(r.a, not_event))
    assert r.gone.Parent is r.deleted and r.empty_draft.Parent is r.drafts


def test_recurring_series_needs_whole_series(world, jet):
    o, _, r = world
    sid = wid(r.a, r.weekly)
    with pytest.raises(OutlookError, match="whole_series"):
        o.delete_event(sid)
    with pytest.raises(OutlookError, match="whole_series"):
        o.update_event(sid, subject="x")
    with pytest.raises(OutlookError, match="times of a recurring|can't be changed here"):
        o.update_event(sid, start="2026-11-09T11:00", whole_series=True)
    assert r.weekly.Subject == "Weekly" and r.weekly.Parent is r.cal
    assert o.update_event(sid, location="Room 2", whole_series=True)["location"] == "Room 2"
    o.delete_event(sid, whole_series=True)
    assert r.weekly.Parent is r.deleted


def test_events_in_deleted_items_are_not_changed(world, jet):
    o, _, r = world
    with pytest.raises(OutlookError, match="Deleted Items"):
        o.update_event(wid(r.a, r.gone_appt), subject="x")
    assert r.gone_appt.Subject == "deleted appt"


def test_create_and_update_event(world, jet):
    o, _, r = world
    ev = o.create_event("Plan", "2026-11-12T14:00", "2026-11-12T15:00", location="Room", body="שלום",
                        reminder_minutes=10, free_busy="out of office")
    item = r.cal._items[-1]
    # Outlook holds the wall-clock time the caller asked for, not one shifted by the UTC offset.
    assert item._raw["Start"] == local(2026, 11, 12, 14) and item._raw["End"] == local(2026, 11, 12, 15)
    assert item.BusyStatus == 3 and item.ReminderMinutesBeforeStart == 10
    assert ev["subject"] == "Plan" and ev["reminder_minutes"] == 10 and item.Body == "שלום"
    later = o.update_event(ev["id"], start="2026-11-13T09:00", end="2026-11-13T10:00", remove_reminder=True)
    assert later["start"].startswith("2026-11-13T09:00") and item.ReminderSet is False
    assert item._raw["Start"] == local(2026, 11, 13, 9)
    utc = o.update_event(ev["id"], start="2026-11-13T07:00:00Z", end="2026-11-13T08:00:00Z")
    assert item._raw["Start"] == datetime(2026, 11, 13, 7, tzinfo=UTC).astimezone().replace(tzinfo=None)
    assert utc["start"] == iso(datetime(2026, 11, 13, 7, tzinfo=UTC).astimezone())
    with pytest.raises(OutlookError, match="end before it starts"):
        o.update_event(ev["id"], start="2026-11-14T09:00")
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_event(ev["id"])
    other = o.create_event("B", "2026-11-12", "2026-11-13", calendar_id=wid(r.a, r.birthdays), all_day=True)
    assert other["calendar_id"] == wid(r.a, r.birthdays) and other["start"] == "2026-11-12"
    one_day = o.create_event("C", "2026-11-20", "2026-11-20", all_day=True)
    assert (one_day["start"], one_day["end"]) == ("2026-11-20", "2026-11-21")


def test_update_all_day_event_times(world, jet):
    o, _, r = world
    ev = o.update_event(wid(r.a, r.holiday), end="2026-11-12")
    assert ev["end"] == "2026-11-12"


def test_respond_to_meeting(world, jet):
    o, _, r = world
    old = wid(r.a, r.meeting)
    res = o.respond_to_meeting(old, "accept", "see you")
    resp = r.meeting.responses[-1]
    assert res["organizer_told"] and resp.code == 3 and resp.sent and resp.Body == "see you"
    # Accepting replaces the appointment: the old id is dead, the new one works.
    assert res["id"] == wid(r.a, r.meeting) != old and "new id" in res["note"]
    with pytest.raises(OutlookError, match="could not find"):
        o.read_event(old)
    assert o.read_event(res["id"])["subject"] == "Meeting"
    res = o.respond_to_meeting(res["id"], "decline", send_response=False)
    resp = r.meeting.responses[-1]
    assert resp.code == 4 and not resp.sent and resp.discarded
    assert res["id"] is None and res["organizer_told"] is False and r.meeting not in r.cal._items
    with pytest.raises(OutlookError, match="not a meeting"):
        o.respond_to_meeting(wid(r.a, r.one_off), "accept")
    with pytest.raises(OutlookError, match="response"):
        o.respond_to_meeting(wid(r.a, r.meeting), "maybe")


# -- contacts --------------------------------------------------------------------------


def test_search_contacts(world):
    o, _, r = world
    assert [c["display_name"] for c in o.search_contacts("rina")["contacts"]] == ["Rina K"]  # the list is skipped
    assert [c["first_name"] for c in o.search_contacts("@home")["contacts"]] == ["אבי"]
    assert o.search_contacts("0501234567")["matched"] == 1
    c = o.search_contacts("rina")["contacts"][0]
    assert c["id"] == wid(r.a, r.rina) and c["emails"] == ["rina@x.co"]


def test_create_and_update_contact(world):
    o, _, r = world
    c = o.create_contact(["a@x.co", "b@x.co"], first_name="Dana", company="Acme", note="met at conf")
    item = r.contacts._items[-1]
    assert c["emails"] == ["a@x.co", "b@x.co"] and item.Body == "met at conf" and item.CompanyName == "Acme"
    c = o.update_contact(c["id"], ["z@x.co"], job_title="CTO")
    assert c["emails"] == ["z@x.co"] and item.Email2Address == "" and item.JobTitle == "CTO"
    with pytest.raises(OutlookError, match="at most 3"):
        o.update_contact(c["id"], ["a@x.co", "b@x.co", "c@x.co", "d@x.co"])
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_contact(c["id"])
    with pytest.raises(OutlookError, match="not a contact"):
        o.update_contact(wid(r.a, r.m1), first_name="x")
    with pytest.raises(OutlookError, match="at least"):
        o.create_contact()


def test_no_mail_folders_is_explained():
    app = App()
    app.store("Empty", "STORE-Z")
    o = WindowsOutlook(app=app, com_error=ComError)
    with pytest.raises(OutlookError, match="no mail folders"):
        o.list_folders()


def test_default_folder_constants_match_outlook():
    assert (FOLDER_CALENDAR, FOLDER_CONTACTS, FOLDER_DELETED, FOLDER_DRAFTS) == (9, 10, 3, 16)


# -- review fixes: stores, item kinds, sending ---------------------------------------------


def test_public_folders_hidden_folders_and_archives_stay_out(world, jet):
    o, _, _r = world
    accounts = {f["account"] for f in o.list_folders()}
    assert "Public Folders - me" not in accounts
    assert all(c["name"] != "Team calendar" for c in o.list_calendars())
    assert [c["display_name"] for c in o.search_contacts("rina")["contacts"]] == ["Rina K"]
    # The archive's Inbox doesn't compete with the accounts' inboxes.
    assert o.list_messages()["folder"]["account"] == ME
    # Default searches cover the accounts' own mailboxes, not data files.
    assert o.search("archived", since_days=None)["matched"] == 0
    assert o.search("archived", account="Archive", since_days=None)["matched"] == 1


def test_mail_tools_skip_other_item_kinds(world):
    o, _, r = world
    listed = o.list_messages(folder_id=wid(r.a, r.deleted), limit=50)
    assert [m["subject"] for m in listed["messages"]] == ["old trash"]
    for not_mail in (r.gone_appt, r.gone_contact, r.one_off):
        with pytest.raises(OutlookError, match="not a mail message"):
            o.delete_messages([wid(r.a, not_mail)])
        with pytest.raises(OutlookError, match="not a mail message"):
            o.read_message(wid(r.a, not_mail))


def test_delete_refuses_stores_without_deleted_items(world):
    o, _, r = world
    with pytest.raises(OutlookError, match="permanently|Deleted Items"):
        o.delete_messages([wid(r.archive, r.old_mail._items[0])])
    assert r.old_mail._items[0].Parent is r.old_mail


def test_send_account_is_set_by_reference_and_checked(world, monkeypatch):
    o, _, r = world
    o.create_draft(["s@x.co"], "x", "y", account="gmail")
    draft = r.g_drafts._items[-1]
    assert draft.SendUsingAccount.SmtpAddress == GMAIL and draft._oleobj_.calls == [(64209, 8)]
    # If Outlook doesn't take the account, nothing is saved or sent from the wrong one.
    import fakecom

    monkeypatch.setattr(fakecom.OleObj, "Invoke", lambda self, *a: None)
    before = (len(r.g_drafts._items), len(r.drafts._items))
    with pytest.raises(OutlookError, match="did not accept"):
        o.send_message(["s@x.co"], "x", "y", account="gmail")
    with pytest.raises(OutlookError, match="did not accept"):
        o.create_draft(["s@x.co"], "x", "y", account="gmail")
    assert (len(r.g_drafts._items), len(r.drafts._items)) == before


def test_unresolved_recipient_is_not_sent(world, monkeypatch):
    o, app, _ = world
    import fakecom

    def resolve(self):
        for x in self._items:
            x.Resolved = x.Address != "nobody@nowhere.example"
        return all(x.Resolved for x in self._items)

    monkeypatch.setattr(fakecom.Recipients, "ResolveAll", resolve)
    with pytest.raises(OutlookError, match="nobody@nowhere.example"):
        o.send_message(["ok@x.co", "nobody@nowhere.example"], "hi", "body")
    assert ("send", "hi") not in app.log


def test_send_without_an_outlook_window_starts_send_receive(world):
    o, app, _ = world
    assert "note" not in o.send_message(["s@x.co"], "hi", "body")
    app.Explorers = fakecom_coll([])
    res = o.send_message(["s@x.co"], "hi2", "body")
    assert "Outbox" in res["note"] and app.Session.send_receive == 1


def fakecom_coll(items):
    from fakecom import Coll

    return Coll(items)


def test_search_has_more_and_literal_percent(world):
    o, app, _ = world
    res = o.search("x.co", limit=1, since_days=None)
    assert res["returned"] == 1 and res["has_more"] is True
    o.search("50%", include_body=True)
    assert any("LIKE '%50%'" in e[2] for e in app.log if e[0] == "table" and "textdescription" in e[2])
