import pytest
from conftest import rows

from outlook_app_mcp.mac.osa import FS, GS, RS, US, OutlookError

CALS = rows((7, "me@work.example.com", "Calendar", "true"), (8, "me@work.example.com", "Birthdays", "false"))

WEEKLY = "\r\n".join(  # noqa: FLY002
    [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "BEGIN:VEVENT",
        "UID:weekly-1",
        "DTSTART:20261102T100000",
        "DTEND:20261102T110000",
        "RRULE:FREQ=WEEKLY;BYDAY=MO",
        "EXDATE:20261116T100000",
        "SUMMARY:Weekly",
        "END:VEVENT",
        "END:VCALENDAR",
    ]
)


def event_row(kind, eid, cid, subject, start, end, ics=""):
    return (kind, eid, cid, subject, start, end, "Room 1", "false", "rina@x", "«constant ****eSBu»", ics)


def test_list_events_expands_series_and_sorts(make):
    ev = rows(
        event_row("single", 1, 7, "One-off", "2026-11-5-32400", "2026-11-5-36000"),
        event_row("series", 2, 7, "Weekly", "2026-11-2-36000", "2026-11-2-39600", WEEKLY),
    )
    o, r = make({"read/calendars": CALS, "read/events": ev})
    res = o.list_events("2026-11-04", "2026-11-24")
    args = r.last("read/events")
    assert args[:2] == ["2026-11-4-0", "2026-11-24-0"] and args[2:] == ["7", "8"]
    starts = [(e["subject"], e["start"][:16]) for e in res["events"]]
    # 9th and 23rd; the 16th is excluded by EXDATE; the 2nd is before the range
    assert starts == [("One-off", "2026-11-05T09:00"), ("Weekly", "2026-11-09T10:00"), ("Weekly", "2026-11-23T10:00")]
    assert res["events"][1]["recurring"] is True and res["events"][1]["id"] == "2"
    assert res["events"][0]["free_busy"] == "busy" and res["events"][0]["calendar"] == "Calendar"


def test_bad_series_is_reported_not_fatal(make):
    ev = rows(
        event_row("single", 1, 7, "ok", "2026-11-5-32400", "2026-11-5-36000"),
        event_row("series", 2, 7, "broken", "2026-11-2-36000", "2026-11-2-39600", "not ics"),
    )
    o, _ = make({"read/calendars": CALS, "read/events": ev})
    res = o.list_events("2026-11-04", "2026-11-11")
    assert res["count"] == 1 and res["problems"][0]["id"] == "2"


def test_list_events_validation(make):
    o, _ = make({"read/calendars": CALS})
    with pytest.raises(OutlookError, match="after start"):
        o.list_events("2026-11-10", "2026-11-09")
    with pytest.raises(OutlookError, match="No calendar"):
        o.list_events(calendar_id=99)


def event_raw(attendees="", start="2026-11-10-36000", end="2026-11-10-39600", organizer="me@work.example.com",
              calendar="7", recurring="false"):
    hdr = ["5", calendar, "Review", start, end, "Room", "false", recurring, organizer,
           "«constant ****eSBu»", "false", "true", "15", attendees, ""]
    return FS.join(hdr) + RS + "desc"


def test_read_event_with_attendees(make):
    att = GS.join([US.join(["Avi", "avi@x", "«constant ****ATrq»", "«constant ****ASde»"])])
    o, _ = make({"read/event": event_raw(att)})
    ev = o.read_event(5)
    assert ev["attendees"] == [{"name": "Avi", "address": "avi@x", "type": "required", "status": "declined"}]
    assert ev["reminder_minutes"] == 15 and ev["description"] == "desc" and ev["start"].startswith("2026-11-10T10:00")


def test_create_event_args(make):
    o, r = make({"read/calendars": CALS, "write/event": "5", "read/event": event_raw()})
    o.create_event("Review", "2026-11-10T10:00", "2026-11-10T11:00", location="Room", body="שלום", free_busy="out of office", reminder_minutes=15)
    args = r.last("write/event")
    assert args[:4] == ["create", "", "2026-11-10-36000", "2026-11-10-39600"]
    pairs = dict(zip(args[4::2], args[5::2]))
    assert pairs["subject"] == "Review" and pairs["freebusy"] == "oof" and pairs["reminder"] == "15"
    assert 'dir="rtl"' in pairs["html"]
    with pytest.raises(OutlookError, match="after start"):
        o.create_event("x", "2026-11-10T10:00", "2026-11-10T09:00")
    with pytest.raises(OutlookError, match="No calendar"):
        o.create_event("x", "2026-11-10T10:00", "2026-11-10T11:00", calendar_id=99)


def test_update_event_orders_times_safely(make):
    o, r = make({"read/event": event_raw(), "write/event": "5", "read/calendars": CALS})
    o.update_event(5, start="2026-11-12T10:00", end="2026-11-12T11:00")  # later than old end
    assert r.last("write/event")[2:] == ["end", "2026-11-12-39600", "start", "2026-11-12-36000"]
    o.update_event(5, start="2026-11-09T10:00", end="2026-11-09T11:00")  # earlier
    assert r.last("write/event")[2:] == ["start", "2026-11-9-36000", "end", "2026-11-9-39600"]
    o.update_event(5, remove_reminder=True)
    assert r.last("write/event")[2:] == ["reminder", ""]
    with pytest.raises(OutlookError, match="end before it starts"):
        o.update_event(5, start="2026-11-12T10:00")
    with pytest.raises(OutlookError, match="Nothing"):
        o.update_event(5)


def test_delete_event_refuses_meetings(make):
    att = US.join(["Avi", "avi@x", "required attendee type", "accepted"])
    for raw in (event_raw(att), event_raw(organizer="boss@x")):  # synced invitations can list no attendees
        o, r = make({"read/event": raw, "read/calendars": CALS})
        with pytest.raises(OutlookError, match="decline"):
            o.delete_event(5)
        assert not any(c[0] == "write/event" for c in r.calls)
    o, r = make({"read/event": event_raw(organizer="ME@work.example.com"), "write/event": "5",
                 "read/calendars": CALS})
    assert o.delete_event(5)["deleted"] == "5"
    assert r.last("write/event") == ["delete", "5"]
    o, r = make({"read/event": event_raw(organizer=""), "write/event": "5", "read/calendars": CALS})
    assert o.delete_event(5)["deleted"] == "5"


def test_delete_event_only_from_a_calendar(make):
    """An event outside the calendars (in Deleted Items, say) would be deleted for good."""
    o, r = make({"read/event": event_raw(calendar="106"), "read/calendars": CALS})
    for call in (lambda: o.delete_event(5), lambda: o.update_event(5, subject="x")):
        with pytest.raises(OutlookError, match="not in one of your calendars"):
            call()
    assert not any(c[0] == "write/event" for c in r.calls)


def test_recurring_series_needs_whole_series(make):
    o, r = make({"read/event": event_raw(recurring="true"), "read/calendars": CALS, "write/event": "5"})
    for call in (lambda: o.delete_event(5), lambda: o.update_event(5, subject="x")):
        with pytest.raises(OutlookError, match="whole_series"):
            call()
    with pytest.raises(OutlookError, match="can't be changed here"):
        o.update_event(5, start="2026-11-10T11:00", whole_series=True)
    assert not any(c[0] == "write/event" for c in r.calls)
    o.update_event(5, location="Room 2", whole_series=True)
    assert r.last("write/event") == ["update", "5", "location", "Room 2"]
    o.delete_event(5, whole_series=True)
    assert r.last("write/event") == ["delete", "5"]


def test_respond(make):
    o, r = make({"write/respond": "5", "read/event": event_raw(organizer="boss@x.co"),
                 "read/message": OutlookError("no such message")})
    o.respond_to_meeting(5, "decline", "לא אוכל", send_response=False)
    assert r.last("write/respond") == ["event", "5", "", "decline", "לא אוכל", "0"]
    with pytest.raises(OutlookError):
        o.respond_to_meeting(5, "maybe")


# -- contacts ------------------------------------------------------------------

CONTACTS = rows(
    (1, "Rina K", "Rina", "K", "Example Ltd", "CISO", GS.join(["rina@work.example", "s@home.co"]), "050-123-4567", "", ""),
    (2, "אבי ב", "אבי", "ב", "", "", "avi@x.co", "", "03-555-0000", ""),
)


def test_search_contacts_by_name_email_phone(make):
    o, _ = make({"read/contacts": CONTACTS})
    assert [c["id"] for c in o.search_contacts("rina")["contacts"]] == ["1"]
    assert [c["id"] for c in o.search_contacts("אבי")["contacts"]] == ["2"]
    assert [c["id"] for c in o.search_contacts("@home")["contacts"]] == ["1"]
    assert [c["id"] for c in o.search_contacts("0501234567")["contacts"]] == ["1"]
    assert o.search_contacts("123")["matched"] == 0  # short digit runs don't match phones
    assert o.search_contacts("Rina")["contacts"][0]["emails"] == ["rina@work.example", "s@home.co"]


def contact_raw(cid="2", kind="exchange"):
    fields = [cid, "Noa", "Noa"] + [""] * 21 + ["", "n@x.co", "", "135", "me@work.example.com", kind]
    return FS.join(fields) + RS + "a note"


def _contact_reply(args):
    return OutlookError("No contact with id 99") if args[0] == "99" else contact_raw(args[0])


def test_read_contact_has_every_field(make):
    o, _ = make({"read/contact": _contact_reply})
    c = o.read_contact("2")
    assert c["first_name"] == "Noa" and c["emails"] == ["n@x.co"] and c["note"] == "a note"
    assert c["account"] == "me@work.example.com" and not any(k.startswith("_") for k in c)


def test_delete_contact_only_from_exchange(make):
    o, r = make({"read/contact": lambda a: contact_raw(a[0], kind=""), "write/contact": "2"})
    with pytest.raises(OutlookError, match="deletes such items for good"):
        o.delete_contact("2")
    assert not [c for c in r.calls if c[0] == "write/contact"]
    o, r = make({"read/contact": _contact_reply, "write/contact": "2"})
    assert o.delete_contact("2")["deleted"] == "2"
    assert r.last("write/contact") == ["delete", "2"]


def test_create_and_update_contact(make):
    o, r = make({"read/contacts": CONTACTS, "write/contact": "2", "read/contact": _contact_reply})
    o.create_contact(emails=["n@x.co"], first_name="Noa", mobile="050")
    args = r.last("write/contact")
    assert args[:2] == ["create", ""]
    assert dict(zip(args[2::2], args[3::2])) == {"first": "Noa", "mobile": "050", "emails": "n@x.co"}
    o.update_contact(2, company="RG")
    assert r.last("write/contact") == ["update", "2", "company", "RG"]
    with pytest.raises(OutlookError, match="at least"):
        o.create_contact(note="x")
    with pytest.raises(OutlookError, match="Nothing"):
        o.update_contact(2)
    with pytest.raises(OutlookError, match="No contact"):
        o.update_contact(99, company="x")


def test_new_outlook_mode_blocks_calendar_and_contacts(make):
    """New Outlook shows AppleScript an empty local store (no accounts, one calendar).
    Nothing may read or write it as if it were the user's data."""
    o, r = make({"read/folders": "", "read/calendars": CALS, "read/contacts": CONTACTS})
    calls = [
        lambda: o.list_calendars(),
        lambda: o.list_events(),
        lambda: o.read_event(5),
        lambda: o.create_event("x", "2026-11-10T10:00", "2026-11-10T11:00"),
        lambda: o.update_event(5, subject="x"),
        lambda: o.delete_event(5),
        lambda: o.respond_to_meeting(5, "accept"),
        lambda: o.search_contacts("x"),
        lambda: o.create_contact(first_name="x"),
        lambda: o.update_contact(1, company="x"),
    ]
    for call in calls:
        with pytest.raises(OutlookError, match="Legacy Outlook"):
            call()
    assert {c[0] for c in r.calls} == {"read/folders"}


# -- all-day events --------------------------------------------------------------

def test_all_day_dates_read_in_the_events_zone():
    from outlook_app_mcp.common import all_day_dates

    # synced: midnight UTC, reported as 02:00 Israel time
    assert all_day_dates("2026-11-09T02:00+02:00", "2026-11-10T02:00+02:00", "Etc/GMT") == ("2026-11-09", "2026-11-10")
    # summer: midnight UTC is 03:00 in Israel
    assert all_day_dates("2026-08-25T03:00+03:00", "2026-08-26T03:00+03:00", "Etc/GMT") == ("2026-08-25", "2026-08-26")
    # switched to all-day in Outlook: keeps its hours, local zone, end is exclusive
    assert all_day_dates("2026-11-21T14:00+02:00", "2026-11-21T15:30+02:00", "Asia/Jerusalem") == ("2026-11-21", "2026-11-22")
    # no zone (recurring series from iCalendar): nearest midnight
    assert all_day_dates("2026-11-11T00:00+02:00", "2026-11-12T00:00+02:00") == ("2026-11-11", "2026-11-12")
    assert all_day_dates("2026-11-08T19:00-05:00", "2026-11-09T19:00-05:00", "bogus/zone") == ("2026-11-09", "2026-11-10")
    assert all_day_dates(None, None) == (None, None)


def test_all_day_events_come_back_as_dates(make):
    def row(kind, eid, start, end, all_day, ics=""):
        return (kind, eid, 7, "Sigd", start, end, "", all_day, "", "«constant ****eSFr»", ics, "Etc/GMT")

    # Synced all-day events are stored at midnight UTC; Outlook reports them in local time.
    from datetime import UTC, datetime

    from outlook_app_mcp.mac.osa import to_stamp

    utc_midnight = [to_stamp(datetime(2026, 11, d, tzinfo=UTC)) for d in (9, 10)]
    ev = rows(
        row("single", 1, *utc_midnight, "true"),
        row("single", 2, "2026-11-9-36000", "2026-11-9-39600", "false"),
    )
    o, _ = make({"read/calendars": CALS, "read/events": ev})
    got = [(e["id"], e["start"], e["end"], e["account"]) for e in o.list_events("2026-11-08", "2026-11-12")["events"]]
    assert got[0] == ("1", "2026-11-09", "2026-11-10", "me@work.example.com")
    assert got[1][1].startswith("2026-11-09T10:00")
