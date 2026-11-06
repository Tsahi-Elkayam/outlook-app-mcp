"""Calendar: recurrence, occurrences, attendees, invitations, cancel and respond."""

import pytest
from conftest import FOLDERS, message_raw, rows

from outlook_app_mcp.common import to_html
from outlook_app_mcp.mac.calendar import NO_ATTENDEE_REMOVAL, parse_recurrence
from outlook_app_mcp.mac.osa import FS, GS, RS, US, OutlookError

CALS = rows((7, "me@work.example.com", "Calendar", "true"), (8, "me@work.example.com", "Birthdays", "false"))
ME = "me@work.example.com"

WEEKLY_ICS = "\r\n".join([  # noqa: FLY002
    "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", "UID:weekly-1",
    "DTSTART:20261102T100000", "DTEND:20261102T110000", "RRULE:FREQ=WEEKLY;BYDAY=MO", "EXDATE:20261116T100000",
    "SUMMARY:Weekly", "END:VEVENT", "END:VCALENDAR",
])
ALLDAY_ICS = "\r\n".join([  # noqa: FLY002
    "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", "UID:holiday-1",
    "DTSTART;VALUE=DATE:20261102", "DTEND;VALUE=DATE:20261103", "RRULE:FREQ=WEEKLY;BYDAY=MO",
    "SUMMARY:Day", "END:VEVENT", "END:VCALENDAR",
])
OCC = "2026-11-09T10:00"  # the second Monday of WEEKLY_ICS


def att(name, address, kind="ATrq", status="ASnr"):
    return US.join([name, address, f"«constant ****{kind}»", f"«constant ****{status}»"])


def event_full(eid="5", calendar="7", subject="Review", start="2026-11-10-36000", end="2026-11-10-39600",
               all_day="false", recurring="false", organizer=ME, attendees=(), is_occurrence="false", rec=(),
               ics="", body="desc"):
    """A read/event reply with all 27 header fields."""
    rec = list(rec) + [""] * (8 - len(rec))
    header = [eid, calendar, subject, start, end, "Room", all_day, recurring, organizer, "«constant ****eSBu»",
              "false", "true", "15", GS.join(attendees), "", "", is_occurrence, "", *rec, ics]
    assert len(header) == 27
    return FS.join(header) + RS + body


def series(**kw):
    kw.setdefault("recurring", "true")
    kw.setdefault("start", "2026-11-2-36000")
    kw.setdefault("end", "2026-11-2-39600")
    kw.setdefault("ics", WEEKLY_ICS)
    return event_full(**kw)


def reply_by_id(table):
    def reply(args):
        got = table[args[0]]
        if isinstance(got, Exception):
            raise got
        return got

    return reply


def cal(make, **replies):
    return make({"read/calendars": CALS, **replies})


def no_writes(r, *scripts):
    return not any(c[0] in (scripts or ("write/event", "write/cancel", "write/respond")) for c in r.calls)


# -- parse_recurrence -------------------------------------------------------------


def rec_fields(*tail):
    return [""] * 18 + list(tail) + [""] * (8 - len(tail))


def test_parse_recurrence_none_without_a_pattern():
    assert parse_recurrence([""] * 27) is None
    assert parse_recurrence(["x"] * 10) is None  # short header
    assert parse_recurrence(rec_fields("")) is None


def test_parse_recurrence_weekly_with_days_and_interval():
    f = rec_fields("weekly", "2", GS.join(["monday", "wednesday"]), "", "", "", "no end type", "")
    assert parse_recurrence(f) == {"frequency": "weekly", "interval": 2, "days_of_week": ["monday", "wednesday"],
                                   "until": None, "count": None}
    assert parse_recurrence(rec_fields("weekly", "1", "", "", "", "", "no end type", ""))["days_of_week"] == []


def test_parse_recurrence_monthly_and_yearly_kinds():
    f = parse_recurrence(rec_fields("absolute monthly", "3", "", "15", "", "", "no end type", ""))
    assert f == {"frequency": "monthly", "interval": 3, "day_of_month": 15, "until": None, "count": None}
    f = parse_recurrence(rec_fields("absolute yearly", "1", "", "5", "3", "", "no end type", ""))
    assert f["frequency"] == "yearly" and f["day_of_month"] == 5 and f["month"] == 3
    f = parse_recurrence(rec_fields("daily", "", "", "", "", "", "no end type", ""))
    assert f["frequency"] == "daily" and f["interval"] == 1


def test_parse_recurrence_relative_patterns_are_described_not_settable():
    f = parse_recurrence(rec_fields("relative monthly", "1", GS.join(["friday"]), "", "", "2", "no end type", ""))
    assert f["frequency"] == "monthly" and "ordinal 2" in f["pattern"] and "friday" in f["pattern"]
    assert "can't be set" in f["pattern"]
    f = parse_recurrence(rec_fields("relative yearly", "1", "monday", "", "11", "1", "no end type", ""))
    assert f["frequency"] == "yearly" and f["month"] == 11 and "pattern" in f


def test_parse_recurrence_end_types():
    f = parse_recurrence(rec_fields("daily", "1", "", "", "", "", "end date type", "2026-12-31-0"))
    assert f["until"] == "2026-12-31" and f["count"] is None
    f = parse_recurrence(rec_fields("daily", "1", "", "", "", "", "end numbered type", "10"))
    assert f["count"] == 10 and f["until"] is None
    f = parse_recurrence(rec_fields("daily", "1", "", "", "", "", "end numbered type", "many"))
    assert f["count"] is None


def test_read_event_reports_the_recurrence(make):
    rec = ("weekly", "1", "monday", "", "", "", "end numbered type", "6")
    o, _ = cal(make, **{"read/event": series(rec=rec)})
    ev = o.read_event(5)
    assert ev["is_recurring"] is True and ev["occurrence_start"] is None
    assert ev["recurrence"]["frequency"] == "weekly" and ev["recurrence"]["count"] == 6
    assert not any(k.startswith("_") for k in ev)
    o, _ = cal(make, **{"read/event": event_full()})
    assert o.read_event(5)["recurrence"] is None


def test_read_event_my_response(make):
    boss = event_full(organizer="boss@x.co", attendees=[att("Me", ME, status="ASte"), att("Avi", "avi@x.co")])
    o, _ = cal(make, **{"read/event": boss})
    ev = o.read_event(5)
    assert ev["is_organizer"] is False
    assert ev["my_response"] == "tentative"
    mine = event_full(attendees=[att("Avi", "avi@x.co")])
    o, _ = cal(make, **{"read/event": mine})
    assert o.read_event(5)["my_response"] == "organizer"
    o, _ = cal(make, **{"read/event": event_full()})
    assert o.read_event(5)["my_response"] == "none"


# -- read_event with occurrence_start -------------------------------------------------


def test_read_event_one_occurrence_of_a_series(make):
    o, _ = cal(make, **{"read/event": series()})
    ev = o.read_event(5, occurrence_start=OCC)
    assert ev["start"].startswith("2026-11-09T10:00") and ev["end"].startswith("2026-11-09T11:00")
    assert ev["occurrence_start"] == ev["start"] and ev["is_recurring"] is True
    assert o.read_event(5, occurrence_start="2026-11-09T10:00:30")["start"].startswith("2026-11-09T10:00")


def test_read_event_occurrence_refusals(make):
    o, _ = cal(make, **{"read/event": series()})
    with pytest.raises(OutlookError, match="no occurrence starting at 2026-11-10T10:00"):
        o.read_event(5, occurrence_start="2026-11-10T10:00")
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.read_event(5, occurrence_start="2026-11-16T10:00")  # EXDATE
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.read_event(5, occurrence_start="2026-11-09T10:05")  # a minute off is not the same occurrence
    with pytest.raises(OutlookError, match="not an ISO date"):
        o.read_event(5, occurrence_start="monday")
    o, _ = cal(make, **{"read/event": event_full()})
    with pytest.raises(OutlookError, match="does not repeat"):
        o.read_event(5, occurrence_start=OCC)
    o, _ = cal(make, **{"read/event": series(ics="")})
    with pytest.raises(OutlookError, match="recurrence data"):
        o.read_event(5, occurrence_start=OCC)


def test_read_event_all_day_occurrence_matches_by_date(make):
    o, _ = cal(make, **{"read/event": series(all_day="true", ics=ALLDAY_ICS)})
    ev = o.read_event(5, occurrence_start="2026-11-09")
    assert ev["start"] == "2026-11-09" and ev["end"] == "2026-11-10"
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.read_event(5, occurrence_start="2026-11-10")


# -- list_events: exceptions and occurrences ------------------------------------------


def series_row(eid=2):
    return ("series", eid, 7, "Weekly", "2026-11-2-36000", "2026-11-2-39600", "", "false", ME,
            "«constant ****eSBu»", WEEKLY_ICS, "", "", "", "")


def single_row(eid, subject, start, end, is_occ="", series_id="", orig=""):
    return ("single", eid, 7, subject, start, end, "", "false", ME, "«constant ****eSBu»", "", "", is_occ,
            series_id, orig)


def listed(make, *event_rows):
    o, _ = cal(make, **{"read/events": rows(*event_rows)})
    return o.list_events("2026-11-04", "2026-11-24")["events"]


def test_list_events_drops_a_series_occurrence_that_has_a_changed_exception(make):
    events = listed(make, series_row(),
                    single_row(3, "Moved", "2026-11-9-50400", "2026-11-9-54000", "true", "2", "2026-11-9-36000"))
    got = [(e["subject"], e["start"][:16]) for e in events]
    # the 9th at 10:00 is replaced by the moved copy; the 16th is an EXDATE; the 23rd stays
    assert got == [("Moved", "2026-11-09T14:00"), ("Weekly", "2026-11-23T10:00")]
    moved, regular = events
    assert moved["recurring"] is True and moved["series_id"] == "2" and "occurrence_start" not in moved
    assert regular["recurring"] is True and regular["occurrence_start"] == regular["start"]
    assert regular["id"] == "2"


def test_list_events_exception_of_another_series_drops_nothing(make):
    events = listed(make, series_row(),
                    single_row(3, "Other", "2026-11-9-50400", "2026-11-9-54000", "true", "99", "2026-11-9-36000"))
    assert [(e["subject"], e["start"][:16]) for e in events] == [
        ("Weekly", "2026-11-09T10:00"), ("Other", "2026-11-09T14:00"), ("Weekly", "2026-11-23T10:00")]


def test_list_events_plain_single_events_are_not_recurring(make):
    (ev,) = listed(make, single_row(4, "One-off", "2026-11-5-32400", "2026-11-5-36000", "false"))
    assert ev["recurring"] is False and "series_id" not in ev and "occurrence_start" not in ev


# -- create_event ------------------------------------------------------------------------

START, END = "2026-11-09T10:00", "2026-11-09T11:00"  # a Monday
HEAD = ["create", "", "2026-11-9-36000", "2026-11-9-39600"]


def creator(make, **extra):
    return cal(make, **{"write/event": "5", "read/event": event_full(), **extra})


def rec_value(make, spec, start=START, end=END):
    o, r = creator(make)
    o.create_event("x", start, end, recurrence=spec)
    args = r.last("write/event")
    pairs = dict(zip(args[4::2], args[5::2]))
    return pairs["recurrence"].split(US)


def test_create_event_recurrence_values(make):
    assert rec_value(make, {"frequency": "daily", "interval": 2}) == [
        "daily", "2", "", "", "", "none", "", "2026-11-9-36000"]
    assert rec_value(make, {"frequency": "weekly", "days_of_week": ["wednesday", "Monday"], "count": 5}) == [
        "weekly", "1", GS.join(["monday", "wednesday"]), "", "", "count", "5", "2026-11-9-36000"]
    assert rec_value(make, {"frequency": "weekly"})[2] == "monday"  # defaults to the start's weekday
    assert rec_value(make, {"frequency": "monthly", "day_of_month": 15, "until": "2026-12-31"}) == [
        "absolute monthly", "1", "", "15", "", "date", "2026-12-31-0", "2026-11-9-36000"]
    assert rec_value(make, {"frequency": "yearly"}) == [
        "absolute yearly", "1", "", "9", "11", "none", "", "2026-11-9-36000"]


def test_create_event_recurrence_refusals(make):
    o, r = creator(make)
    for spec, match in [
        ({"frequency": "hourly"}, "frequency"),
        ({"frequency": "daily", "count": 3, "until": "2026-12-01"}, "not both"),
        ({"frequency": "daily", "until": "2026-11-01"}, "before the first"),
        ({"frequency": "daily", "interval": 0}, "interval"),
        ({"frequency": "daily", "days_of_week": ["monday"]}, "only for weekly"),
        ({"frequency": "weekly", "day_of_month": 3}, "only for monthly"),
        ({"frequency": "daily", "bogus": 1}, "unknown key"),
    ]:
        with pytest.raises(OutlookError, match=match):
            o.create_event("x", START, END, recurrence=spec)
    assert no_writes(r)


def test_create_event_full_pair_order(make):
    o, r = creator(make, **{"read/categories": rows(("Work", "")), "write/category": "x"})
    res = o.create_event("Sync", START, END, calendar_id=8, attendees=["a@x.co", "A@x.co", "b@x.co"],
                         optional_attendees=["c@x.co"], categories=["work"], is_private=True,
                         recurrence={"frequency": "daily"})
    args = r.last("write/event")
    assert args[:4] == ["create", "8", "2026-11-9-36000", "2026-11-9-39600"]
    assert args[4:] == [
        "subject", "Sync", "private", "1",
        "recurrence", US.join(["daily", "1", "", "", "", "none", "", "2026-11-9-36000"]),
        "addreq", GS.join(["a@x.co", "A@x.co", "b@x.co"]),
        "addopt", "c@x.co",
        "categories", "Work",
    ]
    assert res["invitations_sent"] is False and "send_update=true" in res["note"]
    assert not any(c[1][0] == "send" for c in r.calls if c[0] == "write/event")


def test_create_event_private_false_and_no_attendees(make):
    o, r = creator(make)
    res = o.create_event("Solo", START, END, is_private=False)
    assert r.last("write/event")[4:] == ["subject", "Solo", "private", "0"]
    assert "invitations_sent" not in res and "note" not in res


def test_create_event_send_invitations_runs_the_send_script(make):
    o, r = creator(make)
    res = o.create_event("Sync", START, END, attendees=["a@x.co"], send_invitations=True)
    sent = [c[1] for c in r.calls if c[0] == "write/event"]
    assert [a[0] for a in sent] == ["create", "send"]
    assert r.last("write/event") == ["send", "5", ""]
    assert res["invitations_sent"] is True and "note" not in res


def test_create_event_send_invitations_needs_attendees(make):
    o, r = creator(make)
    with pytest.raises(OutlookError, match="send_invitations needs attendees"):
        o.create_event("Sync", START, END, send_invitations=True)
    with pytest.raises(OutlookError, match="not an email"):
        o.create_event("Sync", START, END, attendees=["nope"])
    assert no_writes(r)


GOOGLE_FOLDERS = FOLDERS + RS + rows((159, "google:151", "", "Inbox", 5, 900, "inbox", "google")) + GS + FS.join(
    ["151", US.join(["me@gmail.example"] * 6), ""])


def test_create_event_with_attendees_refused_on_a_google_calendar(make):
    gcals = CALS + RS + rows((9, "me@gmail.example", "Gmail", "false"))
    o, r = make({"read/folders": GOOGLE_FOLDERS, "read/calendars": gcals, "write/event": "5",
                 "read/event": event_full(calendar="9")})
    with pytest.raises(OutlookError, match="Google calendar"):
        o.create_event("Sync", START, END, calendar_id=9, attendees=["a@x.co"])
    assert no_writes(r)
    o.create_event("Solo", START, END, calendar_id=9)
    assert r.last("write/event")[:2] == ["create", "9"]


# -- update_event -----------------------------------------------------------------------------

AVI = att("Avi", "avi@x.co")
OPT = att("Opt", "opt@x.co", kind="ATop")


def updater(make, raw, **extra):
    return cal(make, **{"read/event": raw, "write/event": "5", **extra})


def test_update_event_cannot_remove_attendees_or_demote_them(make):
    o, r = updater(make, event_full(attendees=[AVI, OPT]))
    for kwargs in ({"attendees": []}, {"optional_attendees": []}, {"attendees": ["other@x.co"]},
                   {"optional_attendees": ["avi@x.co"]}):
        with pytest.raises(OutlookError) as exc:
            o.update_event(5, **kwargs)
        assert str(exc.value) == NO_ATTENDEE_REMOVAL
    assert no_writes(r)


def test_update_event_adds_attendees_only(make):
    o, r = updater(make, event_full(attendees=[AVI, OPT]))
    res = o.update_event(5, attendees=["AVI@x.co", "new@x.co"], optional_attendees=["opt@x.co", "o2@x.co"])
    assert r.last("write/event") == ["update", "5", "addreq", "new@x.co", "addopt", "o2@x.co"]
    assert res["update_sent"] is False and "attendees were not told" in res["note"]
    o.update_event(5, attendees=["new@x.co", "avi@x.co"])  # optional untouched
    assert r.last("write/event") == ["update", "5", "addreq", "new@x.co"]


def test_update_event_attendees_already_there_changes_nothing(make):
    o, r = updater(make, event_full(attendees=[AVI]))
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_event(5, attendees=["avi@x.co"])
    assert no_writes(r)


def test_update_event_send_update(make):
    o, r = updater(make, event_full(attendees=[AVI]))
    res = o.update_event(5, location="Room 2", send_update=True)
    assert [c[1] for c in r.calls if c[0] == "write/event"] == [["update", "5", "location", "Room 2"],
                                                                 ["send", "5", ""]]
    assert res["update_sent"] is True
    res = o.update_event(5, send_update=True)  # nothing changed, but tell them
    assert r.last("write/event") == ["send", "5", ""] and res["update_sent"] is True
    assert len([c for c in r.calls if c[0] == "write/event"]) == 3


def test_update_event_send_update_needs_attendees(make):
    o, r = updater(make, event_full())
    with pytest.raises(OutlookError, match="send_update needs a meeting with attendees"):
        o.update_event(5, location="x", send_update=True)
    assert not any(c[1][0] == "send" for c in r.calls if c[0] == "write/event")


def test_update_event_only_the_organizer_changes_attendees_or_sends(make):
    o, r = updater(make, event_full(organizer="boss@x.co", attendees=[AVI]))
    for kwargs in ({"attendees": ["avi@x.co", "n@x.co"]}, {"optional_attendees": []}, {"send_update": True}):
        with pytest.raises(OutlookError, match="Only the organizer"):
            o.update_event(5, **kwargs)
    assert no_writes(r)


def test_update_event_one_occurrence(make):
    table = {"5": series(), "50": event_full("50", start="2026-11-9-36000", end="2026-11-9-39600")}
    o, r = cal(make, **{"read/event": reply_by_id(table), "write/event": "50"})
    res = o.update_event(5, occurrence_start=OCC, location="Room 9")
    assert r.last("write/event") == ["occurrence", "5", "2026-11-9-36000", "location", "Room 9"]
    assert res["id"] == "50" and res["occurrence_start"] == res["start"]
    assert res["occurrence_start"].startswith("2026-11-09T10:00")


def test_update_event_one_occurrence_moved_later_sets_end_first(make):
    table = {"5": series(), "50": event_full("50")}
    o, r = cal(make, **{"read/event": reply_by_id(table), "write/event": "50"})
    o.update_event(5, occurrence_start=OCC, start="2026-11-09T14:00", end="2026-11-09T15:00")
    assert r.last("write/event") == ["occurrence", "5", "2026-11-9-36000", "end", "2026-11-9-54000",
                                     "start", "2026-11-9-50400"]
    o.update_event(5, occurrence_start=OCC, start="2026-11-09T09:00", end="2026-11-09T09:30")
    assert r.last("write/event")[3:] == ["start", "2026-11-9-32400", "end", "2026-11-9-34200"]


def test_update_event_one_occurrence_refusals(make):
    o, r = updater(make, series())
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.update_event(5, occurrence_start="2026-11-10T10:00", location="x")
    with pytest.raises(OutlookError, match="end before it starts"):
        o.update_event(5, occurrence_start=OCC, start="2026-11-09T12:00")
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_event(5, occurrence_start=OCC)
    assert no_writes(r)
    o, r = updater(make, event_full())
    with pytest.raises(OutlookError, match="does not repeat"):
        o.update_event(5, occurrence_start=OCC, location="x")


def test_update_event_an_exception_needs_no_whole_series_flag(make):
    o, r = updater(make, event_full(recurring="true", is_occurrence="true"))
    o.update_event(5, subject="Changed")
    assert r.last("write/event") == ["update", "5", "subject", "Changed"]


def test_update_event_categories_and_private(make):
    o, r = updater(make, event_full(), **{"read/categories": rows(("Work", ""))})
    o.update_event(5, categories=["WORK"], is_private=False)
    assert r.last("write/event") == ["update", "5", "private", "0", "categories", "Work"]


# -- delete_event ---------------------------------------------------------------------------------


def test_delete_event_refuses_a_single_occurrence(make):
    o, r = updater(make, series())
    for kwargs in ({"occurrence_start": OCC}, {"occurrence_start": OCC, "whole_series": True}):
        with pytest.raises(OutlookError, match="single occurrence"):
            o.delete_event(5, **kwargs)
    o, r2 = updater(make, event_full(recurring="true", is_occurrence="true"))
    with pytest.raises(OutlookError, match="single occurrence"):
        o.delete_event(5, whole_series=True)
    assert no_writes(r, "write/event") and no_writes(r2, "write/event")


# -- cancel_meeting -----------------------------------------------------------------------------


def test_cancel_meeting_sends_the_cancellation(make):
    o, r = cal(make, **{"read/event": event_full(attendees=[AVI]), "write/cancel": "5"})
    res = o.cancel_meeting(5)
    assert r.last("write/cancel") == ["5", "", ""]
    assert res["cancelled"] == "5" and res["subject"] == "Review" and res["occurrence_start"] is None
    o.cancel_meeting(5, comment="Sorry\n\nplans changed")
    assert r.last("write/cancel") == ["5", "", to_html("Sorry\n\nplans changed")]


def test_cancel_meeting_one_occurrence(make):
    o, r = cal(make, **{"read/event": series(attendees=[AVI]), "write/cancel": "5"})
    res = o.cancel_meeting(5, occurrence_start=OCC)
    assert r.last("write/cancel") == ["5", "2026-11-9-36000", ""]
    assert res["occurrence_start"].startswith("2026-11-09T10:00")
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.cancel_meeting(5, occurrence_start="2026-11-16T10:00")
    assert len([c for c in r.calls if c[0] == "write/cancel"]) == 1


def test_cancel_meeting_needs_organizer_and_attendees(make):
    o, r = cal(make, **{"read/event": event_full(organizer="boss@x.co", attendees=[AVI]), "write/cancel": "5"})
    with pytest.raises(OutlookError, match="Only the organizer"):
        o.cancel_meeting(5)
    o, r2 = cal(make, **{"read/event": event_full(), "write/cancel": "5"})
    with pytest.raises(OutlookError, match="no attendees"):
        o.cancel_meeting(5)
    o, r3 = cal(make, **{"read/event": event_full(calendar="106", attendees=[AVI]), "write/cancel": "5"})
    with pytest.raises(OutlookError, match="not in one of your calendars"):
        o.cancel_meeting(5)
    assert no_writes(r, "write/cancel") and no_writes(r2, "write/cancel") and no_writes(r3, "write/cancel")


# -- respond_to_meeting -------------------------------------------------------------------------------

NO_MESSAGE = OutlookError("Outlook could not find that item")
NO_EVENT = OutlookError("Outlook could not find that item")


def meeting_message(meeting_event_id="", is_meeting="true"):
    head, _, body = message_raw().partition(RS)
    f = head.split(FS)
    f[16] = is_meeting
    f += ["priority normal", "", meeting_event_id]
    return FS.join(f) + RS + body


def responder(make, event=NO_EVENT, message=NO_MESSAGE):
    return cal(make, **{"read/event": _reply(event), "read/message": _reply(message), "write/respond": "5"})


def _reply(value):
    """A reply that is a callable, an exception to raise, or a fixed string."""
    if callable(value):
        return value

    def reply(args):
        if isinstance(value, Exception):
            raise value
        return value

    return reply


def test_respond_event_path_args(make):
    o, r = responder(make, event=event_full(organizer="boss@x.co"))
    res = o.respond_to_meeting(5, "accept")
    assert r.last("write/respond") == ["event", "5", "", "accept", "", "1"]
    assert res == {"id": "5", "response": "accept", "organizer_told": True, "occurrence_start": None}
    res = o.respond_to_meeting(5, "tentative", comment="maybe", send_response=False)
    assert r.last("write/respond") == ["event", "5", "", "tentative", "maybe", "0"]
    assert res["organizer_told"] is False


def test_respond_to_one_occurrence(make):
    o, r = responder(make, event=series(organizer="boss@x.co"))
    res = o.respond_to_meeting(5, "decline", occurrence_start=OCC)
    assert r.last("write/respond") == ["event", "5", "2026-11-9-36000", "decline", "", "1"]
    assert res["occurrence_start"].startswith("2026-11-09T10:00")
    with pytest.raises(OutlookError, match="no occurrence starting"):
        o.respond_to_meeting(5, "decline", occurrence_start="2026-11-16T10:00")
    assert len([c for c in r.calls if c[0] == "write/respond"]) == 1


def test_respond_refuses_the_organizer(make):
    o, r = responder(make, event=event_full(attendees=[AVI]))
    with pytest.raises(OutlookError, match="You organize this meeting"):
        o.respond_to_meeting(5, "accept")
    assert no_writes(r, "write/respond")


def test_respond_message_path(make):
    o, r = responder(make, message=meeting_message())
    res = o.respond_to_meeting(5, "tentative", comment="hm", send_response=False)
    assert r.last("write/respond") == ["message", "5", "", "tentative", "hm", "0"]
    assert res == {"id": "5", "response": "tentative", "organizer_told": False}
    with pytest.raises(OutlookError, match="occurrence_start goes with an event id"):
        o.respond_to_meeting(5, "accept", occurrence_start=OCC)
    assert len([c for c in r.calls if c[0] == "write/respond"]) == 1


def test_respond_message_that_points_at_an_event_answers_the_event(make):
    table = {"5": NO_EVENT, "77": event_full("77", organizer="boss@x.co")}
    o, r = responder(make, event=reply_by_id(table), message=meeting_message("77"))
    res = o.respond_to_meeting(5, "accept")
    assert r.last("write/respond") == ["event", "77", "", "accept", "", "1"]
    assert res["id"] == "77"


def test_respond_ambiguous_id_is_an_error(make):
    o, r = responder(make, event=event_full(organizer="boss@x.co"), message=meeting_message("77"))
    with pytest.raises(OutlookError, match="both a calendar event and a meeting message.*77"):
        o.respond_to_meeting(5, "accept")
    assert no_writes(r, "write/respond")


def test_respond_event_wins_when_the_message_agrees_or_is_not_a_meeting(make):
    for message in (meeting_message("5"), meeting_message(""), meeting_message("77", is_meeting="false")):
        o, r = responder(make, event=event_full(organizer="boss@x.co"), message=message)
        o.respond_to_meeting(5, "accept")
        assert r.last("write/respond")[:2] == ["event", "5"]


def test_respond_to_nothing_and_bad_response(make):
    o, r = responder(make)
    with pytest.raises(OutlookError, match="No calendar event or meeting message with id 5"):
        o.respond_to_meeting(5, "accept")
    o, r = responder(make, message=meeting_message("", is_meeting="false"))
    with pytest.raises(OutlookError, match="No calendar event or meeting message"):
        o.respond_to_meeting(5, "accept")
    with pytest.raises(OutlookError, match="response must be one of"):
        o.respond_to_meeting(5, "maybe")
    with pytest.raises(OutlookError, match="event_id"):
        o.respond_to_meeting("abc", "accept")
    assert no_writes(r, "write/respond")
