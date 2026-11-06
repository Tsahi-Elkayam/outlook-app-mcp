"""Windows backend, 0.3 calendar: occurrences, recurrence, attendees and sending, cancelling, free/busy."""

from __future__ import annotations

from datetime import datetime

import pytest
from fakecom import Appointment
from test_windows import ME, local, wid

from outlook_app_mcp.common import NOT_ORGANIZER, OutlookError
from outlook_app_mcp.windows.com import to_com


def sent_log(app):
    return [e for e in app.log if e[0] == "send_appt"]


def organized_meeting(o, **extra):
    """A meeting this user organizes, saved without sending."""
    return o.create_event("Sync", "2026-11-12T14:00", "2026-11-12T15:00", attendees=["a@x.co"],
                          optional_attendees=["b@x.co"], **extra)


# -- occurrences ---------------------------------------------------------------------------


def test_list_events_gives_occurrence_start_for_recurring_events_only(w):
    o, _, _ = w
    events = {e["subject"]: e for e in o.list_events("2026-11-04", "2026-11-11")["events"]}
    assert events["Weekly"]["occurrence_start"] == events["Weekly"]["start"]
    assert events["Weekly"]["start"].startswith("2026-11-09T10:00")
    assert "occurrence_start" not in events["One-off"]


def test_read_event_has_the_new_fields(w):
    o, _, r = w
    ev = o.read_event(wid(r.a, r.one_off))
    assert ev["recurrence"] is None and ev["occurrence_start"] is None and ev["categories"] == []
    assert ev["is_private"] is False and ev["my_response"] == "none" and ev["is_organizer"] is True
    assert r.one_off._pattern is None  # reading must not turn a plain event into a recurring one
    meeting = o.read_event(wid(r.a, r.meeting))
    assert meeting["is_organizer"] is False
    r.meeting.ResponseStatus = 3
    assert o.read_event(wid(r.a, r.meeting))["my_response"] == "accepted"
    r.meeting.ResponseStatus = 1
    assert o.read_event(wid(r.a, r.meeting))["is_organizer"] is True


def test_read_event_organizer_by_name(w):
    o, _, r = w
    r.meeting.Organizer = ME
    assert o.read_event(wid(r.a, r.meeting))["is_organizer"] is True


def test_read_one_occurrence_of_a_series(w):
    o, _, r = w
    ev = o.read_event(wid(r.a, r.weekly), occurrence_start="2026-11-16T10:00")
    assert ev["start"].startswith("2026-11-16T10:00") and ev["end"].startswith("2026-11-16T11:00")
    assert ev["occurrence_start"] == ev["start"] and ev["id"] == wid(r.a, r.weekly)
    assert ev["recurrence"] == {"frequency": "weekly", "interval": 1, "days_of_week": ["monday"]}
    assert o.read_event(wid(r.a, r.weekly))["occurrence_start"] is None
    with pytest.raises(OutlookError, match="no occurrence starting at"):
        o.read_event(wid(r.a, r.weekly), occurrence_start="2026-11-17T10:00")
    with pytest.raises(OutlookError, match="does not repeat"):
        o.read_event(wid(r.a, r.one_off), occurrence_start="2026-11-05T09:00")


# -- creating: recurrence, categories, privacy ---------------------------------------------


@pytest.mark.parametrize(("rec", "expected", "pattern"), [
    ({"frequency": "daily", "interval": 2, "count": 3},
     {"frequency": "daily", "interval": 2, "until": "2026-11-16"}, (0, 2)),
    ({"frequency": "weekly", "days_of_week": ["monday", "wednesday"], "until": "2026-12-31"},
     {"frequency": "weekly", "interval": 1, "days_of_week": ["monday", "wednesday"], "until": "2026-12-31"}, (1, 1)),
    ({"frequency": "monthly", "interval": 3},
     {"frequency": "monthly", "interval": 3, "day_of_month": 12}, (2, 3)),
    ({"frequency": "yearly"},
     {"frequency": "yearly", "interval": 1, "day_of_month": 12, "month": 11}, (5, 12)),
])
def test_create_recurring_event(w, rec, expected, pattern):
    o, _, r = w
    ev = o.create_event("Series", "2026-11-12T14:00", "2026-11-12T15:00", recurrence=rec)
    item = r.cal._items[-1]
    assert ev["is_recurring"] is True and ev["recurrence"] == expected
    assert (item._pattern.RecurrenceType, item._pattern.Interval) == pattern and item.saved == 1


def test_weekly_recurrence_defaults_to_the_start_weekday_and_sets_the_mask(w):
    o, _, r = w
    o.create_event("Thursdays", "2026-11-12T14:00", "2026-11-12T15:00", recurrence={"frequency": "weekly"})
    assert r.cal._items[-1]._pattern.DayOfWeekMask == 16  # 2026-11-12 is a Thursday
    o.create_event("Two", "2026-11-12T14:00", "2026-11-12T15:00",
                   recurrence={"frequency": "weekly", "days_of_week": ["sunday", "saturday"]})
    assert r.cal._items[-1]._pattern.DayOfWeekMask == 1 + 64


def test_recurrence_is_checked(w):
    o, _, _ = w
    with pytest.raises(OutlookError, match="until or count"):
        o.create_event("x", "2026-11-12T14:00", "2026-11-12T15:00",
                       recurrence={"frequency": "daily", "until": "2026-12-01", "count": 2})
    with pytest.raises(OutlookError, match="frequency"):
        o.create_event("x", "2026-11-12T14:00", "2026-11-12T15:00", recurrence={"frequency": "hourly"})


def test_create_event_categories_and_privacy(w):
    o, _, r = w
    ev = o.create_event("Private", "2026-11-12T14:00", "2026-11-12T15:00", categories=["Blue", "New one"],
                        is_private=True)
    item = r.cal._items[-1]
    assert ev["is_private"] is True and item.Sensitivity == 2 and ev["categories"] == ["Blue", "New one"]
    assert {c.Name for c in w[1].Session.Categories} == {"Blue", "New one"}  # missing ones are created
    ev = o.create_event("Public", "2026-11-13T14:00", "2026-11-13T15:00", is_private=False)
    assert ev["is_private"] is False


# -- creating: attendees ---------------------------------------------------------------------


def test_meeting_is_saved_unsent_unless_asked(w):
    o, app, r = w
    ev = organized_meeting(o)
    item = r.cal._items[-1]
    assert item.MeetingStatus == 1 and not sent_log(app) and item.saved == 1
    assert [(x.Address, x.Type) for x in item.Recipients] == [("a@x.co", 1), ("b@x.co", 2)]
    assert ev["invitations_sent"] is False and "not been invited" in ev["note"] and ev["is_meeting"]
    assert [a["type"] for a in ev["attendees"]] == ["required", "optional"] and ev["is_organizer"] is True


def test_invitations_go_out_only_with_send_invitations(w):
    o, app, r = w
    ev = organized_meeting(o, send_invitations=True)
    assert ev["invitations_sent"] is True and "note" not in ev
    assert sent_log(app) == [("send_appt", "Sync", 1)]
    assert r.cal._items[-1].saved == 1  # saved first (it has no id before), then sent


def test_unresolved_attendees_stop_a_send_and_leave_nothing(w, monkeypatch):
    import fakecom

    o, app, r = w
    monkeypatch.setattr(fakecom.Recipients, "ResolveAll", lambda self: False)
    before = len(r.cal._items)
    with pytest.raises(OutlookError, match="could not resolve"):
        organized_meeting(o, send_invitations=True)
    assert len(r.cal._items) == before and not sent_log(app)
    with pytest.raises(OutlookError, match="needs attendees"):
        o.create_event("x", "2026-11-12T14:00", "2026-11-12T15:00", send_invitations=True)
    with pytest.raises(OutlookError, match="not an email"):
        o.create_event("x", "2026-11-12T14:00", "2026-11-12T15:00", attendees=["nope"])


# -- updating ------------------------------------------------------------------------------------


def test_update_attendees_saves_only_and_says_so(w):
    o, app, r = w
    ev = organized_meeting(o)
    res = o.update_event(ev["id"], attendees=["a@x.co", "c@x.co"], optional_attendees=[], categories=["Blue"],
                         is_private=True)
    item = r.cal._items[-1]
    assert [(x.Address, x.Type) for x in item.Recipients] == [("a@x.co", 1), ("c@x.co", 1)]
    assert res["update_sent"] is False and "not told" in res["note"] and not sent_log(app)
    assert res["categories"] == ["Blue"] and res["is_private"] is True
    # an unchanged list is not rewritten (attendees keep their answers)
    app.log.clear()
    o.update_event(ev["id"], attendees=["C@x.co", "a@x.co"], subject="Sync 2")
    assert not [e for e in app.log if e[0] == "attendees"]


def test_update_with_send_update_sends(w):
    o, app, _r = w
    ev = organized_meeting(o)
    res = o.update_event(ev["id"], location="Room 4", send_update=True)
    assert res["update_sent"] is True and "note" not in res and sent_log(app) == [("send_appt", "Sync", 1)]
    # only sending a saved meeting is allowed without any other change
    assert o.update_event(ev["id"], send_update=True)["update_sent"] is True
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_event(ev["id"])


def test_update_event_attendee_rules(w):
    o, _, r = w
    with pytest.raises(OutlookError, match="Only the organizer"):
        o.update_event(wid(r.a, r.meeting), attendees=["x@x.co"])
    assert NOT_ORGANIZER.startswith("Only the organizer")
    with pytest.raises(OutlookError, match="Only the organizer"):
        o.update_event(wid(r.a, r.meeting), send_update=True)
    with pytest.raises(OutlookError, match="nobody to send"):
        o.update_event(wid(r.a, r.one_off), send_update=True)
    ev = organized_meeting(o)
    with pytest.raises(OutlookError, match="cancel_meeting"):
        o.update_event(ev["id"], attendees=[], optional_attendees=[])
    # a plain event becomes a meeting when attendees are added
    res = o.update_event(wid(r.a, r.one_off), attendees=["a@x.co"])
    assert res["is_meeting"] is True and r.one_off.MeetingStatus == 1


def test_update_one_occurrence_makes_an_exception(w):
    o, _, r = w
    res = o.update_event(wid(r.a, r.weekly), occurrence_start="2026-11-16T10:00", start="2026-11-16T12:00",
                         end="2026-11-16T13:00", location="Elsewhere")
    assert res["start"].startswith("2026-11-16T12:00") and res["occurrence_start"] == res["start"]
    assert res["location"] == "Elsewhere" and res["id"] == wid(r.a, r.weekly)
    occ = r.weekly._pattern.cache[datetime(2026, 11, 16, 10)]
    assert occ.saved == 1 and r.weekly.Location == "" and r.weekly._raw["Start"] == local(2026, 11, 9, 10)
    assert occ not in r.cal._items
    with pytest.raises(OutlookError, match="no occurrence"):
        o.update_event(wid(r.a, r.weekly), occurrence_start="2026-11-17T10:00", location="x")
    # the series itself still refuses without whole_series
    with pytest.raises(OutlookError, match="whole_series"):
        o.update_event(wid(r.a, r.weekly), location="x")


# -- deleting ------------------------------------------------------------------------------------


def test_delete_one_occurrence_is_refused_without_a_permanent_delete(w):
    o, _, r = w
    with pytest.raises(OutlookError, match="never deletes permanently"):
        o.delete_event(wid(r.a, r.weekly), occurrence_start="2026-11-16T10:00")
    with pytest.raises(OutlookError, match="respond_to_meeting"):  # meetings are refused first
        o.delete_event(wid(r.a, r.meeting), occurrence_start="2026-11-16T10:00")
    assert r.weekly.Parent is r.cal


# -- cancelling ------------------------------------------------------------------------------------


def test_cancel_meeting_sends_then_files_in_deleted_items(w):
    o, app, r = w
    ev = organized_meeting(o, send_invitations=True)
    item = r.cal._items[-1]
    res = o.cancel_meeting(ev["id"], comment="Sorry, off")
    assert res["cancelled"] == ev["id"] and "permanently" in res["note"]
    assert item.MeetingStatus == 5 and item.Body == "Sorry, off"
    assert sent_log(app)[-1] == ("send_appt", "Sync", 5)
    assert item.Parent is r.deleted and item not in r.cal._items
    # nothing was ever deleted: Delete() would have failed in the fake


def test_cancel_meeting_when_outlook_files_the_item_itself(w, monkeypatch):
    o, _, r = w
    monkeypatch.setattr(Appointment, "cancel_moves_item", True)
    ev = organized_meeting(o)
    o.cancel_meeting(ev["id"])
    assert r.deleted._items[-1].Subject == "Sync" and r.deleted._items.count(r.deleted._items[-1]) == 1


def test_cancel_meeting_refusals(w):
    o, app, r = w
    with pytest.raises(OutlookError, match="Only the organizer"):
        o.cancel_meeting(wid(r.a, r.meeting))
    with pytest.raises(OutlookError, match="not a meeting"):
        o.cancel_meeting(wid(r.a, r.one_off))
    ev = organized_meeting(o)
    o.cancel_meeting(ev["id"])
    with pytest.raises(OutlookError, match="Deleted Items"):
        o.cancel_meeting(ev["id"])  # it is in Deleted Items now
    assert not [e for e in sent_log(app) if e[1] == "Meeting"]
    r.meeting.MeetingStatus = 5
    r.meeting.ResponseStatus = 1
    with pytest.raises(OutlookError, match="already cancelled"):
        o.cancel_meeting(wid(r.a, r.meeting))


def test_cancel_one_occurrence_sends_its_cancellation_and_moves_nothing(w):
    o, _app, r = w
    ev = o.create_event("Standup", "2026-11-12T09:00", "2026-11-12T09:15", attendees=["a@x.co"],
                        recurrence={"frequency": "daily", "count": 5})
    res = o.cancel_meeting(ev["id"], comment="skip", occurrence_start="2026-11-13T09:00")
    item = r.cal._items[-1]
    occ = item._pattern.cache[datetime(2026, 11, 13, 9)]
    assert res["occurrence_start"] == "2026-11-13T09:00" and occ.MeetingStatus == 5 and occ.Body == "skip"
    assert item.occurrence_sends == [(datetime(2026, 11, 13, 9), 5)]
    assert item.Parent is r.cal and item.MeetingStatus == 1  # the series is untouched


# -- responding ------------------------------------------------------------------------------------


def test_respond_with_a_meeting_request_message_id(w):
    o, _, r = w
    res = o.respond_to_meeting(wid(r.a, r.invite), "tentative", "maybe")
    assert r.meeting.responses[-1].code == 2 and r.meeting.responses[-1].sent
    assert res["id"] == wid(r.a, r.meeting) and res["organizer_told"] is True
    with pytest.raises(OutlookError, match="cancellation"):
        o.respond_to_meeting(wid(r.a, r.cancelled), "accept")
    with pytest.raises(OutlookError, match="not a calendar event or a meeting request"):
        o.respond_to_meeting(wid(r.a, r.m1), "accept")


def test_respond_to_a_request_not_yet_on_the_calendar_adds_it(w):
    o, _, r = w
    r.invite.appointment = None
    o.respond_to_meeting(wid(r.a, r.invite), "accept", send_response=False)
    added = r.invite.appointment
    assert added.Parent is r.cal and added.responses[-1].code == 3 and added.responses[-1].discarded


def test_respond_to_one_occurrence(w):
    o, _, r = w
    r.weekly.MeetingStatus = 3
    res = o.respond_to_meeting(wid(r.a, r.weekly), "decline", occurrence_start="2026-11-16T10:00")
    occ = r.weekly._pattern.cache[datetime(2026, 11, 16, 10)]
    assert occ.responses[-1].code == 4 and r.weekly.responses == []
    assert res["id"] == wid(r.a, r.weekly) and "occurrence" in res["note"] and r.weekly.Parent is r.cal


def test_the_organizer_has_nobody_to_respond_to(w):
    o, _, _ = w
    ev = organized_meeting(o)
    with pytest.raises(OutlookError, match="cancel_meeting"):
        o.respond_to_meeting(ev["id"], "accept")


# -- free/busy ---------------------------------------------------------------------------------------


def test_free_busy_slices_the_requested_window(w):
    o, app, _ = w
    app.free_busy["boss@x.co"] = "0" * 18 + "22" + "11" + "0" * 26
    res = o.get_free_busy(["boss@x.co", "ghost@x.co"], "2026-11-12T09:00", "2026-11-12T11:00", 30)
    assert res["interval_minutes"] == 30 and res["start"].startswith("2026-11-12T09:00")
    boss, ghost = res["people"]
    assert boss == {"address": "boss@x.co", "slots": "2211", "error": None}
    assert ghost["slots"] == "????" and "free/busy" in ghost["error"]
    call = next(e for e in app.log if e[0] == "freebusy")
    assert call == ("freebusy", "boss@x.co", to_com(datetime(2026, 11, 12)), 30, True)


def test_free_busy_past_the_published_range_is_unknown(w):
    o, app, _ = w
    app.free_busy["boss@x.co"] = "0" * 20
    res = o.get_free_busy(["boss@x.co"], "2026-11-12T09:00", "2026-11-12T11:00", 30)
    assert res["people"][0]["slots"] == "00??"


def test_free_busy_checks_its_input(w):
    o, _, _ = w
    with pytest.raises(OutlookError, match="at least one"):
        o.get_free_busy([], "2026-11-12", "2026-11-13", 30)
    with pytest.raises(OutlookError, match="interval_minutes"):
        o.get_free_busy(["a@x.co"], "2026-11-12", "2026-11-13", 1)
    with pytest.raises(OutlookError, match="after start"):
        o.get_free_busy(["a@x.co"], "2026-11-13", "2026-11-12", 30)
