"""macOS calendar: calendars, events (with recurring series expanded), and changes."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta

import icalendar
import recurring_ical_events

from ..common import (
    HAS_ATTENDEES,
    MAX_EVENTS,
    NOT_ORGANIZER,
    RECURRING_SERIES,
    RECURRING_TIMES,
    OutlookError,
    check_attendees,
    check_categories,
    check_event_times,
    check_free_busy,
    check_recurrence,
    check_reminder,
    check_response,
    dates_if_all_day,
    event_bound,
    event_range,
    iso,
    parse_when,
    to_html,
)
from .osa import (
    GS,
    RS,
    US,
    as_bool,
    attendee_type,
    check_id,
    entries,
    enum,
    records,
    stamp_iso,
    to_stamp,
)


def _local(value) -> datetime:
    """icalendar date / naive / aware datetime -> aware local datetime."""
    if isinstance(value, datetime):
        return value.astimezone()  # naive ("floating") times are taken as local
    if isinstance(value, date):
        return datetime.combine(value, time.min).astimezone()
    raise TypeError(value)


def expand_series(ics: str, start: datetime, end: datetime) -> list[tuple[datetime, datetime]]:
    cal = icalendar.Calendar.from_ical(ics)
    out = []
    for occ in recurring_ical_events.of(cal).between(start, end):
        s = _local(occ["DTSTART"].dt)
        e = _local(occ["DTEND"].dt) if "DTEND" in occ else s
        out.append((s, e))
    return out


RECURRENCE_KINDS = {
    "daily": "daily",
    "weekly": "weekly",
    "absolute monthly": "monthly",
    "relative monthly": "monthly",
    "absolute yearly": "yearly",
    "relative yearly": "yearly",
}
STATUS_WORDS = {"accepted": "accepted", "tentatively accepted": "tentative", "declined": "declined",
                "not responded": "not responded"}
NO_ATTENDEE_REMOVAL = (
    "Outlook for Mac's scripting can add attendees but not remove them (or change required to optional); "
    "nothing was changed. Remove attendees in Outlook."
)
NO_OCCURRENCE_DELETE = (
    "A single occurrence of a series can't be removed from here: Outlook removes it from the series for good "
    "instead of moving it to Deleted Items. Delete the whole series with whole_series=true, or remove this "
    "occurrence in Outlook."
)


def parse_recurrence(f: list[str]) -> dict | None:
    """Fields 18-25 of read/event: type, interval, days, day of month, month, ordinal, end type, end data."""
    if len(f) < 26 or not f[18]:
        return None
    kind = f[18]
    rec: dict = {"frequency": RECURRENCE_KINDS.get(kind, kind),
                 "interval": int(f[19]) if f[19].isdigit() else 1}
    if kind == "weekly":
        rec["days_of_week"] = f[20].split(GS) if f[20] else []
    if kind in ("absolute monthly", "absolute yearly") and f[21].isdigit():
        rec["day_of_month"] = int(f[21])
    if kind.endswith("yearly") and f[22].isdigit():
        rec["month"] = int(f[22])
    if kind.startswith("relative"):
        rec["pattern"] = f"nth weekday (ordinal {f[23]}, days {f[20].replace(GS, ', ')}); can't be set from here"
    rec["until"], rec["count"] = None, None
    if f[24] == "end date type" and f[25]:
        rec["until"] = (stamp_iso(f[25]) or "")[:10] or None
    elif f[24] == "end numbered type" and f[25].isdigit():
        rec["count"] = int(f[25])
    return rec


def parse_event(raw: str) -> dict:
    header, _, body = raw.partition(RS)
    f = header.split("\x1f")
    if len(f) < 15:
        raise OutlookError(f"Unexpected reply from Outlook ({len(f)} header fields).")
    attendees = []
    for e in entries(f[13]):
        attendees.append(
            {
                "name": e[0],
                "address": e[1] if len(e) > 1 else "",
                "type": attendee_type(e[2]) if len(e) > 2 else None,
                "status": enum(e[3]) if len(e) > 3 else None,
            }
        )
    ev = dates_if_all_day({
        "id": f[0],
        "calendar_id": f[1] or None,
        "subject": f[2],
        "start": stamp_iso(f[3]),
        "end": stamp_iso(f[4]),
        "location": f[5],
        "all_day": as_bool(f[6]),
        "is_recurring": as_bool(f[7]),
        "organizer": f[8],
        "free_busy": enum(f[9]),
        "is_private": as_bool(f[10]),
        "reminder_minutes": int(f[12]) if as_bool(f[11]) and f[12].lstrip("-").isdigit() else None,
        "attendees": attendees,
        "categories": f[14].split(GS) if f[14] else [],
        "description": body,
    }, f[15] if len(f) > 15 else "")
    ev["is_occurrence"] = as_bool(f[16]) if len(f) > 16 else False
    ev["recurrence"] = parse_recurrence(f)
    ev["_ics"] = f[26] if len(f) > 26 else ""
    return ev


class CalendarMixin:
    """Needs self._run."""

    def list_calendars(self) -> list[dict]:
        self.require_accounts()
        cals = []
        for f in records(self._run("read/calendars", [])):
            if not f[2]:
                continue  # an account's unnamed root, which holds calendars, not events
            account = f[1] or (self.mailbox_account(f[4]) if len(f) > 4 else None) or ""
            cals.append({"id": f[0], "account": account, "name": f[2], "is_default": as_bool(f[3])})
        if not cals:
            raise OutlookError("Outlook returned no calendars. Is it in Legacy Outlook mode?")
        return cals

    def list_events(self, start: str | None = None, end: str | None = None, calendar_id: str | None = None) -> dict:
        start_dt, end_dt = event_range(start, end)
        cals = self.list_calendars()
        by_id = {c["id"]: c for c in cals}
        if calendar_id is not None:
            calendar_id = check_id(calendar_id, "calendar_id")
            if calendar_id not in by_id:
                raise OutlookError(f"No calendar with id {calendar_id}. Use list_calendars.")
            ids = [calendar_id]
        else:
            ids = list(by_id)
        raw = self._run("read/events", [to_stamp(start_dt), to_stamp(end_dt), *ids])
        events, problems = [], []
        rows = records(raw)
        # Changed occurrences (exceptions) are events of their own; the series' expansion
        # still has them at their original time, so those are left out.
        changed = {(f[13], f[14]) for f in rows if len(f) > 14 and f[12] == "true" and f[13] and f[14]}
        for f in rows:
            kind, eid, cid = f[0], f[1], f[2]
            base = {
                "id": eid,
                "calendar": by_id.get(cid, {}).get("name"),
                "calendar_id": cid,
                "account": by_id.get(cid, {}).get("account"),
                "subject": f[3],
                "location": f[6],
                "all_day": as_bool(f[7]),
                "organizer": f[8],
                "free_busy": enum(f[9]),
            }
            if kind == "single":
                tz = f[11] if len(f) > 11 else ""
                is_exception = len(f) > 13 and f[12] == "true" and bool(f[13])
                ev = dates_if_all_day({**base, "start": stamp_iso(f[4]), "end": stamp_iso(f[5]),
                                       "recurring": is_exception}, tz)
                if is_exception:
                    ev["series_id"] = f[13]
                events.append(ev)
                continue
            skip = {stamp_iso(o) for (sid, o) in changed if sid == eid}
            try:
                occurrences = expand_series(f[10], start_dt, end_dt)
            except Exception as exc:  # noqa: BLE001 - a bad series must not hide the rest of the calendar
                problems.append({"id": eid, "subject": f[3], "error": f"could not expand recurrence: {exc}"})
                continue
            for s, e in occurrences:
                if iso(s) in skip:
                    continue
                # iCalendar all-day dates arrive as local midnights: no zone, nearest midnight.
                ev = dates_if_all_day({**base, "start": iso(s), "end": iso(e), "recurring": True})
                ev["occurrence_start"] = ev["start"]  # pass back to read_event / update_event / respond ...
                events.append(ev)
        events.sort(key=lambda e: e["start"] or "")
        result = {
            "start": iso(start_dt),
            "end": iso(end_dt),
            "count": len(events),
            "events": events[:MAX_EVENTS],
            "truncated": len(events) > MAX_EVENTS,
        }
        if problems:
            result["problems"] = problems
        return result

    def _read(self, event_id: str) -> dict:
        event_id = check_id(event_id, "event_id")
        self.require_accounts()
        ev = parse_event(self._run("read/event", [event_id]))
        own = self._own_addresses()
        organizer = (ev["organizer"] or "").strip().casefold()
        ev["is_organizer"] = not organizer or any(a in organizer for a in own)
        mine = next((a for a in ev["attendees"] if (a["address"] or "").casefold() in own), None)
        if ev["is_organizer"]:
            ev["my_response"] = "organizer" if ev["attendees"] else "none"
        elif mine is not None:
            ev["my_response"] = STATUS_WORDS.get(mine["status"] or "", "not responded")
        else:
            ev["my_response"] = "not responded"
        return ev

    def _find_occurrence(self, master: dict, occurrence_start: str) -> tuple[datetime, datetime]:
        """Start and end of the occurrence of `master` that starts at occurrence_start."""
        if not master["is_recurring"]:
            raise OutlookError("That event does not repeat, so it has no occurrences; leave occurrence_start out.")
        when = parse_when(occurrence_start, "occurrence_start")
        if not master["_ics"]:
            raise OutlookError("Outlook did not give the series' recurrence data; open the occurrence in Outlook.")
        for s, e in expand_series(master["_ics"], when - timedelta(days=1), when + timedelta(days=2)):
            same = s.date() == when.date() if master["all_day"] else abs((s - when).total_seconds()) < 60
            if same:
                return s, e
        raise OutlookError(f"The series has no occurrence starting at {occurrence_start}. Use the start that "
                           "list_events shows for it.")

    @staticmethod
    def _public(ev: dict) -> dict:
        return {k: v for k, v in ev.items() if not k.startswith("_")}

    def read_event(self, event_id: str, occurrence_start: str | None = None) -> dict:
        ev = self._read(event_id)
        ev["occurrence_start"] = None
        if occurrence_start:
            s, e = self._find_occurrence(ev, occurrence_start)
            occ = dates_if_all_day({"all_day": ev["all_day"], "start": iso(s), "end": iso(e)})
            ev["start"], ev["end"] = occ["start"], occ["end"]
            ev["occurrence_start"] = ev["start"]
        return self._public(ev)

    @staticmethod
    def _event_pairs(subject, location, body, all_day, free_busy, reminder_minutes, remove_reminder,
                     is_private=None) -> list[str]:
        pairs: list[str] = []
        if subject is not None:
            pairs += ["subject", subject]
        if location is not None:
            pairs += ["location", location]
        if body is not None:
            pairs += ["html", to_html(body)]
        if all_day is not None:
            pairs += ["allday", "1" if all_day else "0"]
        if free_busy is not None:
            pairs += ["freebusy", check_free_busy(free_busy)]
        if remove_reminder:
            pairs += ["reminder", ""]
        elif check_reminder(reminder_minutes) is not None:
            pairs += ["reminder", str(int(reminder_minutes))]
        if is_private is not None:
            pairs += ["private", "1" if is_private else "0"]
        return pairs

    @staticmethod
    def _recurrence_value(rec: dict, start: datetime) -> str:
        kind = {"daily": "daily", "weekly": "weekly", "monthly": "absolute monthly",
                "yearly": "absolute yearly"}[rec["frequency"]]
        if rec.get("until"):
            end_type, data = "date", to_stamp(datetime.combine(date.fromisoformat(rec["until"]), time.min).astimezone())
        elif rec.get("count"):
            end_type, data = "count", str(rec["count"])
        else:
            end_type, data = "none", ""
        return US.join([kind, str(rec["interval"]), GS.join(rec.get("days_of_week", [])),
                        str(rec.get("day_of_month", "")), str(rec.get("month", "")), end_type, data, to_stamp(start)])

    def _calendar_account(self, calendar_id: str | None) -> dict | None:
        cals = self.list_calendars()
        if calendar_id is None:
            return None
        cal = next((c for c in cals if c["id"] == calendar_id), None)
        if cal is None:
            raise OutlookError(f"No calendar with id {calendar_id}. Use list_calendars.")
        return cal

    def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        calendar_id: str | None = None,
        location: str | None = None,
        body: str | None = None,
        all_day: bool | None = None,
        free_busy: str | None = None,
        reminder_minutes: int | None = None,
        attendees=(),
        optional_attendees=(),
        send_invitations: bool = False,
        recurrence: dict | None = None,
        categories=None,
        is_private: bool | None = None,
    ) -> dict:
        s, e = check_event_times(start, end, all_day)
        req, opt = check_attendees(attendees, optional_attendees)
        rec = check_recurrence(recurrence, s)
        cats = check_categories(categories)
        if send_invitations and not (req or opt):
            raise OutlookError("send_invitations needs attendees.")
        self.require_accounts()
        if calendar_id is not None:
            calendar_id = check_id(calendar_id, "calendar_id")
            cal = self._calendar_account(calendar_id)
            if (req or opt) and cal["account"] in {f.account for f in self.folders() if f.is_google}:
                raise OutlookError("Meetings can't be sent from a Google calendar through Outlook's scripting; "
                                   "use an Exchange calendar, or leave attendees out.")
        pairs = self._event_pairs(subject, location, body, all_day, free_busy, reminder_minutes, False, is_private)
        if rec:
            pairs += ["recurrence", self._recurrence_value(rec, s)]
        if req:
            pairs += ["addreq", GS.join(req)]
        if opt:
            pairs += ["addopt", GS.join(opt)]
        if cats:
            pairs += ["categories", GS.join(self._ensure_categories(cats))]
        eid = self._run("write/event", ["create", calendar_id or "", to_stamp(s), to_stamp(e), *pairs]).strip()
        result = self.read_event(eid)
        if req or opt:
            if send_invitations:
                self._run("write/event", ["send", eid, ""])
            result["invitations_sent"] = bool(send_invitations)
            if not send_invitations:
                result["note"] = ("Saved with attendees but not sent. Send it with "
                                  "update_event(event_id, send_update=true).")
        return result

    def update_event(
        self,
        event_id: str,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
        all_day: bool | None = None,
        free_busy: str | None = None,
        reminder_minutes: int | None = None,
        remove_reminder: bool = False,
        whole_series: bool = False,
        occurrence_start: str | None = None,
        attendees=None,
        optional_attendees=None,
        categories=None,
        is_private: bool | None = None,
        send_update: bool = False,
    ) -> dict:
        event_id = check_id(event_id, "event_id")
        cats = check_categories(categories)
        current = self._read(event_id)
        self._check_live(current)
        if (attendees is not None or optional_attendees is not None or send_update) and not current["is_organizer"]:
            raise OutlookError(NOT_ORGANIZER)
        occ = None
        if occurrence_start:
            occ = self._find_occurrence(current, occurrence_start)
            old_s, old_e = occ
        elif current["is_recurring"] and not current["is_occurrence"]:
            if start or end or all_day is not None:
                raise OutlookError(RECURRING_TIMES)
            if not whole_series:
                raise OutlookError(RECURRING_SERIES)
        if not occ:
            old_s, old_e = event_bound(current["start"]), event_bound(current["end"])
        new_s = parse_when(start, "start") if start else None
        new_e = parse_when(end, "end") if end else None
        final_s, final_e = new_s or old_s, new_e or old_e
        if final_s and final_e and final_e <= final_s:
            raise OutlookError("After this change the event would end before it starts; pass both start and end.")
        times: list[str] = []
        if new_s and new_e:
            # Moving later: push the end out first so start never passes end.
            if old_e and new_s >= old_e:
                times = ["end", to_stamp(new_e), "start", to_stamp(new_s)]
            else:
                times = ["start", to_stamp(new_s), "end", to_stamp(new_e)]
        elif new_s:
            times = ["start", to_stamp(new_s)]
        elif new_e:
            times = ["end", to_stamp(new_e)]
        pairs = times + self._event_pairs(subject, location, body, all_day, free_busy, reminder_minutes,
                                          remove_reminder, is_private)
        pairs += self._attendee_pairs(current, attendees, optional_attendees)
        if cats is not None:
            pairs += ["categories", GS.join(self._ensure_categories(cats))]
        if not pairs and not send_update:
            raise OutlookError("Nothing to change.")
        target = event_id
        if pairs:
            if occ:
                target = self._run("write/event", ["occurrence", event_id, to_stamp(occ[0]), *pairs]).strip()
            else:
                self._run("write/event", ["update", event_id, *pairs])
        after = self.read_event(target)
        if occ:
            after["occurrence_start"] = after["start"]
        if after["attendees"] and current["is_organizer"]:
            if send_update:
                self._run("write/event", ["send", target, ""])
                after["update_sent"] = True
            else:
                after["update_sent"] = False
                after["note"] = ("Changed your copy only; the attendees were not told. "
                                 "Pass send_update=true to send them the update.")
        elif send_update:
            raise OutlookError("send_update needs a meeting with attendees; this event has none.")
        return after

    @staticmethod
    def _attendee_pairs(current: dict, attendees, optional_attendees) -> list[str]:
        if attendees is None and optional_attendees is None:
            return []
        have = {"required": set(), "optional": set()}
        for a in current["attendees"]:
            kind = "optional" if (a.get("type") or "").startswith("optional") else "required"
            have[kind].add((a["address"] or "").casefold())
        req, opt = check_attendees(attendees or [], optional_attendees or [])
        want_req = {x.casefold() for x in req} if attendees is not None else have["required"]
        want_opt = {x.casefold() for x in opt} if optional_attendees is not None else have["optional"]
        if (have["required"] - want_req) or (have["optional"] - want_opt):
            raise OutlookError(NO_ATTENDEE_REMOVAL)
        add_req = [x for x in req if x.casefold() not in have["required"]] if attendees is not None else []
        add_opt = [x for x in opt if x.casefold() not in have["optional"]] if optional_attendees is not None else []
        pairs: list[str] = []
        if add_req:
            pairs += ["addreq", GS.join(add_req)]
        if add_opt:
            pairs += ["addopt", GS.join(add_opt)]
        return pairs

    def _check_live(self, ev: dict) -> None:
        """The event is in one of the calendars list_calendars shows, not in Deleted Items
        or another folder, where Outlook's delete would be permanent."""
        if ev["calendar_id"] not in {c["id"] for c in self.list_calendars()}:
            raise OutlookError(
                f"Event {ev['id']} is not in one of your calendars (it may be in Deleted Items); nothing was changed."
            )

    def _own_addresses(self) -> set[str]:
        return {f.account.casefold() for f in self.folders() if "@" in f.account}

    def delete_event(self, event_id: str, whole_series: bool = False, occurrence_start: str | None = None) -> dict:
        event_id = check_id(event_id, "event_id")
        ev = self._read(event_id)
        self._check_live(ev)
        organizer = (ev["organizer"] or "").strip().casefold()
        if ev["attendees"] or (organizer and organizer not in self._own_addresses()):
            # Invitations synced from other services can list no attendees; an organizer
            # who isn't you still makes it someone's meeting.
            raise OutlookError(HAS_ATTENDEES)
        if occurrence_start or ev["is_occurrence"]:
            raise OutlookError(NO_OCCURRENCE_DELETE)
        if ev["is_recurring"] and not whole_series:
            raise OutlookError(RECURRING_SERIES)
        self._run("write/event", ["delete", event_id])
        return {"deleted": event_id, "subject": ev["subject"], "start": ev["start"]}

    def cancel_meeting(self, event_id: str, comment: str = "", occurrence_start: str | None = None) -> dict:
        event_id = check_id(event_id, "event_id")
        ev = self._read(event_id)
        self._check_live(ev)
        if not ev["is_organizer"]:
            raise OutlookError(NOT_ORGANIZER)
        if not ev["attendees"]:
            raise OutlookError("This event has no attendees, so there is no one to tell; use delete_event.")
        occ = self._find_occurrence(ev, occurrence_start)[0] if occurrence_start else None
        self._run("write/cancel", [event_id, to_stamp(occ) if occ else "", to_html(comment) if comment else ""])
        return {"cancelled": event_id, "subject": ev["subject"],
                "occurrence_start": iso(occ) if occ else None,
                "note": "Outlook sent the cancellation to the attendees and took the meeting off your calendar."}

    def respond_to_meeting(self, event_id: str, response: str, comment: str = "", send_response: bool = True,
                           occurrence_start: str | None = None) -> dict:
        the_id = check_id(event_id, "event_id")
        check_response(response)
        self.require_accounts()
        ev = self._try(lambda: self._read(the_id))
        msg = self._try(lambda: self.read_message(the_id, max_chars=1))
        if msg is not None and not msg.get("is_meeting"):
            msg = None
        if ev is not None and msg is not None and msg.get("meeting_event_id") not in (None, the_id):
            raise OutlookError(
                f"{the_id} is both a calendar event and a meeting message on Outlook for Mac. Pass the event id "
                f"({msg['meeting_event_id']}) from read_message's meeting_event_id, or the event id from list_events."
            )
        if ev is None and msg is not None and msg.get("meeting_event_id"):
            ev = self._read(msg["meeting_event_id"])
        if ev is not None:
            if ev["is_organizer"]:
                raise OutlookError("You organize this meeting; there is nothing to respond to. "
                                   "Use update_event or cancel_meeting.")
            occ = self._find_occurrence(ev, occurrence_start)[0] if occurrence_start else None
            self._run("write/respond", ["event", ev["id"], to_stamp(occ) if occ else "", response, comment or "",
                                        "1" if send_response else "0"])
            return {"id": ev["id"], "response": response, "organizer_told": send_response,
                    "occurrence_start": iso(occ) if occ else None}
        if msg is not None:
            if occurrence_start:
                raise OutlookError("occurrence_start goes with an event id, not a meeting message.")
            self._run("write/respond", ["message", the_id, "", response, comment or "",
                                        "1" if send_response else "0"])
            return {"id": the_id, "response": response, "organizer_told": send_response}
        raise OutlookError(f"No calendar event or meeting message with id {the_id}.")

    @staticmethod
    def _try(fn):
        try:
            return fn()
        except OutlookError:
            return None
