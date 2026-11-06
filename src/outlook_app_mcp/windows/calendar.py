"""Windows calendar: calendars, events (Outlook expands recurring series itself), and changes."""

from __future__ import annotations

from ..common import (
    ALREADY_DELETED,
    DELETE_NOTE,
    HAS_ATTENDEES,
    MAX_EVENTS,
    NOT_ORGANIZER,
    RECURRING_SERIES,
    RECURRING_TIMES,
    WEEKDAYS,
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
)
from .com import (
    ATTENDEE_TYPES,
    BUSY_CODES,
    BUSY_STATUS,
    DAY_MASK,
    FOLDER_CALENDAR,
    ITEM_APPOINTMENT,
    MEETING_CANCELED,
    MEETING_NONE,
    MEETING_ORGANIZED,
    MEETING_RESPONSES,
    MY_RESPONSE,
    RECURRENCE_TYPES,
    RECURS_DAILY,
    RECURS_MONTH_NTH,
    RECURS_MONTHLY,
    RECURS_WEEKLY,
    RECURS_YEAR_NTH,
    RECURS_YEARLY,
    RESPONSE_ORGANIZER,
    RESPONSE_STATUS,
    SENSITIVITY_PRIVATE,
    categories,
    com_errors,
    com_iso,
    day_to_com,
    each,
    from_com,
    jet_local,
    make_id,
    prop,
    recipient_address,
    to_com,
    walk_folders,
)

MAX_SCAN = 20000  # occurrences looked at per calendar; a guard against endless series
DISCARD = 1  # OlInspectorClose.olDiscard
SENSITIVITY_NORMAL = 0
RECIPIENT_REQUIRED, RECIPIENT_OPTIONAL = 1, 2
NO_ONE_LEFT = "That would leave the meeting with no attendees; to call it off use cancel_meeting."
NO_OCCURRENCE_DELETE = (
    "A single occurrence of a series can't be removed from here: in classic Outlook that is only possible by "
    "deleting it for good (it does not go to Deleted Items), and this server never deletes permanently. Delete "
    "the whole series with whole_series=true, or remove this occurrence in Outlook."
)
UPDATE_NOT_SENT = (
    "Saved in your calendar only; the attendees were not told. Call update_event with send_update=true to send "
    "the update."
)


class WinCalendarMixin:
    """Needs self._app, self._ns and the store helpers of WinMailMixin."""

    _jet = staticmethod(jet_local)

    def _calendars(self) -> list[tuple[dict, object]]:
        self.require_accounts()
        try:
            default = self._ns.GetDefaultFolder(FOLDER_CALENDAR).EntryID
        except Exception:  # noqa: BLE001 - no default calendar
            default = ""
        out = []
        for st, store in self._open_stores():
            key = store.key
            for f in walk_folders(st.GetRootFolder(), ITEM_APPOINTMENT):
                info = {
                    "id": make_id(key, f.EntryID),
                    "account": store.account,
                    "name": prop(f, "Name", "") or "",
                    "is_default": f.EntryID == default,
                }
                out.append((info, f))
        if not out:
            raise OutlookError("Outlook shows no calendars.")
        return out

    @com_errors
    def list_calendars(self) -> list[dict]:
        return [info for info, _ in self._calendars()]

    def _event_summary(self, it, key: str, cal: dict) -> dict:
        ev = {
            "id": make_id(key, it.EntryID),
            "calendar": cal["name"],
            "calendar_id": cal["id"],
            "account": cal["account"],
            "subject": prop(it, "Subject", "") or "",
            "location": prop(it, "Location", "") or "",
            "all_day": bool(prop(it, "AllDayEvent", False)),
            "organizer": prop(it, "Organizer", "") or "",
            "free_busy": BUSY_STATUS.get(prop(it, "BusyStatus", None), ""),
            "start": com_iso(prop(it, "Start")),
            "end": com_iso(prop(it, "End")),
            "recurring": bool(prop(it, "IsRecurring", False)),
        }
        dates_if_all_day(ev)
        if ev["recurring"]:
            ev["occurrence_start"] = ev["start"]  # what to pass back to read_event / update_event / respond ...
        return ev

    @com_errors
    def list_events(self, start: str | None = None, end: str | None = None, calendar_id: str | None = None) -> dict:
        start_dt, end_dt = event_range(start, end)
        cals = self._calendars()
        if calendar_id is not None:
            cals = [c for c in cals if c[0]["id"] == calendar_id.strip()]
            if not cals:
                raise OutlookError(f"No calendar with id {calendar_id}. Use list_calendars.")
        events, problems = [], []
        for info, folder in cals:
            key = info["id"].split(".", 1)[0]
            items = folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            found = items.Restrict(f"[Start] < '{self._jet(end_dt)}' AND [End] > '{self._jet(start_dt)}'")
            it, seen = found.GetFirst(), 0
            while it is not None:
                seen += 1
                if seen > MAX_SCAN:
                    problems.append({"calendar_id": info["id"], "error": f"stopped after {MAX_SCAN} items"})
                    break
                s, e = from_com(prop(it, "Start")), from_com(prop(it, "End"))
                if s is not None and s >= end_dt:
                    break  # sorted by start: nothing later can be in range
                if s is not None and e is not None and e > start_dt:
                    events.append(self._event_summary(it, key, info))
                it = found.GetNext()
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

    def _event_item(self, event_id: str):
        """A calendar item, not a message or contact (their ids look the same)."""
        it, key = self._item(event_id, "event_id")
        if not str(prop(it, "MessageClass", "") or "").startswith("IPM.Appointment"):
            raise OutlookError(f"{event_id} is not a calendar event ({prop(it, 'MessageClass', '?')}).")
        return it, key

    # -- occurrences, recurrence, organizer ---------------------------------------

    def _occurrence(self, master, occurrence_start: str):
        """One occurrence of a recurring event, the way Outlook hands it out: an appointment of its own that
        becomes an exception once it is changed and saved. `occurrence_start` is the start list_events shows."""
        if not prop(master, "IsRecurring", False):
            raise OutlookError("That event does not repeat, so it has no occurrences; leave occurrence_start out.")
        when = parse_when(occurrence_start, "occurrence_start")
        try:
            return master.GetRecurrencePattern().GetOccurrence(to_com(when))
        except self._com_error as exc:
            raise OutlookError(f"The series has no occurrence starting at {occurrence_start}. Use the start that "
                               "list_events shows for it.") from exc

    @staticmethod
    def _recurrence_of(pattern) -> dict:
        """A RecurrencePattern as the dict create_event takes. A series that ends on a date and one that ends
        after a count look alike to the object model, so the end is always given as a date."""
        kind = prop(pattern, "RecurrenceType", RECURS_DAILY)
        interval = int(prop(pattern, "Interval", 1) or 1)
        if kind == RECURS_WEEKLY:
            mask = int(prop(pattern, "DayOfWeekMask", 0) or 0)
            out = {"frequency": "weekly", "interval": interval,
                   "days_of_week": [d for d in WEEKDAYS if mask & DAY_MASK[d]]}
        elif kind in (RECURS_MONTHLY, RECURS_MONTH_NTH):
            out = {"frequency": "monthly", "interval": interval}
        elif kind in (RECURS_YEARLY, RECURS_YEAR_NTH):
            out = {"frequency": "yearly", "interval": max(1, interval // 12)}
        else:
            out = {"frequency": "daily", "interval": interval}
        if kind in (RECURS_MONTHLY, RECURS_YEARLY):
            out["day_of_month"] = prop(pattern, "DayOfMonth", None)
        elif kind in (RECURS_MONTH_NTH, RECURS_YEAR_NTH):
            out["pattern"] = "on the Nth weekday of the month (can be read here, only set in Outlook)"
        if kind in (RECURS_YEARLY, RECURS_YEAR_NTH):
            out["month"] = prop(pattern, "MonthOfYear", None)
        if not prop(pattern, "NoEndDate", True):
            end = from_com(prop(pattern, "PatternEndDate"))
            if end is not None:
                out["until"] = end.date().isoformat()
        return out

    @staticmethod
    def _apply_recurrence(it, rec: dict, start) -> None:
        """Make a new event repeat. The event's own start and end must be set before, and the item saved after."""
        pattern = it.GetRecurrencePattern()
        pattern.RecurrenceType = RECURRENCE_TYPES[rec["frequency"]]
        pattern.Interval = rec["interval"] * (12 if rec["frequency"] == "yearly" else 1)
        if rec["frequency"] == "weekly":
            pattern.DayOfWeekMask = sum(DAY_MASK[d] for d in rec["days_of_week"])
        if rec["frequency"] in ("monthly", "yearly"):
            pattern.DayOfMonth = rec["day_of_month"]
        if rec["frequency"] == "yearly":
            pattern.MonthOfYear = rec["month"]
        pattern.PatternStartDate = day_to_com(start)
        if "until" in rec:
            pattern.PatternEndDate = day_to_com(parse_when(rec["until"]))
        elif "count" in rec:
            pattern.Occurrences = rec["count"]
        else:
            pattern.NoEndDate = True

    def _my_names(self) -> set[str]:
        names = set()
        for acc in each(self._ns.Accounts):
            for attr in ("SmtpAddress", "DisplayName", "UserName"):
                v = (prop(acc, attr, "") or "").strip().casefold()
                if v:
                    names.add(v)
        return names

    def _is_organizer(self, it) -> bool:
        """True for a plain appointment (it is yours), and for a meeting you organize."""
        if prop(it, "ResponseStatus", None) == RESPONSE_ORGANIZER:
            return True
        status = prop(it, "MeetingStatus", MEETING_NONE)
        if status in (MEETING_NONE, MEETING_ORGANIZED, MEETING_CANCELED):
            return True
        return (prop(it, "Organizer", "") or "").strip().casefold() in self._my_names()

    def _event_dict(self, it, master, key: str, occurrence: bool = False) -> dict:
        attendees = []
        for r in each(it.Recipients):
            status = RESPONSE_STATUS.get(prop(r, "MeetingResponseStatus", None))
            attendees.append({
                "name": prop(r, "Name", "") or "",
                "address": recipient_address(r),
                "type": ATTENDEE_TYPES.get(prop(r, "Type", None)),
                "status": status,
            })
        reminder = prop(it, "ReminderMinutesBeforeStart", None) if prop(it, "ReminderSet", False) else None
        recurring = bool(prop(it, "IsRecurring", False))
        ev = {
            "id": make_id(key, master.EntryID),
            "calendar_id": make_id(key, master.Parent.EntryID),
            "subject": prop(it, "Subject", "") or "",
            "start": com_iso(prop(it, "Start")),
            "end": com_iso(prop(it, "End")),
            "location": prop(it, "Location", "") or "",
            "all_day": bool(prop(it, "AllDayEvent", False)),
            "is_recurring": recurring,
            "is_meeting": prop(it, "MeetingStatus", MEETING_NONE) != MEETING_NONE,
            "organizer": prop(it, "Organizer", "") or "",
            "free_busy": BUSY_STATUS.get(prop(it, "BusyStatus", None), ""),
            "is_private": prop(it, "Sensitivity", 0) == SENSITIVITY_PRIVATE,
            "reminder_minutes": int(reminder) if reminder is not None else None,
            "attendees": attendees,
            "categories": categories(it),
            "recurrence": self._recurrence_of(it.GetRecurrencePattern()) if recurring else None,
            "my_response": MY_RESPONSE.get(prop(it, "ResponseStatus", 0), "none"),
            "is_organizer": self._is_organizer(it),
            "description": prop(it, "Body", "") or "",
        }
        dates_if_all_day(ev)
        ev["occurrence_start"] = ev["start"] if occurrence else None
        return ev

    @com_errors
    def read_event(self, event_id: str, occurrence_start: str | None = None) -> dict:
        self.require_accounts()
        master, key = self._event_item(event_id)
        if occurrence_start:
            return self._event_dict(self._occurrence(master, occurrence_start), master, key, occurrence=True)
        return self._event_dict(master, master, key)

    def _set_fields(self, it, subject, location, body, all_day, free_busy, reminder_minutes, remove_reminder) -> bool:
        changed = False
        if subject is not None:
            it.Subject = subject
            changed = True
        if location is not None:
            it.Location = location
            changed = True
        if body is not None:
            it.Body = body
            changed = True
        if all_day is not None:
            it.AllDayEvent = bool(all_day)
            changed = True
        if free_busy is not None:
            it.BusyStatus = BUSY_CODES[check_free_busy(free_busy)]
            changed = True
        if remove_reminder:
            it.ReminderSet = False
            changed = True
        elif check_reminder(reminder_minutes) is not None:
            it.ReminderSet = True
            it.ReminderMinutesBeforeStart = int(reminder_minutes)
            changed = True
        return changed

    @com_errors
    def create_event(self, subject: str, start: str, end: str, calendar_id: str | None = None, location=None,
                     body=None, all_day=None, free_busy=None, reminder_minutes=None, attendees=(),
                     optional_attendees=(), send_invitations: bool = False, recurrence=None, categories=None,
                     is_private=None) -> dict:
        """Create an event. With attendees it is a meeting: it is only saved (unsent) unless
        send_invitations is true, which sends the invitations at once."""
        s, e = check_event_times(start, end, all_day)
        if free_busy is not None:
            check_free_busy(free_busy)
        check_reminder(reminder_minutes)
        required, optional = check_attendees(attendees, optional_attendees)
        rec = check_recurrence(recurrence, s)
        cats = check_categories(categories)
        if send_invitations and not (required or optional):
            raise OutlookError("send_invitations needs attendees.")
        cals = self._calendars()
        if calendar_id is not None:
            chosen = [c for c in cals if c[0]["id"] == calendar_id.strip()]
            if not chosen:
                raise OutlookError(f"No calendar with id {calendar_id}. Use list_calendars.")
        else:
            chosen = [c for c in cals if c[0]["is_default"]] or cals[:1]
        info, folder = chosen[0]
        it = folder.Items.Add(ITEM_APPOINTMENT)
        it.Start = to_com(s)
        it.End = to_com(e)
        self._set_fields(it, subject, location, body, all_day, free_busy, reminder_minutes, False)
        if is_private is not None:
            it.Sensitivity = SENSITIVITY_PRIVATE if is_private else SENSITIVITY_NORMAL
        if cats:
            self._apply_categories(it, cats)
        if required or optional:
            it.MeetingStatus = MEETING_ORGANIZED
            for addr_list, kind in ((required, RECIPIENT_REQUIRED), (optional, RECIPIENT_OPTIONAL)):
                for a in addr_list:
                    it.Recipients.Add(a).Type = kind
            resolved = it.Recipients.ResolveAll()
            if send_invitations:
                self._check_resolved(it, resolved)  # before anything is saved: nothing is left behind
        if rec:
            self._apply_recurrence(it, rec, s)
        it.Save()
        if send_invitations:
            it.Send()
        key = info["id"].split(".", 1)[0]
        result = self.read_event(make_id(key, it.EntryID))
        if required or optional:
            result["invitations_sent"] = bool(send_invitations)
            if not send_invitations:
                result["note"] = ("Saved without sending: the attendees have not been invited. "
                                  "update_event with send_update=true sends the invitations.")
        return result

    @staticmethod
    def _attendees_of(it, kind: int) -> list[str]:
        return [recipient_address(r) for r in each(it.Recipients) if prop(r, "Type", None) == kind]

    @com_errors
    def update_event(self, event_id: str, subject=None, start=None, end=None, location=None, body=None, all_day=None,
                     free_busy=None, reminder_minutes=None, remove_reminder: bool = False,
                     whole_series: bool = False, occurrence_start=None, attendees=None, optional_attendees=None,
                     categories=None, is_private=None, send_update: bool = False) -> dict:
        self.require_accounts()
        master, key = self._event_item(event_id)
        recurring_series = bool(prop(master, "IsRecurring", False)) and not occurrence_start
        if recurring_series:
            if start or end or all_day is not None:
                raise OutlookError(RECURRING_TIMES)
            if not whole_series:
                raise OutlookError(RECURRING_SERIES)
        if free_busy is not None:
            check_free_busy(free_busy)
        check_reminder(reminder_minutes)
        cats = check_categories(categories)
        required = None if attendees is None else check_attendees(attendees, [])[0]
        optional = None if optional_attendees is None else check_attendees([], optional_attendees)[1]
        if self._under_deleted(master, key):
            raise OutlookError("That event is in Deleted Items; restore it in Outlook before changing it.")
        it = self._occurrence(master, occurrence_start) if occurrence_start else master
        current = self._event_dict(it, master, key, occurrence=bool(occurrence_start))
        old_s, old_e = event_bound(current["start"]), event_bound(current["end"])
        new_s = parse_when(start, "start") if start else None
        new_e = parse_when(end, "end") if end else None
        final_s, final_e = new_s or old_s, new_e or old_e
        if final_s and final_e and final_e <= final_s:
            raise OutlookError("After this change the event would end before it starts; pass both start and end.")
        # Attendees: only the organizer changes who is invited or sends an update.
        have_req, have_opt = self._attendees_of(it, RECIPIENT_REQUIRED), self._attendees_of(it, RECIPIENT_OPTIONAL)
        new_req = have_req if required is None else required
        new_opt = have_opt if optional is None else optional
        people_changed = ({a.casefold() for a in new_req} != {a.casefold() for a in have_req}
                          or {a.casefold() for a in new_opt} != {a.casefold() for a in have_opt})
        if (people_changed or send_update) and not self._is_organizer(it):
            raise OutlookError(NOT_ORGANIZER)
        if (have_req or have_opt) and not (new_req or new_opt):
            raise OutlookError(NO_ONE_LEFT)
        if send_update and not (new_req or new_opt):
            raise OutlookError("This event has no attendees, so there is nobody to send an update to.")
        changed = False
        if new_s and new_e:
            # Moving later: push the end out first so start never passes end.
            if old_e and new_s >= old_e:
                it.End = to_com(new_e)
                it.Start = to_com(new_s)
            else:
                it.Start = to_com(new_s)
                it.End = to_com(new_e)
            changed = True
        elif new_s:
            it.Start = to_com(new_s)
            changed = True
        elif new_e:
            it.End = to_com(new_e)
            changed = True
        changed = self._set_fields(it, subject, location, body, all_day, free_busy, reminder_minutes,
                                   remove_reminder) or changed
        if is_private is not None:
            it.Sensitivity = SENSITIVITY_PRIVATE if is_private else SENSITIVITY_NORMAL
            changed = True
        if cats is not None:
            self._apply_categories(it, cats)
            changed = True
        if people_changed:
            # The address string makes Outlook rebuild just that kind of attendee; resources are untouched.
            if {a.casefold() for a in new_req} != {a.casefold() for a in have_req}:
                it.RequiredAttendees = "; ".join(new_req)
            if {a.casefold() for a in new_opt} != {a.casefold() for a in have_opt}:
                it.OptionalAttendees = "; ".join(new_opt)
            if prop(it, "MeetingStatus", MEETING_NONE) == MEETING_NONE:
                it.MeetingStatus = MEETING_ORGANIZED
            resolved = it.Recipients.ResolveAll()
            if send_update:
                self._check_resolved(it, resolved)
            changed = True
        if not changed and not send_update:
            raise OutlookError("Nothing to change.")
        if send_update:
            it.Send()
        else:
            it.Save()
        if occurrence_start:
            result = self._event_dict(it, master, key, occurrence=True)
        else:
            result = self.read_event(event_id)
        if result["attendees"] and result["is_organizer"]:
            result["update_sent"] = bool(send_update)
            if changed and not send_update:
                result["note"] = UPDATE_NOT_SENT
        return result

    @com_errors
    def delete_event(self, event_id: str, whole_series: bool = False, occurrence_start: str | None = None) -> dict:
        ev = self.read_event(event_id)
        if ev["attendees"] or ev["is_meeting"]:
            raise OutlookError(HAS_ATTENDEES)
        if occurrence_start:
            raise OutlookError(NO_OCCURRENCE_DELETE)
        if ev["is_recurring"] and not whole_series:
            raise OutlookError(RECURRING_SERIES)
        it, key = self._event_item(event_id)
        if self._under_deleted(it, key):
            raise OutlookError(f"Event {event_id} {ALREADY_DELETED}")
        # Moved, never Delete(): Delete on an item in Deleted Items is permanent.
        it.Move(self._deleted_folder_obj(key))
        return {"deleted": event_id, "subject": ev["subject"], "start": ev["start"], "note": DELETE_NOTE}

    @com_errors
    def cancel_meeting(self, event_id: str, comment: str = "", occurrence_start: str | None = None) -> dict:
        """Cancel a meeting you organize: the attendees get the cancellation, then the event leaves your calendar
        (moved to Deleted Items, never deleted). One occurrence is cancelled by sending its cancellation; Outlook
        then removes that occurrence itself."""
        self.require_accounts()
        master, key = self._event_item(event_id)
        if self._under_deleted(master, key):
            raise OutlookError("That meeting is in Deleted Items; restore it in Outlook before cancelling it.")
        it = self._occurrence(master, occurrence_start) if occurrence_start else master
        status = prop(it, "MeetingStatus", MEETING_NONE)
        if status == MEETING_NONE:
            raise OutlookError("That event is not a meeting, so there is nothing to cancel; use delete_event.")
        if status == MEETING_CANCELED:
            raise OutlookError("That meeting is already cancelled.")
        if not self._is_organizer(it):
            raise OutlookError(NOT_ORGANIZER)
        subject = prop(it, "Subject", "") or ""
        entry = master.EntryID
        it.MeetingStatus = MEETING_CANCELED
        if comment:
            it.Body = comment
        it.Send()
        if not occurrence_start:
            # Outlook often files the cancelled item in Deleted Items by itself; otherwise it is moved there.
            try:
                left = self._ns.GetItemFromID(entry, self._store(key).store_id)
            except self._com_error:
                left = None
            if left is not None and not self._under_deleted(left, key):
                left.Move(self._deleted_folder_obj(key))
        result = {"cancelled": event_id, "subject": subject, "note": "The cancellation was sent to the attendees. "
                  + DELETE_NOTE}
        if occurrence_start:
            result["occurrence_start"] = occurrence_start
        return result

    def _meeting_target(self, event_id: str):
        """(appointment, store key) for respond_to_meeting. The id may also be that of a meeting-request
        message, which stands for the event it announces."""
        it, key = self._item(event_id, "event_id")
        cls = prop(it, "MessageClass", "") or ""
        if cls.startswith("IPM.Appointment"):
            return it, key
        if not cls.startswith("IPM.Schedule.Meeting"):
            raise OutlookError(f"{event_id} is not a calendar event or a meeting request ({cls or '?'}).")
        if not cls.startswith("IPM.Schedule.Meeting.Request"):
            raise OutlookError("That message is a reply or a cancellation, not an invitation; there is nothing "
                               "to respond to.")
        appt = it.GetAssociatedAppointment(False) or it.GetAssociatedAppointment(True)
        if appt is None:
            raise OutlookError("Outlook has no calendar event for that meeting request.")
        return appt, self._key_of(appt)

    @com_errors
    def respond_to_meeting(self, event_id: str, response: str, comment: str = "", send_response: bool = True,
                           occurrence_start: str | None = None) -> dict:
        check_response(response)
        self.require_accounts()
        master, key = self._meeting_target(event_id)
        it = self._occurrence(master, occurrence_start) if occurrence_start else master
        status = prop(it, "MeetingStatus", MEETING_NONE)
        if status == MEETING_NONE:
            raise OutlookError("That event is not a meeting; there is nothing to respond to.")
        if self._is_organizer(it):
            raise OutlookError("You organize this meeting, so there is nobody to respond to; use cancel_meeting "
                               "to call it off.")
        if self._under_deleted(master, key):
            raise OutlookError("That meeting is in Deleted Items; restore it in Outlook before responding.")
        reply = it.Respond(MEETING_RESPONSES[response], True)
        if send_response and reply is not None:
            if comment:
                reply.Body = comment
            reply.Send()
        elif reply is not None:
            reply.Close(DISCARD)
        # Accepting or tentatively accepting replaces the appointment with a copy that has a
        # new EntryID; Outlook says to read EntryID again. A declined meeting leaves the calendar.
        result = {"response": response, "organizer_told": bool(send_response and reply is not None)}
        if response == "decline" and not occurrence_start:
            result["id"] = None
            result["note"] = "Outlook removes a declined meeting from the calendar."
        else:
            new_eid = prop(master, "EntryID", None)
            result["id"] = make_id(key, new_eid) if new_eid else None
            if response == "decline":
                result["note"] = "Outlook removes a declined occurrence from the calendar."
            elif result["id"] != event_id:
                result["note"] = "Outlook replaced the event with a copy; use the new id from now on."
        return result
