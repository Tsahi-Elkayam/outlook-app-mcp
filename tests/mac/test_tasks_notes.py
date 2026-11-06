"""Tasks and notes: listing, reading, writing and the Exchange-only delete."""

import pytest
from conftest import rows

from outlook_app_mcp.common import to_html
from outlook_app_mcp.mac.osa import FS, GS, RS, OutlookError

NOT_FLAGGED = "«constant ****FlNF»"
DONE = "«constant ****FlCo»"
OPEN = "«constant ****FlNC»"
HIGH = "«constant ****PrHi»"
NORMAL = "«constant ****PrNr»"

TASK_FOLDERS = (
    ("F", 20, "Tasks", "me@work.example.com", "exchange"),
    ("F", 21, "Tasks", "me@school.example", "imap"),
    ("F", 22, "Tasks", "", ""),
)


def task(tid, subject, due="", start="", done="", flag=OPEN, prio=NORMAL, folder=20, cats=()):
    return ("T", tid, subject, due, start, done, flag, prio, folder, GS.join(cats))


TASKS = rows(
    *TASK_FOLDERS,
    task(1, "Pay rent", due="2026-11-12-0", prio=HIGH, cats=("Home",)),
    task(2, "Call Rina", due="2026-11-09-0"),
    task(3, "Someday", folder=21),
    task(4, "File taxes", due="2026-10-30-0", done="2026-11-01-0", flag=DONE, cats=("Money", "Home")),
    task(5, "Alpha task", due="2026-11-12-0"),
    task(6, "Local only", folder=22, due="2026-12-01-0"),
)


def task_raw(tid="9", subject="Pay rent", due="2026-11-12-0", reminder="", body="remember", folder=20,
             flag=OPEN, done="", cats=()):
    header = ["T", tid, subject, due, "", done, flag, HIGH, str(folder), GS.join(cats), reminder]
    return FS.join(header) + RS + body


# -- tasks: listing --------------------------------------------------------------------


def test_list_tasks_sorts_by_due_date_then_subject_with_undated_last(make):
    o, _ = make({"read/tasks": TASKS})
    res = o.list_tasks()
    assert [t["id"] for t in res["tasks"]] == ["2", "5", "1", "6", "3"]
    assert res["count"] == 5 and res["returned"] == 5 and res["has_more"] is False


def test_list_tasks_fields(make):
    o, _ = make({"read/tasks": TASKS})
    t = next(t for t in o.list_tasks()["tasks"] if t["id"] == "1")
    assert t == {"id": "1", "subject": "Pay rent", "due": "2026-11-12", "start": None, "completed": False,
                 "completed_on": None, "priority": "high", "categories": ["Home"], "folder": "Tasks",
                 "account": "me@work.example.com"}
    local = next(t for t in o.list_tasks()["tasks"] if t["id"] == "6")
    assert local["account"] is None and local["priority"] == "normal"


def test_list_tasks_hides_completed_unless_asked(make):
    o, _ = make({"read/tasks": TASKS})
    assert "4" not in [t["id"] for t in o.list_tasks()["tasks"]]
    done = next(t for t in o.list_tasks(include_completed=True)["tasks"] if t["id"] == "4")
    assert done["completed"] is True and done["completed_on"] == "2026-11-01"
    assert o.list_tasks(include_completed=True)["count"] == 6


def test_list_tasks_query_matches_subject_or_category(make):
    o, _ = make({"read/tasks": TASKS})
    assert [t["id"] for t in o.list_tasks(query="RENT")["tasks"]] == ["1"]
    assert [t["id"] for t in o.list_tasks(query="home", include_completed=True)["tasks"]] == ["4", "1"]
    assert o.list_tasks(query="  ")["count"] == 5


def test_list_tasks_due_before_excludes_undated_and_the_day_itself(make):
    o, _ = make({"read/tasks": TASKS})
    assert [t["id"] for t in o.list_tasks(due_before="2026-11-12")["tasks"]] == ["2"]
    assert [t["id"] for t in o.list_tasks(due_before="2026-11-13")["tasks"]] == ["2", "5", "1"]


def test_list_tasks_limit_and_validation(make):
    o, _ = make({"read/tasks": TASKS})
    res = o.list_tasks(limit=2)
    assert res["returned"] == 2 and res["has_more"] is True and res["count"] == 5
    with pytest.raises(OutlookError, match="limit"):
        o.list_tasks(limit=0)
    with pytest.raises(OutlookError, match="due_before"):
        o.list_tasks(due_before="tomorrow")


def test_tasks_need_legacy_outlook(make):
    o, r = make({"read/folders": "", "read/tasks": TASKS})
    for call in (o.list_tasks, lambda: o.read_task(1), lambda: o.create_task("x"), lambda: o.delete_task(1),
                 o.list_notes, lambda: o.read_note(1), lambda: o.create_note("x"), lambda: o.delete_note(1)):
        with pytest.raises(OutlookError, match="Legacy Outlook"):
            call()
    assert {c[0] for c in r.calls} == {"read/folders"}


# -- tasks: reading -----------------------------------------------------------------------


def test_read_task(make):
    raw = task_raw(reminder="2026-11-11-32400", body="line1\nline2", cats=("Home",))
    o, r = make({"read/tasks": TASKS, "read/task": raw})
    t = o.read_task("9")
    assert r.last("read/task") == ["9"]
    assert t["id"] == "9" and t["subject"] == "Pay rent" and t["due"] == "2026-11-12" and t["priority"] == "high"
    assert t["body"] == "line1\nline2" and t["reminder"].startswith("2026-11-11T09:00")
    assert t["folder"] == "Tasks" and t["account"] == "me@work.example.com" and t["categories"] == ["Home"]


def test_read_task_without_reminder_and_completed(make):
    o, _ = make({"read/tasks": TASKS, "read/task": task_raw(flag=DONE, done="2026-11-10-0")})
    t = o.read_task(9)
    assert t["reminder"] is None and t["completed"] is True and t["completed_on"] == "2026-11-10"


def test_read_task_id_must_be_a_number(make):
    o, r = make({"read/tasks": TASKS, "read/task": task_raw()})
    with pytest.raises(OutlookError, match="task_id"):
        o.read_task("abc")
    assert not any(c[0] == "read/task" for c in r.calls)


# -- tasks: writing -------------------------------------------------------------------------


def writer(**extra):
    return {"read/tasks": TASKS, "read/task": task_raw(), "write/task": "9", **extra}


def test_create_task_pairs(make):
    o, r = make(writer(**{"read/categories": rows(("Home", "")), "write/category": "x"}))
    res = o.create_task("  Pay rent ", body="remember\n\nnow", due="2026-11-12T15:30", start="2026-11-10",
                        reminder="2026-11-11T09:00", priority="HIGH", categories=["home"])
    args = r.last("write/task")
    assert args[:2] == ["create", ""]
    assert args[2:] == [
        "subject", "Pay rent", "html", to_html("remember\n\nnow"),
        "due", "2026-11-12-0",  # tasks keep days: the time of day is dropped
        "start", "2026-11-10-0",
        "reminder", "2026-11-11-32400",
        "priority", "high",
        "categories", "Home",
    ]
    assert res["id"] == "9" and r.last("read/task") == ["9"]


def test_create_task_minimal_and_in_an_account(make):
    o, r = make(writer())
    o.create_task("x")
    assert r.last("write/task") == ["create", "", "subject", "x"]
    o.create_task("x", account="school")
    assert r.last("write/task")[:2] == ["create", "21"]


def test_create_task_refusals(make):
    o, r = make(writer())
    with pytest.raises(OutlookError, match="subject"):
        o.create_task("   ")
    with pytest.raises(OutlookError, match="start must not be after due"):
        o.create_task("x", due="2026-11-10", start="2026-11-11")
    with pytest.raises(OutlookError, match="No task folder of an account matching 'nobody'"):
        o.create_task("x", account="nobody")
    with pytest.raises(OutlookError, match="importance"):
        o.create_task("x", priority="urgent")
    with pytest.raises(OutlookError, match="not an ISO date"):
        o.create_task("x", due="someday")
    assert not any(c[0] == "write/task" for c in r.calls)


def test_update_task_pairs_and_clearing(make):
    o, r = make(writer())
    o.update_task(9, subject="New", due="", start="", remove_reminder=True, complete=True, priority="low")
    assert r.last("write/task") == ["update", "9", "subject", "New", "due", "", "start", "", "reminder", "",
                                    "priority", "low", "complete", "1"]
    o.update_task(9, complete=False)
    assert r.last("write/task") == ["update", "9", "complete", "0"]
    o.update_task(9, due="2026-12-24T08:00", reminder="2026-12-23T09:00")
    assert r.last("write/task") == ["update", "9", "due", "2026-12-24-0", "reminder", "2026-12-23-32400"]


def test_update_task_remove_reminder_wins_over_a_new_one(make):
    o, r = make(writer())
    o.update_task(9, reminder="2026-12-23T09:00", remove_reminder=True)
    assert r.last("write/task") == ["update", "9", "reminder", ""]


def test_update_task_refusals(make):
    o, r = make(writer())
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_task(9)
    with pytest.raises(OutlookError, match="start must not be after due"):
        o.update_task(9, due="2026-11-01", start="2026-11-02")
    with pytest.raises(OutlookError, match="task_id"):
        o.update_task("x", subject="a")
    assert not any(c[0] == "write/task" for c in r.calls)
    o, r = make(writer(**{"read/task": OutlookError("Outlook could not find that item")}))
    with pytest.raises(OutlookError, match="could not find"):
        o.update_task(9, subject="a")
    assert not any(c[0] == "write/task" for c in r.calls)


def test_delete_task_only_in_exchange_folders(make):
    o, r = make(writer())
    res = o.delete_task(1)
    assert r.last("write/task") == ["delete", "1"]
    assert res["deleted"] == "1" and res["subject"] == "Pay rent" and "nothing was deleted permanently" in res["note"]
    for tid in (3, 6):  # IMAP folder, On My Computer
        with pytest.raises(OutlookError, match="deletes such items for good"):
            o.delete_task(tid)
    assert [c[1] for c in r.calls if c[0] == "write/task"] == [["delete", "1"]]


def test_delete_task_names_the_folder_and_unknown_ids(make):
    o, _ = make(writer())
    with pytest.raises(OutlookError, match="me@school.example"):
        o.delete_task(3)
    with pytest.raises(OutlookError, match="On My Computer"):
        o.delete_task(6)
    with pytest.raises(OutlookError, match="No task with id 99"):
        o.delete_task(99)


# -- notes ----------------------------------------------------------------------------------

NOTE_FOLDERS = (
    ("F", 30, "Notes", "me@work.example.com", "exchange"),
    ("F", 31, "Notes", "", ""),
)


def note(nid, title, modified, folder=30, preview=""):
    return ("N", nid, title, modified, folder, preview)


NOTES = rows(
    *NOTE_FOLDERS,
    note(1, "Wifi", "2026-11-02-1000", preview="password is on the fridge"),
    note(2, "Ideas", "2026-11-05-1000", preview="a course on notes"),
    note(3, "Draft", "2026-11-04-1000", folder=31, preview="x"),
    note(4, "Undated", "", preview="no date"),
)


def note_raw(nid="7", title="Wifi", modified="2026-11-02-1000", cats=(), body="the body"):
    return FS.join([nid, title, modified, "30", GS.join(cats)]) + RS + body


def test_list_notes_newest_first_with_fields(make):
    o, _ = make({"read/notes": NOTES})
    res = o.list_notes()
    assert [n["id"] for n in res["notes"]] == ["2", "3", "1", "4"]
    ideas = res["notes"][0]
    assert ideas["title"] == "Ideas" and ideas["preview"] == "a course on notes" and ideas["folder"] == "Notes"
    assert ideas["account"] == "me@work.example.com" and ideas["modified"].startswith("2026-11-05T")
    assert res["notes"][1]["account"] is None and res["notes"][3]["modified"] is None


def test_list_notes_query_looks_at_title_and_preview(make):
    o, _ = make({"read/notes": NOTES})
    assert [n["id"] for n in o.list_notes(query="FRIDGE")["notes"]] == ["1"]
    assert [n["id"] for n in o.list_notes(query="ideas")["notes"]] == ["2"]
    assert o.list_notes(query="zzz")["count"] == 0


def test_list_notes_limit(make):
    o, _ = make({"read/notes": NOTES})
    res = o.list_notes(limit=1)
    assert res["returned"] == 1 and res["has_more"] is True and res["count"] == 4
    with pytest.raises(OutlookError):
        o.list_notes(limit=0)


def test_read_note(make):
    o, r = make({"read/note": note_raw(cats=("Home", "Red"), body="line1\nline2")})
    n = o.read_note("7")
    assert r.last("read/note") == ["7"]
    assert n == {"id": "7", "title": "Wifi", "body": "line1\nline2", "modified": n["modified"],
                 "categories": ["Home", "Red"]}
    assert n["modified"].startswith("2026-11-02T")
    with pytest.raises(OutlookError, match="note_id"):
        o.read_note("seven")


def test_create_note_args(make):
    o, r = make({"read/notes": NOTES, "write/note": "7", "read/note": note_raw()})
    o.create_note(" Wifi ", "line1\nline2")
    assert r.last("write/note") == ["create", "", "title", "Wifi", "html", to_html("line1\nline2")]
    o.create_note("Empty")
    assert r.last("write/note")[-2:] == ["html", to_html("")]
    o.create_note("Local", account="work")
    assert r.last("write/note")[:2] == ["create", "30"]
    with pytest.raises(OutlookError, match="title"):
        o.create_note("  ")
    with pytest.raises(OutlookError, match="No notes folder of an account matching 'nobody'"):
        o.create_note("x", account="nobody")


def test_update_note(make):
    o, r = make({"write/note": "7", "read/note": note_raw()})
    o.update_note(7, title=" New ", body="text")
    assert r.last("write/note") == ["update", "7", "title", "New", "html", to_html("text")]
    o.update_note(7, body="")
    assert r.last("write/note") == ["update", "7", "html", to_html("")]
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_note(7)
    with pytest.raises(OutlookError, match="title"):
        o.update_note(7, title=" ")
    assert len([c for c in r.calls if c[0] == "write/note"]) == 2


def test_update_note_unknown_id_writes_nothing(make):
    o, r = make({"write/note": "7", "read/note": OutlookError("Outlook could not find that item")})
    with pytest.raises(OutlookError, match="could not find"):
        o.update_note(7, title="x")
    assert not any(c[0] == "write/note" for c in r.calls)


def test_delete_note_only_in_exchange_folders(make):
    o, r = make({"read/notes": NOTES, "write/note": "1"})
    res = o.delete_note(1)
    assert r.last("write/note") == ["delete", "1"] and res["deleted"] == "1" and res["title"] == "Wifi"
    with pytest.raises(OutlookError, match="Note 3 is in the folder 'Notes' of On My Computer"):
        o.delete_note(3)
    with pytest.raises(OutlookError, match="No note with id 99"):
        o.delete_note(99)
    assert [c[1] for c in r.calls if c[0] == "write/note"] == [["delete", "1"]]


def test_update_task_checks_new_due_against_current_start(make):
    """A due date before the task's existing start is refused even when start isn't passed."""
    task = FS.join(["T", "5", "Report", "", "2026-11-20-0", "", "«constant ****FlNC»", "«constant ****PrNr»",
                    "145", "", ""]) + RS + ""
    folders = FS.join(["F", "145", "Tasks", "me@work.example.com", "exchange"])
    o, r = make({"read/tasks": folders, "read/task": task, "write/task": "5"})
    with pytest.raises(OutlookError, match="start must not be after due"):
        o.update_task("5", due="2026-11-01")
    assert not [c for c in r.calls if c[0] == "write/task"]
