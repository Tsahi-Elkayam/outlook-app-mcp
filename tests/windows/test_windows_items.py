"""Windows backend, 0.3 items: categories, tasks, notes, contacts, contact groups, the directory."""

from __future__ import annotations

from datetime import datetime

import pytest
from test_windows import wid

from outlook_app_mcp.common import OutlookError
from outlook_app_mcp.windows.com import NO_DATE

# -- categories -------------------------------------------------------------------------------


def test_categories_list_and_create(w):
    o, app, _ = w
    assert o.list_categories() == {"categories": []}
    assert o.create_category("  Blue ") == {"name": "Blue"}
    app.Session.Categories._items[0].Color = 8
    assert o.list_categories() == {"categories": [{"name": "Blue", "color": "blue"}]}
    assert o.create_category("blue") == {"name": "Blue"}  # exists: nothing added, no error from Outlook
    assert len(app.Session.Categories._items) == 1
    for bad in ("", "a,b", "x;y"):
        with pytest.raises(OutlookError):
            o.create_category(bad)


def test_categorize_modes_and_creation(w):
    o, app, r = w
    ids = [wid(r.a, r.m1), wid(r.a, r.m3)]
    assert o.categorize("message", ids, ["Work", "Urgent"]) == {"changed": 2, "categories": ["Work", "Urgent"]}
    assert r.m1.Categories == "Work, Urgent" and r.m1.saved == 1
    assert {c.Name for c in app.Session.Categories} == {"Work", "Urgent"}
    o.categorize("message", ids, ["work", "Extra"])  # case-insensitive: Work is kept as spelled
    assert r.m1.Categories == "Work, Urgent, Extra"
    o.categorize("message", ids[:1], ["urgent"], mode="remove")
    assert r.m1.Categories == "Work, Extra" and r.m3.Categories == "Work, Urgent, Extra"
    o.categorize("message", ids[:1], ["Only"], mode="set")
    assert r.m1.Categories == "Only"
    o.categorize("message", ids[:1], [], mode="set")
    assert r.m1.Categories == ""
    with pytest.raises(OutlookError, match="at least one"):
        o.categorize("message", ids, [], mode="add")
    with pytest.raises(OutlookError, match="mode"):
        o.categorize("message", ids, ["x"], mode="toggle")
    with pytest.raises(OutlookError, match="item_type"):
        o.categorize("folder", ids, ["x"])


def test_categorize_every_kind_of_item(w):
    o, _, r = w
    o.categorize("event", [wid(r.a, r.one_off)], ["E"])
    o.categorize("task", [wid(r.a, r.rent)], ["T"])
    o.categorize("contact", [wid(r.a, r.rina)], ["C"])
    o.categorize("note", [wid(r.a, r.idea)], ["N"])
    assert (r.one_off.Categories, r.rent.Categories, r.rina.Categories, r.idea.Categories) == ("E", "T", "C", "N")
    assert o.read_event(wid(r.a, r.one_off))["categories"] == ["E"]


def test_categorize_checks_every_id_before_changing_anything(w):
    o, app, r = w
    with pytest.raises(OutlookError, match="not a task"):
        o.categorize("task", [wid(r.a, r.rent), wid(r.a, r.m1)], ["X"])
    assert r.rent.Categories == "" and not app.Session.Categories._items
    with pytest.raises(OutlookError, match="not a note"):
        o.categorize("note", [wid(r.a, r.rent)], ["X"])
    with pytest.raises(OutlookError, match="not a calendar event"):
        o.categorize("event", [wid(r.a, r.m1)], ["X"])
    with pytest.raises(OutlookError, match="not a mail message"):
        o.categorize("message", [wid(r.a, r.rent)], ["X"])
    with pytest.raises(OutlookError, match="not a contact"):
        o.categorize("contact", [wid(r.a, r.m1)], ["X"])


# -- tasks -------------------------------------------------------------------------------------


def test_list_tasks(w):
    o, _, r = w
    res = o.list_tasks()
    assert [t["subject"] for t in res["tasks"]] == ["Pay rent", "No date"]  # due first, undated last, no request
    rent, undated = res["tasks"]
    assert rent == {"id": wid(r.a, r.rent), "subject": "Pay rent", "due": "2026-11-10", "start": None,
                    "completed": False, "completed_on": None, "priority": "normal", "categories": [],
                    "folder": "Tasks", "account": "me@work.example.com"}
    assert undated["priority"] == "high" and undated["categories"] == ["Home"]
    done = o.list_tasks(include_completed=True)["tasks"]
    assert [t["subject"] for t in done] == ["Old chore", "Pay rent", "No date"]
    assert done[0]["completed"] is True and done[0]["completed_on"]
    assert [t["subject"] for t in o.list_tasks(query="RENT")["tasks"]] == ["Pay rent"]
    assert [t["subject"] for t in o.list_tasks(query="home")["tasks"]] == ["No date"]  # category
    assert [t["subject"] for t in o.list_tasks(due_before="2026-11-11")["tasks"]] == ["Pay rent"]
    assert o.list_tasks(due_before="2026-11-10")["count"] == 0  # strictly before
    page = o.list_tasks(limit=1)
    assert (page["count"], page["returned"], page["has_more"]) == (2, 1, True)
    assert "Task request" not in [t["subject"] for t in o.list_tasks(include_completed=True)["tasks"]]


def test_read_task(w):
    o, _, r = w
    r.rent.Body = "monthly"
    r.rent.ReminderSet = True
    r.rent.set_local("ReminderTime", datetime(2026, 11, 9, 8, 0))
    t = o.read_task(wid(r.a, r.rent))
    assert t["body"] == "monthly" and t["reminder"].startswith("2026-11-09T08:00")
    assert o.read_task(wid(r.a, r.undated))["reminder"] is None
    for not_task in (r.m1, r.request, r.idea):
        with pytest.raises(OutlookError, match="not a task"):
            o.read_task(wid(r.a, not_task))


def test_create_task(w):
    o, _, r = w
    t = o.create_task("Write report", body="Q4", due="2026-11-20", start="2026-11-15", reminder="2026-11-19T09:00",
                      priority="high", categories=["Work"])
    item = r.tasks._items[-1]
    assert (t["subject"], t["due"], t["start"], t["priority"], t["categories"]) == (
        "Write report", "2026-11-20", "2026-11-15", "high", ["Work"])
    assert t["body"] == "Q4" and t["reminder"].startswith("2026-11-19T09:00") and t["id"] == wid(r.a, item)
    assert item._raw["DueDate"] == datetime(2026, 11, 20) and item.Importance == 2 and item.saved == 1
    # a time of day is dropped: tasks hold dates
    o.create_task("Timed", due="2026-11-21T17:30")
    assert r.tasks._items[-1]._raw["DueDate"] == datetime(2026, 11, 21)
    with pytest.raises(OutlookError, match="start must not be after due"):
        o.create_task("Backwards", due="2026-11-01", start="2026-11-02")
    with pytest.raises(OutlookError, match="subject"):
        o.create_task("  ")
    with pytest.raises(OutlookError, match="importance"):
        o.create_task("x", priority="soon")
    with pytest.raises(OutlookError, match="has no Tasks folder"):
        o.create_task("x", account="gmail")


def test_update_task(w):
    o, _, r = w
    tid = wid(r.a, r.rent)
    t = o.update_task(tid, subject="Pay rent now", due="2026-11-12", priority="low", complete=True)
    assert t["subject"] == "Pay rent now" and t["due"] == "2026-11-12" and t["priority"] == "low"
    assert t["completed"] is True and t["completed_on"] and r.rent.saved == 1
    assert o.update_task(tid, complete=False)["completed"] is False
    assert o.update_task(tid, due="")["due"] is None and r.rent._raw["DueDate"].year == 4501
    r.rent.ReminderSet = True
    assert o.update_task(tid, remove_reminder=True)["reminder"] is None
    assert o.update_task(tid, reminder="2026-11-09T07:00")["reminder"].startswith("2026-11-09T07:00")
    assert o.update_task(tid, reminder="")["reminder"] is None
    assert o.update_task(tid, categories=["A", "B"])["categories"] == ["A", "B"]
    assert o.update_task(tid, start="2026-11-01", due="2026-11-02")["start"] == "2026-11-01"
    with pytest.raises(OutlookError, match="start must not be after due"):
        o.update_task(tid, start="2026-11-09")
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_task(tid)
    with pytest.raises(OutlookError, match="not both"):
        o.update_task(tid, reminder="2026-11-09T07:00", remove_reminder=True)


def test_delete_task_moves_it(w):
    o, _, r = w
    res = o.delete_task(wid(r.a, r.rent))
    assert res["deleted"] == wid(r.a, r.rent) and "permanently" in res["note"] and r.rent.Parent is r.deleted
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_task(wid(r.a, r.rent))
    with pytest.raises(OutlookError, match="Deleted Items"):
        o.update_task(wid(r.a, r.rent), subject="x")
    with pytest.raises(OutlookError, match="not a task"):
        o.delete_task(wid(r.a, r.m1))


# -- notes ---------------------------------------------------------------------------------------


def test_list_and_read_notes(w):
    o, _, r = w
    res = o.list_notes()
    assert [n["title"] for n in res["notes"]] == ["Idea", "Shopping"]  # newest first
    shopping = res["notes"][1]
    assert shopping["preview"] == "milk eggs" and shopping["folder"] == "Notes"
    assert shopping["modified"].startswith("2026-11-01T10:00") and shopping["id"] == wid(r.a, r.shopping)
    assert [n["title"] for n in o.list_notes(query="EGGS")["notes"]] == ["Shopping"]
    assert o.list_notes(limit=1)["has_more"] is True
    note = o.read_note(wid(r.a, r.shopping))
    assert note == {"id": wid(r.a, r.shopping), "title": "Shopping", "body": "milk\neggs",
                    "modified": note["modified"], "categories": []}
    with pytest.raises(OutlookError, match="not a note"):
        o.read_note(wid(r.a, r.rent))


def test_create_and_update_note(w):
    o, _, r = w
    n = o.create_note("Ideas", "one\ntwo")
    item = r.notes._items[-1]
    assert item.Body == "Ideas\r\none\r\ntwo" and n["title"] == "Ideas" and n["body"] == "one\ntwo"
    assert n["modified"] is not None
    n = o.update_note(n["id"], title="Plans")
    assert item.Subject == "Plans" and n["body"] == "one\ntwo"
    n = o.update_note(n["id"], body="three")
    assert item.Body == "Plans\r\nthree" and n["title"] == "Plans"
    n = o.update_note(n["id"], body="")
    assert item.Body == "Plans" and n["body"] == ""
    assert o.create_note("Just a title")["body"] == ""
    with pytest.raises(OutlookError, match="single line"):
        o.create_note("a\nb")
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_note(n["id"])
    with pytest.raises(OutlookError, match="has no Notes folder"):
        o.create_note("x", account="gmail")


def test_delete_note_moves_it(w):
    o, _, r = w
    res = o.delete_note(wid(r.a, r.idea))
    assert res["title"] == "Idea" and "permanently" in res["note"] and r.idea.Parent is r.deleted
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_note(wid(r.a, r.idea))
    with pytest.raises(OutlookError, match="Deleted Items"):
        o.update_note(wid(r.a, r.idea), title="x")


# -- contacts ------------------------------------------------------------------------------------


def test_read_contact_has_every_field(w):
    o, _, r = w
    c = o.read_contact(wid(r.a, r.rina))
    assert c["display_name"] == "Rina K" and c["emails"] == ["rina@x.co"] and c["mobile"] == "050-123-4567"
    for key in ("middle_name", "nickname", "department", "office", "home_street", "home_city", "home_state",
                "home_zip", "home_country", "business_street", "business_city", "business_state",
                "business_zip", "business_country", "web_page", "note"):
        assert c[key] == ""
    assert c["birthday"] is None and c["categories"] == []
    with pytest.raises(OutlookError, match="not a contact"):
        o.read_contact(wid(r.a, r.m1))
    with pytest.raises(OutlookError, match="not a contact"):
        o.read_contact(wid(r.a, r.team))  # a group has its own tools


def test_create_and_update_contact_with_the_new_fields(w):
    o, app, r = w
    c = o.create_contact(["a@x.co"], first_name="Dana", middle_name="M", nickname="D", department="R&D",
                         office="Tel Aviv", home_street="1 Main", home_city="Haifa", home_state="IL",
                         home_zip="31000", home_country="Israel", business_street="2 Biz", business_city="TLV",
                         business_state="IL", business_zip="61000", business_country="Israel",
                         web_page="https://x.example", birthday="1990-05-17", categories=["Friends"],
                         note="met at conf")
    item = r.contacts._items[-1]
    assert (item.MiddleName, item.NickName, item.Department, item.OfficeLocation) == ("M", "D", "R&D", "Tel Aviv")
    assert (item.HomeAddressPostalCode, item.BusinessAddressCountry, item.WebPage) == (
        "31000", "Israel", "https://x.example")
    assert item._raw["Birthday"] == datetime(1990, 5, 17) and item.Categories == "Friends"
    assert c["birthday"] == "1990-05-17" and c["categories"] == ["Friends"] and c["home_city"] == "Haifa"
    assert c["note"] == "met at conf" and {x.Name for x in app.Session.Categories} == {"Friends"}
    c = o.update_contact(c["id"], birthday="", nickname="", categories=[], office="Jerusalem")
    assert c["birthday"] is None and item._raw["Birthday"].year == 4501
    assert c["nickname"] == "" and c["categories"] == [] and c["office"] == "Jerusalem"
    assert o.update_contact(c["id"], birthday="2000-01-02")["birthday"] == "2000-01-02"
    assert o.update_contact(c["id"], categories=["A", "B"])["categories"] == ["A", "B"]
    with pytest.raises(OutlookError, match="Unknown contact field"):
        o.update_contact(c["id"], shoe_size="9")
    with pytest.raises(OutlookError, match="ISO date"):
        o.update_contact(c["id"], birthday="someday")
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_contact(c["id"])


def test_search_contacts_shows_categories(w):
    o, _, r = w
    r.rina.Categories = "VIP, Friends"
    assert o.search_contacts("rina")["contacts"][0]["categories"] == ["VIP", "Friends"]


def test_delete_contact_moves_it(w):
    o, _, r = w
    res = o.delete_contact(wid(r.a, r.avi))
    assert res["name"] == "אבי" and "permanently" in res["note"] and r.avi.Parent is r.deleted
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_contact(wid(r.a, r.gone_contact))
    with pytest.raises(OutlookError, match="not a contact"):
        o.delete_contact(wid(r.a, r.m1))


# -- contact groups ------------------------------------------------------------------------------


def test_contact_group_lifecycle(w):
    o, app, r = w
    g = o.create_contact_group("Friends", ["a@x.co", "b@x.co", "a@x.co"])
    item = r.contacts._items[-1]
    assert g["name"] == "Friends" and g["id"] == wid(r.a, item) and item.saved == 1
    assert [m["address"] for m in g["members"]] == ["a@x.co", "b@x.co"]
    assert [x["name"] for x in o.list_contact_groups()["groups"]] == ["Team rina", "Friends"]
    assert [x["name"] for x in o.list_contact_groups("fri")["groups"]] == ["Friends"]
    g = o.update_contact_group(g["id"], name="Pals", add_members=["c@x.co", "a@x.co"], remove_members=["b@x.co"])
    assert g["name"] == "Pals" and [m["address"] for m in g["members"]] == ["a@x.co", "c@x.co"]
    assert ("removemember", "Pals", "b@x.co") in app.log
    with pytest.raises(OutlookError, match="Not members"):
        o.update_contact_group(g["id"], remove_members=["zed@x.co"])
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_contact_group(g["id"])
    with pytest.raises(OutlookError, match="not an email"):
        o.update_contact_group(g["id"], add_members=["bad"])
    res = o.delete_contact_group(g["id"])
    assert res["name"] == "Pals" and "permanently" in res["note"] and item.Parent is r.deleted
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_contact_group(g["id"])


def test_contact_group_checks(w):
    o, _, r = w
    with pytest.raises(OutlookError, match="name"):
        o.create_contact_group(" ", [])
    with pytest.raises(OutlookError, match="not an email"):
        o.create_contact_group("x", ["bad"])
    before = len(r.contacts._items)
    assert o.create_contact_group("Empty", [])["members"] == []
    assert len(r.contacts._items) == before + 1
    with pytest.raises(OutlookError, match="not a contact group"):
        o.delete_contact_group(wid(r.a, r.rina))
    with pytest.raises(OutlookError, match="not a contact group"):
        o.update_contact_group(wid(r.a, r.rina), name="x")


# -- organization directory ----------------------------------------------------------------------


def test_lookup_directory(w, directory):
    o, _, _ = w
    res = o.lookup_directory("dan")
    assert [p["name"] for p in res["people"]] == ["Dana Levi", "Danny Cohen"] and "note" not in res
    assert res["people"][0] == {"name": "Dana Levi", "address": "dana@corp.example", "title": "CTO",
                                "department": "R&D", "office": "Tel Aviv", "phone": "03-5551234"}
    assert [p["name"] for p in o.lookup_directory("Dana Levi")["people"]] == ["Dana Levi"]  # found once
    assert o.lookup_directory("staff")["people"] == []  # lists are not people
    assert len(o.lookup_directory("dan", limit=1)["people"]) == 1
    with pytest.raises(OutlookError, match="must not be empty"):
        o.lookup_directory(" ")


def test_lookup_directory_says_when_it_stopped_early(w, directory, monkeypatch):
    from outlook_app_mcp.windows import misc

    o, _, _ = w
    monkeypatch.setattr(misc, "GAL_SCAN_CAP", 1)
    res = o.lookup_directory("danny")
    assert res["people"] == [] and "first 1 entries" in res["note"]


def test_lookup_directory_needs_exchange(w):
    o, _, _ = w
    with pytest.raises(OutlookError, match="Exchange / Microsoft 365"):
        o.lookup_directory("dana")


def test_no_date_constant_reads_as_none():
    from outlook_app_mcp.windows.com import from_com

    assert from_com(NO_DATE) is None


def test_categories_are_joined_with_the_users_list_separator(w, monkeypatch):
    from outlook_app_mcp.windows import com

    o, _, r = w
    monkeypatch.setattr(com, "list_separator", lambda: ";")
    o.categorize("message", [wid(r.a, r.m1)], ["A", "B"])
    assert r.m1.Categories == "A; B"
    assert com.categories(r.m1) == ["A", "B"]  # both separators are understood when reading
    assert com.list_separator.__name__  # the real one is asked of Windows only; elsewhere it falls back


def test_list_separator_falls_back_off_windows():
    from outlook_app_mcp.windows.com import list_separator

    assert list_separator() in {",", ";"}
