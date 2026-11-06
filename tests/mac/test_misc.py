"""Accounts, categories, automatic replies, free/busy, directory, contact groups and contact fields."""

import pytest
from conftest import FOLDERS, rows
from test_calendar_contacts import contact_raw

from outlook_app_mcp.common import to_html
from outlook_app_mcp.mac.contacts import FIELDS
from outlook_app_mcp.mac.misc import parse_free_busy
from outlook_app_mcp.mac.osa import FS, GS, RS, US, OutlookError

ACCOUNTS = rows(
    ("me@work.example.com", "Work", "exchange", "1"),
    ("old@mail.example", "", "imap", "0"),
)
GMAIL_FOLDERS = FOLDERS + RS + rows(
    (159, "google:151", "", "Inbox", 5, 900, "inbox", "google"),
    (162, "google:151", "", "Sent Mail", 0, 40, "sent", "google"),
) + GS + FS.join(["151", US.join(["me@gmail.example"] * 6), ""])


# -- accounts --------------------------------------------------------------------


def test_list_accounts_adds_google_accounts_from_folders(make):
    o, _ = make({"read/accounts": ACCOUNTS, "read/folders": GMAIL_FOLDERS})
    accounts = o.list_accounts()["accounts"]
    assert accounts == [
        {"address": "me@work.example.com", "name": "Work", "kind": "exchange", "is_default": True},
        {"address": "old@mail.example", "name": "old@mail.example", "kind": "imap", "is_default": False},
        {"address": "me@gmail.example", "name": "me@gmail.example", "kind": "google", "is_default": False},
    ]


def test_list_accounts_unresolved_google_account_is_left_out(make):
    unresolved = FOLDERS + RS + rows((159, "google:151", "", "Inbox", 0, 1, "inbox", "google")) + GS + FS.join(["151", "", ""])
    o, _ = make({"read/accounts": ACCOUNTS, "read/folders": unresolved})
    assert [a["address"] for a in o.list_accounts()["accounts"]] == ["me@work.example.com", "old@mail.example"]


def test_list_accounts_first_is_default_when_outlook_names_none(make):
    none_default = rows(("a@x.co", "A", "exchange", "0"), ("b@x.co", "B", "exchange", "0"))
    o, _ = make({"read/accounts": none_default})
    assert [a["is_default"] for a in o.list_accounts()["accounts"]] == [True, False]


def test_list_accounts_needs_legacy_outlook(make):
    o, _ = make({"read/folders": "", "read/accounts": ACCOUNTS})
    with pytest.raises(OutlookError, match="Legacy Outlook"):
        o.list_accounts()


def test_send_and_receive_syncs_only_scriptable_accounts(make):
    o, r = make({"read/accounts": ACCOUNTS, "read/folders": GMAIL_FOLDERS, "write/sync": ""})
    res = o.send_and_receive()
    assert r.last("write/sync") == ["me@work.example.com", "old@mail.example"]
    assert res == {"status": "started", "accounts": ["me@work.example.com", "old@mail.example"]}


def test_send_and_receive_with_only_google_accounts(make):
    only_google = rows((159, "google:151", "", "Inbox", 0, 1, "inbox", "google")) + GS + FS.join(["151", US.join(["g@x.co"] * 6), ""])
    o, r = make({"read/accounts": "", "read/folders": only_google, "write/sync": ""})
    with pytest.raises(OutlookError, match="can be synced"):
        o.send_and_receive()
    assert not any(c[0] == "write/sync" for c in r.calls)


# -- categories ------------------------------------------------------------------

CATEGORIES = rows(("Red", "65535,0,0"), ("Teal", "0,32896,32768"), ("None", ""), ("Odd", "x,y"))


def test_list_categories_converts_16_bit_colors(make):
    o, _ = make({"read/categories": CATEGORIES})
    assert o.list_categories()["categories"] == [
        {"name": "Red", "color": "#ff0000"},
        {"name": "Teal", "color": "#008080"},
        {"name": "None", "color": None},
        {"name": "Odd", "color": None},
    ]


def test_create_category_reuses_an_existing_name_in_any_case(make):
    o, r = make({"read/categories": CATEGORIES, "write/category": "x"})
    assert o.create_category(" red ") == {"name": "Red"}
    assert not any(c[0] == "write/category" for c in r.calls)
    assert o.create_category("Blue") == {"name": "Blue"}
    assert r.last("write/category") == ["create", "Blue"]


@pytest.mark.parametrize("bad", ["", "  ", "a,b", "a;b"])
def test_create_category_refuses_bad_names(make, bad):
    o, r = make({"read/categories": CATEGORIES})
    with pytest.raises(OutlookError):
        o.create_category(bad)
    assert not any(c[0] == "write/category" for c in r.calls)


def _item_cats(mapping):
    def reply(args):
        assert args[0] in ("message", "task", "event", "contact", "note")
        return rows(*[(i, GS.join(mapping[i])) for i in args[1:] if i in mapping])

    return reply


def test_categorize_add_keeps_order_and_creates_missing_categories(make):
    o, r = make({"read/item_categories": _item_cats({"5": ["Red"], "6": []}), "read/categories": CATEGORIES,
                 "write/category": "2"})
    res = o.categorize("message", [5, 6], ["red", "Brand new"])
    assert r.last("read/item_categories") == ["message", "5", "6"]
    created = [c[1] for c in r.calls if c[0] == "write/category" and c[1][0] == "create"]
    assert created == [["create", "Brand new"]]
    # "red" resolves to the existing spelling "Red", which message 5 already has
    assert r.last("write/category") == ["assign", "message", "5", GS.join(["Red", "Brand new"]), "6",
                                        GS.join(["Red", "Brand new"])]
    assert res == {"changed": 2, "categories": ["Red", "Brand new"], "mode": "add"}


def test_categorize_remove_does_not_create_categories(make):
    o, r = make({"read/item_categories": _item_cats({"5": ["Red", "Teal"]}), "read/categories": CATEGORIES,
                 "write/category": "1"})
    o.categorize("task", ["5"], ["RED", "Never existed"], mode="remove")
    assert r.last("write/category") == ["assign", "task", "5", "Teal"]
    assert not any(c[0] == "write/category" and c[1][0] == "create" for c in r.calls)


def test_categorize_set_replaces_and_may_clear(make):
    o, r = make({"read/item_categories": _item_cats({"5": ["Red", "Teal"]}), "read/categories": CATEGORIES,
                 "write/category": "1"})
    o.categorize("event", [5], ["Teal"], mode="set")
    assert r.last("write/category") == ["assign", "event", "5", "Teal"]
    o.categorize("event", [5], [], mode="set")
    assert r.last("write/category") == ["assign", "event", "5", ""]
    with pytest.raises(OutlookError, match="at least one"):
        o.categorize("event", [5], [], mode="add")


def test_categorize_refuses_missing_ids_and_bad_input(make):
    o, r = make({"read/item_categories": _item_cats({"5": []}), "read/categories": CATEGORIES, "write/category": "1"})
    with pytest.raises(OutlookError, match="No message with id 9"):
        o.categorize("message", [5, 9], ["Red"])
    with pytest.raises(OutlookError, match="item_type"):
        o.categorize("folder", [5], ["Red"])
    with pytest.raises(OutlookError, match="mode"):
        o.categorize("message", [5], ["Red"], mode="toggle")
    with pytest.raises(OutlookError, match="item_ids"):
        o.categorize("message", [], ["Red"])
    assert not any(c[0] == "write/category" for c in r.calls)


# -- automatic replies -------------------------------------------------------------

AUTO_ACCOUNTS = rows(("me@work.example.com", "Work", "exchange", "1"), ("old@mail.example", "Old", "imap", "0"),
                     ("two@work.example.com", "Two", "exchange", "0"))


def auto_raw(enabled="false", scheduled="false", start="", end="", ext="true", known="false",
             internal="<p>in</p>", external="<p>out</p>"):
    return FS.join([enabled, scheduled, start, end, ext, known, internal, external])


def test_get_auto_reply_parses_fields_and_audience(make):
    raw = auto_raw("true", "true", "2026-11-10-32400", "2026-11-12-61200", "true", "true")
    o, r = make({"read/accounts": AUTO_ACCOUNTS, "read/autoreply": raw})
    res = o.get_auto_reply()
    assert r.last("read/autoreply") == ["me@work.example.com"]
    assert res["enabled"] and res["scheduled"] and res["external_audience"] == "known"
    assert res["start"].startswith("2026-11-10T09:00") and res["end"].startswith("2026-11-12T17:00")
    assert res["internal_message"] == "<p>in</p>" and res["external_message"] == "<p>out</p>"
    assert o.get_auto_reply()["account"] == "me@work.example.com"
    r.replies["read/autoreply"] = auto_raw(ext="false")
    res = o.get_auto_reply()
    assert res["external_audience"] == "none" and res["start"] is None and res["end"] is None
    r.replies["read/autoreply"] = auto_raw(ext="true", known="false")
    assert o.get_auto_reply()["external_audience"] == "all"


def test_auto_reply_is_exchange_only(make):
    o, r = make({"read/accounts": AUTO_ACCOUNTS, "read/folders": GMAIL_FOLDERS, "read/autoreply": auto_raw(),
                 "write/autoreply": "ok"})
    for call in (lambda: o.get_auto_reply("gmail.example"), lambda: o.get_auto_reply("old@mail"),
                 lambda: o.set_auto_reply("gmail.example", enabled=True)):
        with pytest.raises(OutlookError, match="Exchange"):
            call()
    with pytest.raises(OutlookError, match="No account matches"):
        o.get_auto_reply("nobody")
    with pytest.raises(OutlookError, match="several accounts"):
        o.get_auto_reply("work.example.com")
    assert not any(c[0] in ("read/autoreply", "write/autoreply") for c in r.calls)


def test_auto_reply_default_account_needs_an_exchange_account(make):
    o, _ = make({"read/accounts": rows(("old@mail.example", "Old", "imap", "1"))})
    with pytest.raises(OutlookError, match="has none"):
        o.get_auto_reply()
    o, _ = make({"read/accounts": rows(("i@x.co", "I", "imap", "1"), ("a@x.co", "A", "exchange", "0"),
                                        ("b@x.co", "B", "exchange", "0"))})
    with pytest.raises(OutlookError, match="pass account"):
        o.get_auto_reply()
    o, _ = make({"read/accounts": rows(("i@x.co", "I", "imap", "1"), ("only@x.co", "A", "exchange", "0")),
                 "read/autoreply": auto_raw()})
    assert o.get_auto_reply()["account"] == "only@x.co"


def test_set_auto_reply_argument_order_and_schedule(make):
    o, r = make({"read/accounts": AUTO_ACCOUNTS, "read/autoreply": auto_raw(ext="true", known="false"),
                 "write/autoreply": "ok"})
    res = o.set_auto_reply(enabled=True, internal_message="Away", start="2026-11-10T09:00", end="2026-11-12T17:00")
    assert r.last("write/autoreply") == [
        "me@work.example.com", "1", "1", "2026-11-10-32400", "2026-11-12-61200",
        to_html("Away"), "<p>out</p>",  # external message untouched; audience stays "all"
        "1", "0",
    ]
    assert res["account"] == "me@work.example.com"


def test_set_auto_reply_audience_and_unscheduled(make):
    o, r = make({"read/accounts": AUTO_ACCOUNTS, "read/autoreply": auto_raw(), "write/autoreply": "ok"})
    o.set_auto_reply(enabled=False, external_message="Gone", external_audience="known")
    assert r.last("write/autoreply") == ["me@work.example.com", "0", "0", "", "", "<p>in</p>", to_html("Gone"), "1", "1"]
    o.set_auto_reply(enabled=True, external_audience="none")
    assert r.last("write/autoreply")[7:] == ["0", "0"]
    o.set_auto_reply(enabled=True, external_audience="all")
    assert r.last("write/autoreply")[7:] == ["1", "0"]


def test_set_auto_reply_refusals(make):
    o, r = make({"read/accounts": AUTO_ACCOUNTS, "read/autoreply": auto_raw(), "write/autoreply": "ok"})
    with pytest.raises(OutlookError, match="both start and end"):
        o.set_auto_reply(enabled=True, start="2026-11-10T09:00")
    with pytest.raises(OutlookError, match="end must be after"):
        o.set_auto_reply(enabled=True, start="2026-11-10T09:00", end="2026-11-10T09:00")
    with pytest.raises(OutlookError, match="external_audience"):
        o.set_auto_reply(enabled=True, external_audience="everyone")
    assert not any(c[0] == "write/autoreply" for c in r.calls)


# -- free/busy ---------------------------------------------------------------------

DAY = "2026-11-09 08:00:00 +0200"


def person(address, *statuses, start=DAY):
    out = [address, start]
    for s in statuses:
        out += ["", "", s]
    return out


def test_parse_free_busy_maps_statuses_to_digits():
    raw = GS.join(person("a@x.co", "free", "tentative", "busy", "out of office", "working elsewhere"))
    assert parse_free_busy(raw, ["a@x.co"], 5) == [{"address": "a@x.co", "slots": "01234", "error": None}]


def test_parse_free_busy_unknown_status_is_a_question_mark_and_pads_or_trims():
    raw = GS.join(person("a@x.co", " BUSY ", "weird"))
    assert parse_free_busy(raw, ["a@x.co"], 4)[0]["slots"] == "2???"
    assert parse_free_busy(raw, ["a@x.co"], 1)[0]["slots"] == "2"


def test_parse_free_busy_several_people_and_missing_ones():
    raw = GS.join(person("A@x.co", "busy", "free") + person("b@x.co", "tentative", "tentative"))
    res = parse_free_busy(raw, ["a@x.co", "b@x.co", "c@x.co"], 2)
    assert [(p["address"], p["slots"], p["error"] is None) for p in res] == [
        ("a@x.co", "20", True), ("b@x.co", "11", True), ("c@x.co", "??", False)]
    assert "no free/busy" in res[2]["error"]


def test_parse_free_busy_a_subject_that_looks_like_an_address_is_not_a_person():
    items = ["a@x.co", DAY, "b@x.co", "", "busy", "", "", "free"]  # subject "b@x.co" is followed by "", not a date
    assert parse_free_busy(GS.join(items), ["a@x.co", "b@x.co"], 2)[0]["slots"] == "20"


def test_parse_free_busy_empty_answer():
    res = parse_free_busy("", ["a@x.co"], 3)
    assert res[0]["slots"] == "???" and res[0]["error"]


def test_get_free_busy_request_and_result(make):
    raw = GS.join(person("a@x.co", "free", "busy", "busy", "out of office"))
    o, r = make({"read/accounts": ACCOUNTS, "read/freebusy": raw})
    res = o.get_free_busy(["a@x.co"], "2026-11-09T08:00", "2026-11-09T10:00", 30)
    assert r.last("read/freebusy") == ["me@work.example.com", "2026-11-9-28800", "2026-11-9-36000", "30", "a@x.co"]
    assert res["interval_minutes"] == 30 and res["start"].startswith("2026-11-09T08:00")
    assert res["people"] == [{"address": "a@x.co", "slots": "0223", "error": None}]


def test_get_free_busy_rounds_partial_last_interval_up(make):
    o, _ = make({"read/accounts": ACCOUNTS, "read/freebusy": GS.join(person("a@x.co", "busy", "busy", "busy"))})
    res = o.get_free_busy(["a@x.co"], "2026-11-09T08:00", "2026-11-09T09:10", 30)
    assert res["people"][0]["slots"] == "222"


def test_get_free_busy_refusals(make):
    o, r = make({"read/accounts": ACCOUNTS, "read/freebusy": ""})
    for args, match in [
        (([], "2026-11-09", "2026-11-10"), "at least one"),
        ((["nope"], "2026-11-09", "2026-11-10"), "not an email"),
        ((["a@x.co"], "2026-11-10", "2026-11-09"), "after start"),
        ((["a@x.co"], "2026-11-09", "2027-01-09"), "6 weeks"),
    ]:
        with pytest.raises(OutlookError, match=match):
            o.get_free_busy(*args)
    with pytest.raises(OutlookError, match="5-1440"):
        o.get_free_busy(["a@x.co"], "2026-11-09", "2026-11-10", 3)
    assert not any(c[0] == "read/freebusy" for c in r.calls)
    o, _ = make({"read/accounts": rows(("old@mail.example", "Old", "imap", "1"))})
    with pytest.raises(OutlookError, match="needs an Exchange"):
        o.get_free_busy(["a@x.co"], "2026-11-09", "2026-11-10")


# -- directory -----------------------------------------------------------------------


def dir_rec(name, address):
    return f"{name}{US}{address}"


def test_lookup_directory_returns_people_without_duplicates(make):
    raw = rows((dir_rec("Noa Levi", "noa@x.co"),), (dir_rec("Noa L", "NOA@x.co"),), (dir_rec("Room 4", ""),),
               (dir_rec("Avi", "avi@x.co"),))
    o, r = make({"read/accounts": ACCOUNTS, "read/directory": raw})
    res = o.lookup_directory(" Noa ")
    assert r.last("read/directory") == ["me@work.example.com", "Noa"]
    assert res["query"] == "Noa" and res["account"] == "me@work.example.com"
    assert [p["address"] for p in res["people"]] == ["noa@x.co", "avi@x.co"]
    assert res["people"][0] == {"name": "Noa Levi", "address": "noa@x.co", "title": "", "department": "",
                                "office": "", "phone": ""}
    assert [p["address"] for p in o.lookup_directory("Noa", limit=1)["people"]] == ["noa@x.co"]


def test_lookup_directory_multi_word_falls_back_to_first_word_and_filters(make):
    def reply(args):
        if args[1] == "Noa Levi":
            return ""
        assert args[1] == "Noa"
        return rows((dir_rec("Noa Cohen", "noa.c@x.co"),), (dir_rec("Noa Levi", "noa.l@x.co"),),
                    (dir_rec("Noa B", "noa@levi-law.co"),))

    o, r = make({"read/accounts": ACCOUNTS, "read/directory": reply})
    res = o.lookup_directory("Noa Levi")
    assert [c[1][1] for c in r.calls if c[0] == "read/directory"] == ["Noa Levi", "Noa"]
    assert [p["address"] for p in res["people"]] == ["noa.l@x.co", "noa@levi-law.co"]


def test_lookup_directory_no_fallback_when_the_full_query_found_people(make):
    o, r = make({"read/accounts": ACCOUNTS, "read/directory": rows((dir_rec("Noa Levi", "n@x.co"),))})
    o.lookup_directory("Noa Levi")
    assert len([c for c in r.calls if c[0] == "read/directory"]) == 1


def test_lookup_directory_refusals(make):
    o, _ = make({"read/accounts": ACCOUNTS, "read/directory": ""})
    with pytest.raises(OutlookError, match="empty"):
        o.lookup_directory("   ")
    with pytest.raises(OutlookError, match="limit"):
        o.lookup_directory("x", limit=0)
    assert o.lookup_directory("nobody")["people"] == []
    o, _ = make({"read/accounts": rows(("old@mail.example", "Old", "imap", "1"))})
    with pytest.raises(OutlookError, match="needs an Exchange"):
        o.lookup_directory("x")


# -- contact groups ------------------------------------------------------------------


def group_rec(gid="3", name="Team", members=(("Avi", "avi@x.co"), ("Noa", "noa@x.co")), kind="exchange",
              account="me@work.example.com"):
    return (gid, name, GS.join(f"{n}{US}{a}" for n, a in members), "9", account, kind)


def test_list_contact_groups_hides_private_fields_and_filters(make):
    o, _ = make({"read/groups": rows(group_rec(), group_rec("4", "Family", (), "", ""))})
    groups = o.list_contact_groups()["groups"]
    assert groups[0] == {"id": "3", "name": "Team", "account": "me@work.example.com",
                         "members": [{"name": "Avi", "address": "avi@x.co"}, {"name": "Noa", "address": "noa@x.co"}]}
    assert groups[1]["members"] == [] and groups[1]["account"] is None
    assert [g["id"] for g in o.list_contact_groups("FAM")["groups"]] == ["4"]


def test_update_contact_group_sends_the_whole_member_list(make):
    o, r = make({"read/groups": rows(group_rec()), "write/group": "3"})
    o.update_contact_group(3, add_members=["new@x.co", "NOA@x.co", "new@x.co"], remove_members=["AVI@x.co"])
    assert r.last("write/group") == ["update", "3", "", GS.join([f"Noa{US}noa@x.co", f"new@x.co{US}new@x.co"])]


def test_update_contact_group_rename_keeps_members(make):
    o, r = make({"read/groups": rows(group_rec()), "write/group": "3"})
    o.update_contact_group(3, name=" Crew ")
    assert r.last("write/group") == ["update", "3", "Crew", "-"]


def test_update_contact_group_refusals(make):
    o, r = make({"read/groups": rows(group_rec()), "write/group": "3"})
    with pytest.raises(OutlookError, match="Not in the group: zed@x.co"):
        o.update_contact_group(3, remove_members=["zed@x.co"])
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_contact_group(3)
    with pytest.raises(OutlookError, match="No contact group with id 99"):
        o.update_contact_group(99, name="x")
    with pytest.raises(OutlookError, match="not an email"):
        o.update_contact_group(3, add_members=["bad"])
    assert not any(c[0] == "write/group" for c in r.calls)


def test_create_contact_group_dedupes_members(make):
    o, r = make({"read/groups": rows(group_rec()), "write/group": "3"})
    res = o.create_contact_group(" Team ", ["a@x.co", "a@x.co", "b@x.co"])
    assert r.last("write/group") == ["create", "", "Team", GS.join([f"a@x.co{US}a@x.co", f"b@x.co{US}b@x.co"])]
    assert res["id"] == "3" and "_kind" not in res
    with pytest.raises(OutlookError, match="name"):
        o.create_contact_group("  ")


def test_delete_contact_group_only_in_exchange(make):
    o, r = make({"read/groups": rows(group_rec(), group_rec("4", "Local", kind="", account="")), "write/group": "3"})
    res = o.delete_contact_group(3)
    assert r.last("write/group") == ["delete", "3"] and res["deleted"] == "3" and res["name"] == "Team"
    with pytest.raises(OutlookError, match="deletes such items for good"):
        o.delete_contact_group(4)
    assert [c[1] for c in r.calls if c[0] == "write/group"] == [["delete", "3"]]


# -- full contact fields ---------------------------------------------------------------


def test_create_contact_maps_every_field_to_its_script_key(make):
    o, r = make({"write/contact": "2", "read/contact": contact_raw("2")})
    fields = {arg: f"v-{arg}" for arg in FIELDS}
    o.create_contact(**fields)
    args = r.last("write/contact")
    assert args[:2] == ["create", ""]
    assert dict(zip(args[2::2], args[3::2])) == {key: f"v-{arg}" for arg, key in FIELDS.items()}
    assert len(FIELDS) == 23


def test_contact_birthday_emails_categories_pairs(make):
    o, r = make({"write/contact": "2", "read/contact": contact_raw("2"), "read/categories": rows(("VIP", "")),
                 "write/category": "New"})
    o.update_contact(2, birthday="1990-05-17", emails=["a@x.co", "b@x.co"], categories=["vip", "New"])
    assert r.last("write/contact") == ["update", "2", "birthday", "1990-5-17-43200", "emails",
                                       GS.join(["a@x.co", "b@x.co"]), "categories", GS.join(["VIP", "New"])]
    assert [c[1] for c in r.calls if c[0] == "write/category"] == [["create", "New"]]


def test_contact_birthday_cannot_be_cleared(make):
    o, r = make({"write/contact": "2", "read/contact": contact_raw("2")})
    with pytest.raises(OutlookError, match="not clear it"):
        o.update_contact(2, birthday="")
    with pytest.raises(OutlookError, match="not clear it"):
        o.create_contact(first_name="Noa", birthday="")
    with pytest.raises(OutlookError, match="not an ISO date"):
        o.update_contact(2, birthday="soon")
    assert not any(c[0] == "write/contact" for c in r.calls)


def test_contact_unknown_field_and_bad_emails_refused(make):
    o, r = make({"write/contact": "2", "read/contact": contact_raw("2")})
    with pytest.raises(OutlookError, match="Unknown contact field"):
        o.update_contact(2, shoe_size="43")
    with pytest.raises(OutlookError, match="not an email"):
        o.update_contact(2, emails=["nope"])
    assert not any(c[0] == "write/contact" for c in r.calls)


def test_create_contact_in_a_chosen_address_book(make):
    books = rows((11, "", ""), (12, "me@work.example.com", "exchange"))
    o, r = make({"read/address_books": books, "write/contact": "2", "read/contact": contact_raw("2")})
    o.create_contact(first_name="Noa", account="WORK.example")
    assert r.last("write/contact")[:2] == ["create", "12"]
    with pytest.raises(OutlookError, match="No address book"):
        o.create_contact(first_name="Noa", account="other")
