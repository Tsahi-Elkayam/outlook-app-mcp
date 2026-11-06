"""Mail: paging, search filters, conversations, drafts, folders, attachments, signatures, flags."""

from pathlib import Path

import pytest
from conftest import FOLDERS, a, message_raw, rows

from outlook_app_mcp.common import to_html
from outlook_app_mcp.mac.osa import FS, GS, RS, US, OutlookError

ME = "me@work.example.com"
NF, NC = "«constant ****FlNF»", "«constant ****FlNC»"

IDX = rows(
    (1, 101, "2026-11-4-28800", "true", NF, "older", a("Rina", "rina@x")),
    (2, 101, "2026-11-4-43200", "false", NC, "newer", a("Avi", "avi-c@x")),
    (3, 101, "", "true", "", "no date", a("", "")),
)


def ids(res):
    return [m["id"] for m in res["messages"]]


# -- list_messages: offset -----------------------------------------------------------------


def top_block(n):
    recs = [(i, 101, f"2026-11-4-{60000 - i * 100}", "true", NF, f"m{i}", a("S", "s@x")) for i in range(1, n + 1)]
    return "OK" + RS + rows(*recs)


def test_list_messages_offset_reads_offset_plus_limit_off_the_top(make):
    o, r = make({"read/recent": top_block(5)})
    res = o.list_messages(limit=2, offset=3)
    assert r.last("read/recent") == ["101", "5", ""]
    assert ids(res) == ["4", "5"] and res["offset"] == 3 and res["returned"] == 2
    assert res["matched"] is None and res["has_more"] is True


def test_list_messages_offset_past_a_short_block_is_empty_and_complete(make):
    o, _ = make({"read/recent": top_block(2)})
    res = o.list_messages(limit=5, offset=3)
    assert res["messages"] == [] and res["matched"] == 2 and res["has_more"] is False


def test_list_messages_offset_in_the_scan_path(make):
    stamps = rows((1, "2026-11-4-28800"), (2, "2026-11-4-43200"), (3, "2026-11-3-100"))

    def summaries(args):
        by_id = {r.split(FS)[0]: r for r in IDX.split(RS)}
        return RS.join(by_id[i] for i in args)

    o, r = make({"read/stamps": stamps, "read/summaries": summaries})
    res = o.list_messages(unread_only=True, limit=1, offset=1)
    assert r.last("read/summaries") == ["1"]
    assert ids(res) == ["1"] and res["matched"] == 3 and res["has_more"] is True
    res = o.list_messages(unread_only=True, limit=5, offset=2)
    assert r.last("read/summaries") == ["3"] and res["has_more"] is False
    res = o.list_messages(unread_only=True, limit=5, offset=3)
    assert res["messages"] == [] and res["has_more"] is False


def test_list_messages_negative_offset(make):
    o, r = make({"read/recent": top_block(1)})
    with pytest.raises(OutlookError, match="offset"):
        o.list_messages(offset=-1)
    assert not any(c[0] == "read/recent" for c in r.calls)


# -- search filters -------------------------------------------------------------------------------


def searcher(make, **extra):
    return make({"read/window": IDX, "read/index": IDX, **extra})


def test_search_needs_a_query_or_a_filter(make):
    o, _ = searcher(make)
    with pytest.raises(OutlookError, match="query or at least one filter"):
        o.search("")
    with pytest.raises(OutlookError, match="include_body needs a query"):
        o.search("", from_address="x", include_body=True)
    assert o.search("", flagged=False)["matched"] == 2  # False is a filter too


def test_search_from_address(make):
    o, r = searcher(make)
    assert ids(o.search("", from_address="AVI-C")) == ["2"]
    assert r.last("read/window")[0] == str(90 * 86400)
    assert ids(o.search("", from_address="@x")) == ["2", "1"]  # newest first; the nameless sender has no address
    assert o.search("", from_address="zzz")["matched"] == 0


def test_search_flagged(make):
    o, _ = searcher(make)
    assert ids(o.search("", flagged=True)) == ["2"]
    assert ids(o.search("", flagged=False)) == ["1", "3"]


def test_search_after_and_before_use_received_time(make):
    o, r = searcher(make)
    res = o.search("", after="2026-11-04T10:00")
    assert ids(res) == ["2"] and res["since_days"] is None
    assert r.last("read/window")[0].isdigit()
    assert ids(o.search("", before="2026-11-04T10:00")) == ["1"]
    assert ids(o.search("", after="2026-11-04T08:00", before="2026-11-04T12:00")) == ["1"]  # after is inclusive, before exclusive
    assert o.search("", before="2026-11-05")["matched"] == 2  # the undated message never matches a date filter
    with pytest.raises(OutlookError, match="not an ISO date"):
        o.search("", after="yesterday")


def test_search_filters_combine_with_the_query(make):
    o, _ = searcher(make)
    assert ids(o.search("older", flagged=False)) == ["1"]
    assert ids(o.search("older", flagged=True)) == []


FACTS = rows(
    (1, GS.join(["lee@x.co"]), 0, ""),
    (2, GS.join(["bob@x.co", "Lee@Y.co"]), 2, GS.join(["Red", "Blue"])),
)


def test_search_to_address_asks_read_facts_for_recipients(make):
    o, r = searcher(make, **{"read/facts": FACTS})
    res = o.search("", to_address=" LEE@")
    assert r.last("read/facts") == ["t", str(90 * 86400), "101"]
    assert ids(res) == ["2", "1"]
    assert ids(o.search("", to_address="bob@")) == ["2"]
    assert ids(o.search("", from_address="avi", to_address="lee@y")) == ["2"]


def test_search_attachment_and_category_filters(make):
    o, r = searcher(make, **{"read/facts": FACTS})
    assert ids(o.search("", has_attachments=True)) == ["2"] and r.last("read/facts")[0] == "a"
    assert ids(o.search("", has_attachments=False)) == ["1"]  # message 3 has no facts record and drops out
    assert ids(o.search("", category=" red ")) == ["2"] and r.last("read/facts")[0] == "c"
    assert ids(o.search("", category="green")) == []
    o.search("", to_address="x", has_attachments=True, category="red")
    assert r.last("read/facts")[0] == "tac"


def test_search_facts_only_for_folders_with_hits_and_not_when_nothing_matched(make):
    o, r = searcher(make, **{"read/facts": FACTS})
    o.search("", from_address="zzz", to_address="x")
    assert not any(c[0] == "read/facts" for c in r.calls)
    o.search("", flagged=True, category="red", since_days=None)
    assert r.last("read/facts") == ["c", "", "101"]


def test_search_pages_with_offset(make):
    o, _ = searcher(make)
    first = o.search("", from_address="@x", limit=1)
    assert ids(first) == ["2"] and first["matched"] == 2 and first["has_more"] is True
    second = o.search("", from_address="@x", limit=1, offset=1)
    assert ids(second) == ["1"] and second["has_more"] is False and second["offset"] == 1
    with pytest.raises(OutlookError, match="offset"):
        o.search("", from_address="x", offset=-1)


# -- get_conversation --------------------------------------------------------------------------


def with_conversation(conv="321", folder_id=101):
    head, _, body = message_raw(folder_id=folder_id).partition(RS)
    return FS.join(head.split(FS) + ["priority normal", conv, ""]) + RS + body


def test_get_conversation_reads_the_accounts_folders_oldest_first(make):
    o, r = make({"read/message": with_conversation(), "read/conversation": IDX})
    res = o.get_conversation(55)
    assert r.last("read/message") == ["55", "0", "0"]
    assert r.last("read/conversation") == ["321", "101", "102", "103", "105"]  # empty folder 104 skipped
    assert res["conversation_id"] == "321" and res["count"] == 3 and res["returned"] == 3
    assert ids(res) == ["3", "1", "2"]  # undated first, then oldest to newest
    assert res["messages"][1]["folder"] == "Inbox" and res["messages"][1]["account"] == ME


def test_get_conversation_limit_and_validation(make):
    o, _ = make({"read/message": with_conversation(), "read/conversation": IDX})
    res = o.get_conversation(55, limit=2)
    assert res["count"] == 3 and res["returned"] == 2 and len(res["messages"]) == 2
    with pytest.raises(OutlookError, match="limit"):
        o.get_conversation(55, limit=0)


def test_get_conversation_without_an_id(make):
    o, r = make({"read/message": message_raw(), "read/conversation": IDX})
    with pytest.raises(OutlookError, match="no conversation id"):
        o.get_conversation(55)
    assert not any(c[0] == "read/conversation" for c in r.calls)


# -- update_draft ----------------------------------------------------------------------------------

ATTS = [("a.pdf", "10", "application/pdf"), ("b.png", "20", "image/png"), ("c.txt", "30", "text/plain")]


def drafter(make, folder_id=105, **kw):
    return make({"read/message": message_raw(folder_id=folder_id, atts=ATTS, **kw), "write/draft": "55"})


def test_update_draft_pairs_in_order(make, tmp_path):
    f = tmp_path / "new.pdf"
    f.write_bytes(b"x")
    o, r = drafter(make)
    res = o.update_draft(55, to=["a@x.co", "b@x.co"], cc=["c@x.co"], bcc=[], subject="נושא חדש", body="שלום",
                         add_attachments=[str(f)], remove_attachments=["a.pdf", 3], importance="HIGH")
    assert r.last("write/draft") == [
        "55",
        "to", GS.join(["a@x.co", "b@x.co"]),
        "cc", "c@x.co",
        "bcc", "",  # an empty list clears the recipients
        "subject", "נושא חדש",
        "html", to_html("שלום"),
        "add", str(f.resolve()),
        "remove", GS.join(["3", "1"]),  # highest index first
        "importance", "high",
    ]
    assert res == {"id": "55", "status": "draft"}


def test_update_draft_html_body_is_kept(make):
    o, r = drafter(make)
    o.update_draft(55, body="<p>x</p>", body_is_html=True)
    assert r.last("write/draft") == ["55", "html", "<p>x</p>"]


def test_update_draft_remove_attachment_by_name_or_index_once(make):
    o, r = drafter(make)
    o.update_draft(55, remove_attachments=["b.png", "2", 2])
    assert r.last("write/draft") == ["55", "remove", "2"]


def test_update_draft_refusals(make):
    o, r = drafter(make, folder_id=101)
    with pytest.raises(OutlookError, match="not in a Drafts folder"):
        o.update_draft(55, subject="x")
    o, r = drafter(make)
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_draft(55)
    with pytest.raises(OutlookError, match="not an email"):
        o.update_draft(55, to=["bad"])
    with pytest.raises(OutlookError, match="No attachment 'nope.doc'.*1: a.pdf"):
        o.update_draft(55, remove_attachments=["nope.doc"])
    with pytest.raises(OutlookError, match="absolute"):
        o.update_draft(55, add_attachments=["rel.txt"])
    with pytest.raises(OutlookError, match="importance"):
        o.update_draft(55, importance="urgent")
    with pytest.raises(OutlookError, match="draft_id"):
        o.update_draft("x", subject="s")
    assert not any(c[0] == "write/draft" for c in r.calls)


def test_update_draft_drops_the_folder_cache(make):
    o, r = drafter(make)
    o.folders()
    o.update_draft(55, subject="s")
    o.folders()
    assert [c[0] for c in r.calls].count("read/folders") == 2


# -- move_folder / delete_folder ----------------------------------------------------------------

FOLDER_ROWS = (
    (106, ME, "Work", "Deleted Items", 0, 3, "deleted"),
    (107, ME, "Work", "Deleted Items/Old", 0, 1, ""),
    (108, ME, "Work", "Empty/Sub", 0, 0, ""),
    (201, "other@x.co", "Other", "Inbox", 0, 1, "inbox"),
    (202, "other@x.co", "Other", "Stuff", 0, 1, ""),
    (170, "google:151", "", "Label", 0, 1, "", "google"),
    (171, "google:151", "", "Trash", 0, 0, "deleted", "google"),
)
BIG = FOLDERS + RS + rows(*FOLDER_ROWS) + GS + FS.join(["151", US.join(["me@gmail.example"] * 6), ""])


def folders_world(make, **extra):
    return make({"read/folders": BIG, "write/folder": "x", **extra})


def test_move_folder(make):
    o, r = folders_world(make)
    res = o.move_folder(104, 102)
    assert r.last("write/folder") == ["move", "104", "102", ""]
    assert res == {"id": "104", "old_path": "Empty", "parent_id": "102"}


@pytest.mark.parametrize("folder, parent, match", [
    (101, 102, "inbox folder.*can't be moved or deleted"),
    (170, 102, "Google label"),
    (104, 202, "within its own account"),
    (104, 104, "into itself"),
    (104, 108, "into itself"),  # 108 is Empty/Sub, below 104
    (104, 106, "delete_folder"),
    (104, 107, "delete_folder"),
    (104, 999, "No mail folder with id 999"),
    (999, 102, "No mail folder with id 999"),
])
def test_move_folder_refusals(make, folder, parent, match):
    o, r = folders_world(make)
    with pytest.raises(OutlookError, match=match):
        o.move_folder(folder, parent)
    assert not any(c[0] == "write/folder" for c in r.calls)


def test_move_folder_between_google_and_exchange_is_refused(make):
    o, r = folders_world(make)
    with pytest.raises(OutlookError):
        o.move_folder(104, 170)
    assert not any(c[0] == "write/folder" for c in r.calls)


def test_delete_folder_moves_it_into_the_accounts_deleted_items(make):
    o, r = folders_world(make)
    res = o.delete_folder(104)
    assert r.last("write/folder") == ["move", "104", "106", ""]
    assert res["deleted"] == "104" and res["path"] == "Empty" and "nothing was deleted permanently" in res["note"]


def test_delete_folder_refusals(make):
    o, r = folders_world(make)
    for fid, match in [(101, "can't be moved or deleted"), (106, "can't be moved or deleted"),
                       (107, "already in a deleted-items folder"), (170, "Google label"),
                       (202, "no deleted folder"), (999, "No mail folder")]:
        with pytest.raises(OutlookError, match=match):
            o.delete_folder(fid)
    assert not any(c[0] == "write/folder" for c in r.calls)


def test_delete_folder_rereads_the_folders_first(make):
    o, r = folders_world(make)
    o.folders()
    o.delete_folder(104)
    assert [c[0] for c in r.calls].count("read/folders") == 2


# -- save_attachment "all" ---------------------------------------------------------------------------


def writes_files(args):
    Path(args[2]).write_bytes(b"data-" + args[1].encode())
    return args[2]


def test_save_attachment_all(make, tmp_path):
    atts = [("a.pdf", "5", "application/pdf"), ("a.pdf", "5", "application/pdf"), ("b.png", "6", "image/png")]
    o, r = make({"read/message": message_raw(atts=atts), "write/save_attachment": writes_files})
    res = o.save_attachment(55, "ALL", str(tmp_path))
    saved = res["saved"]
    assert [Path(s["path"]).name for s in saved] == ["a.pdf", "a (1).pdf", "b.png"]  # never overwrites
    assert [c[1][:2] for c in r.calls if c[0] == "write/save_attachment"] == [["55", "1"], ["55", "2"], ["55", "3"]]
    assert [s["name"] for s in saved] == ["a.pdf", "a.pdf", "b.png"] and saved[2]["size"] == len(b"data-3")
    assert (tmp_path / "a (1).pdf").read_bytes() == b"data-2"


def test_save_attachment_all_without_attachments(make, tmp_path):
    o, r = make({"read/message": message_raw(), "write/save_attachment": writes_files})
    with pytest.raises(OutlookError, match="has no attachments"):
        o.save_attachment(55, "all", str(tmp_path))
    assert not any(c[0] == "write/save_attachment" for c in r.calls)


def test_save_attachment_all_refuses_a_bad_folder_before_saving(make, tmp_path):
    o, r = make({"read/message": message_raw(atts=ATTS), "write/save_attachment": writes_files})
    with pytest.raises(OutlookError, match="not an existing folder"):
        o.save_attachment(55, "all", str(tmp_path / "missing"))
    assert not any(c[0] == "write/save_attachment" for c in r.calls)


# -- signatures ---------------------------------------------------------------------------------------

SIGS = rows(("Work", "Lee, DevOps"), ("Personal", "T."))


def signatures(args):
    return SIGS if not args else {"Work": "<b>Lee</b>", "Personal": "<i>T.</i>"}[args[0]]


def test_list_signatures(make):
    o, r = make({"read/signatures": signatures})
    assert o.list_signatures() == {"signatures": [{"name": "Work", "preview": "Lee, DevOps"},
                                                  {"name": "Personal", "preview": "T."}]}
    assert r.last("read/signatures") == []
    o, _ = make({"read/signatures": ""})
    assert o.list_signatures() == {"signatures": []}


def test_create_draft_appends_the_signature(make):
    o, r = make({"read/signatures": signatures, "write/compose": "77"})
    o.create_draft(["a@x.co"], "s", "Hello", signature="work")  # matched without regard to case
    assert [c[1] for c in r.calls if c[0] == "read/signatures"] == [[], ["Work"]]
    assert r.last("write/compose")[3] == to_html("Hello") + "<br><b>Lee</b>"
    o.send_message(["a@x.co"], "s", "<p>Hi</p>", body_is_html=True, signature="Personal")
    assert r.last("write/compose")[3] == "<p>Hi</p><br><i>T.</i>"


def test_create_draft_without_a_signature_reads_none(make):
    o, r = make({"read/signatures": signatures, "write/compose": "77"})
    o.create_draft(["a@x.co"], "s", "Hello")
    assert r.last("write/compose")[3] == to_html("Hello")
    assert not any(c[0] == "read/signatures" for c in r.calls)


def test_unknown_signature_is_refused_and_nothing_is_written(make):
    o, r = make({"read/signatures": signatures, "write/compose": "77"})
    with pytest.raises(OutlookError, match=r"No signature named 'Nope'. Signatures: Work, Personal"):
        o.create_draft(["a@x.co"], "s", "Hello", signature="Nope")
    o, r2 = make({"read/signatures": "", "write/compose": "77"})
    with pytest.raises(OutlookError, match="Signatures: none"):
        o.create_draft(["a@x.co"], "s", "Hello", signature="Work")
    assert not any(c[0] == "write/compose" for c in r.calls + r2.calls)


def test_reply_appends_the_signature(make):
    o, r = make({"read/signatures": signatures, "write/reply": "88", "read/locate": rows((5, 101))})
    o.reply_to_message(5, "thanks", signature="Work")
    assert r.last("write/reply")[2] == to_html("thanks") + "<br><b>Lee</b>"
    o.reply_to_message(5, "", signature="Work")  # no text, only the signature
    assert r.last("write/reply")[2] == "<br><b>Lee</b>"
    with pytest.raises(OutlookError, match="No signature"):
        o.forward_message(5, ["a@x.co"], signature="Nope")


# -- mark_messages due_date ---------------------------------------------------------------------------


def test_flag_with_a_due_date(make):
    o, r = make({"write/mark": "2"})
    assert o.mark_messages([5, 6], "flag", due_date="2026-11-20") == {"action": "flag", "changed": 2}
    assert r.last("write/mark") == ["flag", "2026-11-20-0", "5", "6"]
    o.mark_messages([5], "flag", due_date="2026-11-20T09:30")
    assert r.last("write/mark") == ["flag", "2026-11-20-34200", "5"]
    o.mark_messages([5], "flag")
    assert r.last("write/mark") == ["flag", "", "5"]


def test_due_date_only_with_flag(make):
    o, r = make({"write/mark": "1"})
    for action in ("read", "unread", "unflag", "complete"):
        with pytest.raises(OutlookError, match="due_date goes with action 'flag'"):
            o.mark_messages([5], action, due_date="2026-11-20")
    with pytest.raises(OutlookError, match="not an ISO date"):
        o.mark_messages([5], "flag", due_date="next week")
    assert not any(c[0] == "write/mark" for c in r.calls)
    o.mark_messages([5], "read", due_date="")  # an empty date means none
    assert r.last("write/mark") == ["read", "", "5"]
