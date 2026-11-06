from pathlib import Path

import pytest
from conftest import FOLDERS, a, message_raw, rows

from outlook_app_mcp.mac.osa import FS, GS, RS, FolderChanged, OutlookError

FOLDERS_WITH_DELETED = FOLDERS + RS + rows(
    (106, "me@work.example.com", "Work", "Deleted Items", 0, 3, "deleted"),
    (107, "me@work.example.com", "Work", "Deleted Items/Old", 0, 1, ""),
)

# -- folders -------------------------------------------------------------------


def test_folders_parse_roles_and_cache(make):
    o, r = make({})
    fs = o.folders()
    assert [f.id for f in fs] == ["101", "102", "103", "104", "105"]
    assert fs[0].role == "inbox" and fs[1].as_dict()["role"] is None
    o.folders()
    assert [c[0] for c in r.calls] == ["read/folders"]


def test_no_accounts_means_new_outlook(make):
    o, _ = make({"read/folders": ""})
    with pytest.raises(OutlookError, match="Legacy Outlook"):
        o.folders()


def test_unknown_account_lists_accounts(make):
    o, _ = make({})
    with pytest.raises(OutlookError, match="me@work.example.com"):
        o.folders("nobody")


# -- reading -------------------------------------------------------------------

IDX = rows(
    (1, 101, "2026-11-4-28800", "true", "«constant ****FlNF»", "older", a("Rina", "rina@x")),
    (2, 101, "2026-11-4-43200", "false", "«constant ****FlNC»", "שלום", a("אבי", "avi-c@x")),
    (3, 101, "", "true", "", "no date", a("", "")),
)


STAMPS = rows((1, "2026-11-4-28800"), (2, "2026-11-4-43200"), (3, ""))


def summaries_for(args):
    by_id = {r.split(FS)[0]: r for r in IDX.split(RS)}
    return RS.join(by_id[i] for i in args)


def test_list_defaults_to_inbox_newest_first(make):
    o, r = make({"read/recent": "UNSORTED", "read/stamps": STAMPS, "read/summaries": summaries_for})
    res = o.list_messages(limit=2)
    assert r.last("read/recent") == ["101", "2", ""]
    assert r.last("read/stamps") == ["101", "", "0"]
    assert r.last("read/summaries") == ["2", "1"]  # only the newest two are read in full
    assert res["matched"] == 3 and res["returned"] == 2 and res["has_more"] is True
    assert [m["id"] for m in res["messages"]] == ["2", "1"]
    m = res["messages"][0]
    assert m["subject"] == "שלום" and m["is_read"] is False and m["flag"] == "flagged"
    assert m["from"] == {"name": "אבי", "address": "avi-c@x"}
    assert m["received"].startswith("2026-11-04T12:00")
    assert m["folder"] == "Inbox" and m["account"] == "me@work.example.com"


def test_list_reads_the_top_block_when_order_is_proven(make):
    o, r = make({"read/recent": "OK" + RS + IDX})
    res = o.list_messages(limit=5, since_days=1)
    assert r.last("read/recent") == ["101", "5", "86400"]
    assert not any(c[0] in ("read/index", "read/stamps") for c in r.calls)
    assert res["matched"] == 3 and res["has_more"] is False and res["messages"][0]["id"] == "2"


def test_full_block_may_have_more(make):
    o, _ = make({"read/recent": "OK" + RS + IDX})
    res = o.list_messages(limit=3)  # inbox holds 10
    assert res["matched"] is None and res["has_more"] is True and res["returned"] == 3


def test_recent_retries_then_falls_back(make):
    o, r = make({"read/recent": [FolderChanged("x"), FolderChanged("x")], "read/stamps": STAMPS, "read/summaries": summaries_for})
    assert o.list_messages()["matched"] == 3
    assert [c[0] for c in r.calls].count("read/recent") == 2


def test_unread_goes_straight_to_the_scan(make):
    o, r = make({"read/stamps": "", "read/summaries": ""})
    res = o.list_messages(folder_id=102, unread_only=True, since_days=2)
    assert r.last("read/stamps") == ["102", "172800", "1"]
    assert not any(c[0] in ("read/recent", "read/summaries") for c in r.calls)
    assert res["matched"] == 0 and res["has_more"] is False


def test_list_validation(make):
    o, _ = make({"read/stamps": "", "read/recent": "OK"})
    with pytest.raises(OutlookError, match="No mail folder with id 999"):
        o.list_messages(folder_id=999)
    with pytest.raises(OutlookError):
        o.list_messages(limit=0)
    with pytest.raises(OutlookError):
        o.list_messages(since_days=0)
    assert o.list_messages(limit=10_000)["returned"] == 0


def test_scan_retried_once_then_reported(make):
    o, _ = make({"read/stamps": [FolderChanged("x"), STAMPS], "read/summaries": summaries_for})
    assert o.list_messages(unread_only=True)["matched"] == 3
    o, _ = make({"read/stamps": [FolderChanged("x"), FolderChanged("x")]})
    with pytest.raises(OutlookError, match="kept changing"):
        o.list_messages(unread_only=True)


def test_search_index_retried_once_then_reported(make):
    ok = rows((1, 101, "2026-11-4-1", "true", "", "s", a("n", "x@y")))
    o, _ = make({"read/window": [FolderChanged("x"), ok]})
    assert o.search("s", folder_id=101)["matched"] == 1
    o, _ = make({"read/window": [FolderChanged("x"), FolderChanged("x")]})
    with pytest.raises(OutlookError, match="kept changing"):
        o.search("s", folder_id=101)


def test_search_subject_and_sender_skips_empty_folders(make):
    o, r = make({"read/window": IDX, "read/index": IDX})
    assert [m["id"] for m in o.search("OLDER")["messages"]] == ["1"]
    assert r.last("read/window") == [str(90 * 86400), "101", "102", "103", "105"]
    o.search("older", since_days=None)  # all time: the plain scan
    assert r.last("read/index") == ["", "0", "101", "102", "103", "105"]
    assert [m["id"] for m in o.search("avi-c")["messages"]] == ["2"]
    with pytest.raises(OutlookError):
        o.search("  ")


def test_search_body_passes_query_as_data(make):
    evil = 'x" & (do shell script "rm -rf ~") & "'
    o, r = make({"read/index": IDX, "read/body_search": "3"})
    res = o.search(evil, folder_id=101, since_days=None, include_body=True)
    assert r.last("read/body_search") == [evil, "", "101"]
    assert [m["id"] for m in res["messages"]] == ["3"]


def test_read_message(make):
    raw = message_raw(atts=[("a.pdf", "1024", "application/pdf"), ("b.png", "", "")], body="line1\nline2 " + RS + " x")
    o, r = make({"read/message": raw})
    m = o.read_message(55, max_chars=8)
    assert r.last("read/message") == ["55", "0", "0"]
    assert m["subject"] == "נושא" and m["from"] == {"name": "Rina", "address": "rina@x"}
    assert m["received"].startswith("2026-11-04T10:00")
    assert m["to"] == [{"name": "Lee", "address": "lee@x"}] and m["cc"] == []
    assert m["attachments"][0] == {"index": 1, "name": "a.pdf", "size": 1024, "content_type": "application/pdf"}
    assert m["attachments"][1]["size"] is None and m["attachments"][1]["content_type"] is None
    assert m["flag"] is None and m["categories"] == ["Red", "Blue"]
    assert m["body"] == "line1\nli" and m["body_truncated"] is True


def test_read_message_bad_reply(make):
    o, _ = make({"read/message": "1\x1f2"})
    with pytest.raises(OutlookError, match="Unexpected reply"):
        o.read_message(1)


# -- changing ------------------------------------------------------------------


def test_mark_validates_and_invalidates_cache(make):
    o, r = make({"write/mark": "2"})
    o.folders()
    assert o.mark_messages([5, 6], "read") == {"action": "read", "changed": 2}
    assert r.last("write/mark") == ["read", "", "5", "6"]
    o.folders()
    assert [c[0] for c in r.calls].count("read/folders") == 2
    with pytest.raises(OutlookError, match="action"):
        o.mark_messages([5], "delete")
    with pytest.raises(OutlookError, match="empty"):
        o.mark_messages([], "read")


def test_move_maps_new_ids(make):
    o, r = make({"write/move": RS.join(["5", ""])})
    res = o.move_messages([5, 6], 102)
    assert r.last("write/move") == ["102", "5", "6"]
    assert res["new_ids"] == {"5": "5", "6": None} and res["to"]["path"] == "Inbox/Projects"
    assert res["moved"] == 2 and "new id" in res["note"]
    with pytest.raises(OutlookError):
        o.move_messages([5], 999)


def test_delete_passes_protected_folders(make):
    folders = FOLDERS_WITH_DELETED
    o, r = make({"read/folders": folders, "read/locate": rows((7, 101)), "write/delete": "1"})
    assert o.delete_messages([7])["deleted"] == 1
    assert r.last("write/delete") == ["106", "7"]


def test_delete_refuses_accounts_without_deleted_items(make):
    """Some IMAP and POP accounts have no deleted-items folder Outlook can name; whether
    a delete there is permanent can't be told, so nothing is deleted."""
    o, r = make({"read/locate": rows((7, 101)), "write/delete": "1"})
    with pytest.raises(OutlookError, match="no deleted folder"):
        o.delete_messages([7])
    assert not any(c[0] == "write/delete" for c in r.calls)


def test_delete_rereads_folders_first(make):
    o, r = make({"read/folders": FOLDERS_WITH_DELETED, "read/locate": rows((7, 101)), "write/delete": "1"})
    o.folders()
    o.delete_messages([7])
    assert [c[0] for c in r.calls].count("read/folders") == 2


def test_delete_refuses_messages_already_deleted(make):
    o, r = make({"read/folders": FOLDERS_WITH_DELETED, "read/locate": rows((7, 101), (8, 107))})
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_messages([7, 8])
    assert not any(c[0] == "write/delete" for c in r.calls)
    o, _ = make({"read/folders": FOLDERS_WITH_DELETED, "read/locate": rows((7, 106))})
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_messages([7])


def test_delete_unknown_message(make):
    o, _ = make({"read/locate": ""})
    with pytest.raises(OutlookError, match="No message with id 7"):
        o.delete_messages([7])


def test_compose_draft_builds_html_and_recipients(make, tmp_path):
    f = tmp_path / "doc.pdf"
    f.write_bytes(b"x")
    o, r = make({"write/compose": "77"})
    res = o.create_draft(["s@x.co", "t@x.co"], "נושא", "שלום\n\nתודה", cc=["c@x.co"], attachments=[str(f)], account="work.example")
    args = r.last("write/compose")
    assert args[0] == "draft" and args[1] == "me@work.example.com" and args[2] == "נושא"
    assert 'dir="rtl"' in args[3] and "שלום" in args[3]
    assert args[4] == GS.join(["s@x.co", "t@x.co"]) and args[5] == "c@x.co" and args[6] == ""
    assert args[7] == str(f)
    assert res == {"id": "77", "status": "draft", "account": "me@work.example.com"}


def test_compose_checks(make, tmp_path):
    o, _ = make({"write/compose": "1"})
    with pytest.raises(OutlookError, match="recipient"):
        o.send_message([], "s", "b")
    with pytest.raises(OutlookError, match="not an email"):
        o.send_message(["bad"], "s", "b")
    with pytest.raises(OutlookError, match="absolute"):
        o.create_draft(["a@b.co"], attachments=["rel.txt"])
    with pytest.raises(OutlookError, match="not a file"):
        o.create_draft(["a@b.co"], attachments=[str(tmp_path / "missing")])
    assert o.create_draft()["status"] == "draft"  # empty draft is fine
    o.create_draft(["a@b.co"], body="<p>x</p>", body_is_html=True)


def test_send_only_real_drafts_with_recipients(make):
    o, _ = make({"read/message": message_raw(folder_id=105), "write/send_draft": "55"})
    assert o.send_draft(55)["status"] == "sent"
    o, _ = make({"read/message": message_raw(folder_id=101)})
    with pytest.raises(OutlookError, match="not in a Drafts"):
        o.send_draft(55)
    o, _ = make({"read/message": message_raw(folder_id=105, to=())})
    with pytest.raises(OutlookError, match="no recipients"):
        o.send_draft(55)


def test_reply_and_forward_args(make):
    o, r = make({"write/reply": "88", "read/locate": rows((5, 101))})
    res = o.reply_to_message(5, "תודה", reply_all=True)
    args = r.last("write/reply")
    assert args[:2] == ["5", "replyall"] and 'dir="rtl"' in args[2] and args[3] == "תודה"
    # the kept draft goes to the Drafts folder of the original's account (105)
    assert args[4:] == ["", "0", "105", "", ""] and res["status"] == "draft"
    o.forward_message(5, ["x@y.co"], send=True)
    assert r.last("write/reply")[1] == "forward" and r.last("write/reply")[4:] == ["x@y.co", "1", "105", "", ""]
    with pytest.raises(OutlookError, match="at least one"):
        o.forward_message(5, [], send=True)


def test_folder_create_and_rename(make):
    o, r = make({"write/folder": "200"})
    assert o.create_folder("Projects 2026", parent_id=101)["id"] == "200"
    assert r.last("write/folder") == ["create", "101", "", "Projects 2026"]
    o.create_folder("Top")
    assert r.last("write/folder") == ["create", "", "me@work.example.com", "Top"]
    o.rename_folder(102, "Proj")
    assert r.last("write/folder") == ["rename", "102", "", "Proj"]
    with pytest.raises(OutlookError, match="inbox folder"):
        o.rename_folder(101, "x")
    with pytest.raises(OutlookError):
        o.create_folder("a/b", parent_id=101)


def test_save_attachment_never_overwrites(make, tmp_path):
    (tmp_path / "a.pdf").write_bytes(b"old")

    def save(args):
        from pathlib import Path

        Path(args[2]).write_bytes(b"new!")
        return args[2]

    raw = message_raw(atts=[("a.pdf", "4", "application/pdf"), ("x/../evil:name", "1", "")])
    o, r = make({"read/message": raw, "write/save_attachment": save})
    res = o.save_attachment(55, "a.pdf", str(tmp_path))
    assert res["path"] == str((tmp_path / "a (1).pdf").resolve()) and res["size"] == 4
    assert (tmp_path / "a.pdf").read_bytes() == b"old"
    res = o.save_attachment(55, "2", str(tmp_path))
    assert Path(res["path"]).parent == tmp_path.resolve() and r.last("write/save_attachment")[1] == "2"
    with pytest.raises(OutlookError, match="Attachments: 1: a.pdf"):
        o.save_attachment(55, "nope.doc", str(tmp_path))
    with pytest.raises(OutlookError, match="not an existing folder"):
        o.save_attachment(55, "a.pdf", str(tmp_path / "missing"))


def test_safe_filename_drops_bidi_marks():
    from outlook_app_mcp.common import safe_filename as _safe_filename

    assert _safe_filename("\u200e\u2068ממצאים\u2069.docx") == "ממצאים.docx"
    assert _safe_filename("a/b:c.txt") == "a_b_c.txt" and _safe_filename("\u200e") == "attachment"
