"""Windows backend, 0.3 mail features: accounts, paging, search filters, read extras, drafts, signatures,
conversations, folders, categories, out of office."""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

import pytest
from fakecom import Attachment, Mail, Recipient
from test_windows import GMAIL, ME, wid

from outlook_app_mcp.common import OutlookError
from outlook_app_mcp.windows.com import PR_OOF_STATE, PR_TRANSPORT_HEADERS

# Real, absolute on every system, and outside the allowed folders (tests allow only tmp_path).
OUTSIDE_FILE = str(Path(__file__).resolve())
OUTSIDE_DIR = str(Path(__file__).resolve().parent)

# -- accounts and send/receive ---------------------------------------------------------


def test_list_accounts(w):
    o, _, _ = w
    accounts = o.list_accounts()["accounts"]
    assert accounts == [
        {"address": ME, "name": ME, "kind": "exchange", "is_default": True},
        {"address": GMAIL, "name": GMAIL, "kind": "imap", "is_default": False},
    ]


def test_google_imap_accounts_are_called_google(w):
    o, app, _ = w
    app.Session.accounts[1].SmtpAddress = "me@gmail.com"
    assert o.list_accounts()["accounts"][1]["kind"] == "google"


def test_send_and_receive_starts_it(w):
    o, app, _ = w
    assert o.send_and_receive() == {"status": "started"}
    assert app.Session.send_receive == 1


# -- paging and search ------------------------------------------------------------------


def test_list_messages_offset_pages(w):
    o, _, _r = w
    first = o.list_messages(account="work", limit=1)
    second = o.list_messages(account="work", limit=1, offset=1)
    third = o.list_messages(account="work", limit=5, offset=2)
    assert first["messages"][0]["id"] != second["messages"][0]["id"]
    assert second["offset"] == 1 and second["has_more"] is True
    assert third["returned"] == third["matched"] - 2 and third["has_more"] is False
    with pytest.raises(OutlookError, match="offset"):
        o.list_messages(offset=-1)


def test_search_filters_without_a_query(w):
    o, _, r = w
    r.m1.Categories = "Blue, Project"
    r.m1.Attachments._items = [Attachment("x.txt")]
    ids = lambda res: [m["subject"] for m in res["messages"]]
    assert ids(o.search(from_address="RINA@x")) == ["older"]
    assert ids(o.search(from_address="avi-c")) == ["שלום"]  # the SMTP address, not the X.500 one
    assert set(ids(o.search(has_attachments=True))) == {"older", "report"}
    assert "older" not in ids(o.search(has_attachments=False))
    assert ids(o.search(flagged=True)) == ["שלום"]
    assert "שלום" not in ids(o.search(flagged=False))
    assert ids(o.search(category="blue")) == ["older"]
    assert ids(o.search(category="green")) == []
    with pytest.raises(OutlookError, match="query or at least one filter"):
        o.search()


def test_search_to_address_opens_candidates(w):
    o, _, _ = w
    assert [m["subject"] for m in o.search(to_address="lee@x")["messages"]] == ["שלום"]
    assert [m["subject"] for m in o.search(to_address="c@x.co")["messages"]] == ["שלום"]  # Cc counts
    assert o.search(to_address="nobody@x.co")["matched"] == 0
    # combines with the text query
    assert o.search("report", to_address="lee@x")["matched"] == 0


def test_search_to_address_scan_is_capped(w, monkeypatch):
    from outlook_app_mcp.windows import mail

    o, _, _ = w
    monkeypatch.setattr(mail, "MAX_RECIPIENT_SCAN", 1)
    res = o.search(to_address="zzz")
    assert "newest 1 candidates" in res["note"]


def test_search_dates_offset_and_has_more(w):
    o, app, _ = w
    after = (datetime.now() - timedelta(days=10)).date().isoformat()
    res = o.search(after=after, since_days=1)  # after overrides since_days
    subjects = {m["subject"] for m in res["messages"]}
    assert res["since_days"] is None and {"older", "שלום"} <= subjects and "report" not in subjects
    before = (datetime.now() - timedelta(days=1)).date().isoformat()
    subjects = {m["subject"] for m in o.search(before=before, since_days=None)["messages"]}
    assert {"older", "report"} <= subjects and "שלום" not in subjects
    assert any('"urn:schemas:httpmail:datereceived" <' in e[2] for e in app.log if e[0] == "table")
    paged = o.search("x.co", limit=1, offset=1, since_days=None)
    assert paged["offset"] == 1 and paged["returned"] == 1 and paged["has_more"] is True
    with pytest.raises(OutlookError, match="later than after"):
        o.search("x", after="2026-11-05", before="2026-11-01")


def test_search_body_text_respects_filters(w):
    o, _, _ = w
    assert o.search("quarterly", include_body=True, since_days=None, from_address="dana")["matched"] == 1
    assert o.search("quarterly", include_body=True, since_days=None, from_address="nobody")["matched"] == 0


# -- reading ----------------------------------------------------------------------------


def test_read_message_extras(w):
    o, _, r = w
    r.m2.Importance = 2
    r.m2.Categories = "Blue; Project"
    r.m2.ConversationID = "CID-1"
    r.m2.PropertyAccessor._props[PR_TRANSPORT_HEADERS] = "Received: from x\r\nSubject: hi"
    m = o.read_message(wid(r.a, r.m2), 100)
    assert (m["importance"], m["categories"], m["conversation_id"]) == ("high", ["Blue", "Project"], "CID-1")
    assert m["body_format"] == "text" and "headers" not in m and m["meeting_event_id"] is None
    assert o.read_message(wid(r.a, r.m2), 100, include_headers=True)["headers"].startswith("Received:")
    assert o.read_message(wid(r.a, r.m1), 100)["conversation_id"] is None
    html = o.read_message(wid(r.a, r.m2), 10, as_html=True)
    assert html["body_format"] == "html" and html["body"] == r.m2.HTMLBody[:10]
    assert html["body_truncated"] is True


def test_meeting_messages_point_to_their_event(w):
    o, _, r = w
    m = o.read_message(wid(r.a, r.invite), 50)
    assert m["is_meeting"] and m["meeting_event_id"] == wid(r.a, r.meeting)
    assert o.read_message(wid(r.a, r.cancelled), 50)["meeting_event_id"] is None  # no event any more


def test_mark_flag_with_due_date(w):
    o, _, r = w
    res = o.mark_messages([wid(r.a, r.m1)], "flag", due_date="2026-11-20")
    assert res == {"action": "flag", "changed": 1, "due_date": "2026-11-20"}
    assert r.m1.FlagStatus == 2 and r.m1._raw["TaskDueDate"] == datetime(2026, 11, 20)
    assert r.m1._raw["TaskStartDate"] == datetime(2026, 11, 20)
    with pytest.raises(OutlookError, match="only goes with"):
        o.mark_messages([wid(r.a, r.m1)], "read", due_date="2026-11-20")
    with pytest.raises(OutlookError, match="ISO date"):
        o.mark_messages([wid(r.a, r.m1)], "flag", due_date="soon")


# -- composing: importance, signatures --------------------------------------------------


@pytest.fixture
def signatures(tmp_path, monkeypatch):
    folder = tmp_path / "Signatures"
    folder.mkdir()
    (folder / "Work.htm").write_text(
        "<html><head><style>p{color:red}</style></head><body><p>Best,<br>Lee</p></body></html>", encoding="utf-8")
    (folder / "Work.txt").write_bytes("Best,\r\nLee".encode("utf-16"))
    (folder / "Home.htm").write_bytes("<html><body><p>Cheers &amp; thanks</p></body></html>".encode("cp1252"))
    monkeypatch.setenv("OUTLOOK_APP_MCP_SIGNATURES_DIR", str(folder))
    return folder


def test_list_signatures(w, signatures):
    o, _, _ = w
    assert o.list_signatures() == {"signatures": [
        {"name": "Home", "preview": "Cheers & thanks"},
        {"name": "Work", "preview": "Best, Lee"},
    ]}


def test_no_signature_folder_means_no_signatures(w, tmp_path, monkeypatch):
    o, _, _ = w
    monkeypatch.setenv("OUTLOOK_APP_MCP_SIGNATURES_DIR", str(tmp_path / "missing"))
    assert o.list_signatures() == {"signatures": []}
    monkeypatch.delenv("OUTLOOK_APP_MCP_SIGNATURES_DIR")
    monkeypatch.delenv("APPDATA", raising=False)
    with pytest.raises(OutlookError, match="APPDATA"):
        o.list_signatures()


def test_draft_with_importance_and_signature(w, signatures):
    o, _, r = w
    res = o.create_draft(["s@x.co"], "hi", "Hello", importance="high", signature="work", account="work")
    draft = r.drafts._items[-1]
    assert res["id"] == wid(r.a, draft) and draft.Importance == 2
    assert draft.HTMLBody.index("Hello") < draft.HTMLBody.index("Best,") and "<style>" not in draft.HTMLBody
    with pytest.raises(OutlookError, match="Signatures: Home, Work"):
        o.create_draft(["s@x.co"], "hi", "Hello", signature="Nope")
    with pytest.raises(OutlookError, match="importance"):
        o.create_draft(["s@x.co"], "hi", "Hello", importance="urgent")


def test_send_with_signature_and_importance(w, signatures):
    o, app, _r = w
    o.send_message(["s@x.co"], "sig", "Hello", importance="low", signature="Home")
    assert ("send", "sig") in app.log


def test_reply_and_forward_with_cc_attachments_signature(w, signatures, tmp_path):
    o, _, r = w
    f = tmp_path / "plan.pdf"
    f.write_bytes(b"x")
    res = o.reply_to_message(wid(r.a, r.m1), "Thanks", reply_all=True, cc=["cc@x.co"], attachments=[str(f)],
                             signature="Work")
    reply = r.drafts._items[-1]
    assert res["status"] == "draft" and not reply.sent
    assert [(x.Address, x.Type) for x in reply.Recipients] == [("cc@x.co", 2)]
    assert reply.Attachments.Item(1).FileName == "plan.pdf"
    html = reply.HTMLBody
    assert html.index("Thanks") < html.index("Best,") < html.index("quoted")
    o.forward_message(wid(r.a, r.m1), ["to@x.co"], "FYI", cc=["cc@x.co"], signature="Home")
    fwd = r.drafts._items[-1]
    assert [(x.Address, x.Type) for x in fwd.Recipients] == [("to@x.co", 1), ("cc@x.co", 2)]
    # a signature alone, without a body, still goes in
    o.reply_to_message(wid(r.a, r.m1), "", signature="Home")
    assert "Cheers" in r.drafts._items[-1].HTMLBody


def test_reply_checks_files_and_addresses_before_creating_anything(w, tmp_path):
    o, _, r = w
    before = len(r.drafts._items)
    with pytest.raises(OutlookError, match="outside the folders"):
        o.reply_to_message(wid(r.a, r.m1), "x", attachments=[OUTSIDE_FILE])
    with pytest.raises(OutlookError, match="not an email"):
        o.reply_to_message(wid(r.a, r.m1), "x", cc=["bad"])
    assert len(r.drafts._items) == before


def test_attach_files_must_be_in_the_allowed_folders(w):
    o, _, _ = w
    with pytest.raises(OutlookError, match="outside the folders"):
        o.create_draft(["s@x.co"], "x", "y", attachments=[OUTSIDE_FILE])


def test_save_all_attachments(w, tmp_path):
    o, _, r = w
    res = o.save_attachment(wid(r.a, r.m3), "all", str(tmp_path))
    assert [s["name"] for s in res["saved"]] == ["a.pdf", "x/../evil:name"]
    assert all((tmp_path / s["path"].split("\\")[-1].split("/")[-1]).exists() for s in res["saved"])
    with pytest.raises(OutlookError, match="no attachments"):
        o.save_attachment(wid(r.a, r.m1), "all", str(tmp_path))
    with pytest.raises(OutlookError, match="outside the folders"):
        o.save_attachment(wid(r.a, r.m3), "all", OUTSIDE_DIR)


# -- update_draft -------------------------------------------------------------------------


def test_update_draft_replaces_recipients_body_and_attachments(w, tmp_path):
    o, _, r = w
    f = tmp_path / "new.txt"
    f.write_bytes(b"n")
    r.draft.Attachments._items = [Attachment("a.txt"), Attachment("b.txt"), Attachment("c.txt")]
    r.draft.Recipients._items.append(Recipient("cc-old@x.co", type_=2))
    res = o.update_draft(wid(r.a, r.draft), to=["n1@x.co", "n2@x.co"], cc=[], bcc=["b@x.co"], subject="New",
                         body="שלום", add_attachments=[str(f)], remove_attachments=["a.txt", 3], importance="low")
    assert res == {"id": wid(r.a, r.draft), "status": "draft"}
    assert [(x.Address, x.Type) for x in r.draft.Recipients] == [("n1@x.co", 1), ("n2@x.co", 1), ("b@x.co", 3)]
    assert r.draft.Subject == "New" and 'dir="rtl"' in r.draft.HTMLBody and r.draft.Importance == 0
    assert [a.FileName for a in r.draft.Attachments] == ["b.txt", "new.txt"]
    assert r.draft.Recipients.resolved and r.draft.saved == 1 and not r.draft.sent
    # untouched parts stay: only the subject changes here
    o.update_draft(wid(r.a, r.draft), subject="Again")
    assert r.draft.Recipients.Count == 3


def test_update_draft_refuses_what_is_not_an_unsent_draft(w):
    o, _, r = w
    with pytest.raises(OutlookError, match="not an unsent draft"):
        o.update_draft(wid(r.a, r.m1), subject="x")
    r.draft.sent = True
    with pytest.raises(OutlookError, match="not an unsent draft"):
        o.update_draft(wid(r.a, r.draft), subject="x")
    r.draft.sent = False
    with pytest.raises(OutlookError, match="Nothing to change"):
        o.update_draft(wid(r.a, r.draft))
    with pytest.raises(OutlookError, match="not an email"):
        o.update_draft(wid(r.a, r.draft), to=["bad"])
    with pytest.raises(OutlookError, match="No attachment"):
        o.update_draft(wid(r.a, r.draft), remove_attachments=["nope.txt"])
    with pytest.raises(OutlookError, match="outside the folders"):
        o.update_draft(wid(r.a, r.draft), add_attachments=[OUTSIDE_FILE])
    assert r.draft.saved == 0


# -- conversations ------------------------------------------------------------------------


def test_get_conversation_gathers_the_thread_across_folders(w):
    o, _, r = w
    now = datetime.now()
    a = Mail("Plan", ("Rina", "rina@x.co"), now - timedelta(days=3))
    b = Mail("RE: Plan", ("Me", ME), now - timedelta(days=2))
    c = Mail("RE: RE: Plan", ("Rina", "rina@x.co"), now - timedelta(days=1))
    other = Mail("RE: Plan", ("Zed", "zed@x.co"), now)  # same subject, another conversation
    elsewhere = Mail("Plan", ("Bob", "bob@y.co"), now)  # another account
    r.inbox.put(a, other)
    r.sent.put(b)
    r.deleted.put(c)
    r.g_inbox.put(elsewhere)
    for m in (a, b, c, other, elsewhere):
        m.ConversationTopic = "Plan"
        m.ConversationID = "C-OTHER" if m in (other, elsewhere) else "C-1"
    res = o.get_conversation(wid(r.a, b), 50)
    assert res["conversation_id"] == "C-1" and res["count"] == 3
    assert [m["subject"] for m in res["messages"]] == ["Plan", "RE: Plan", "RE: RE: Plan"]  # oldest first
    assert [m["folder"] for m in res["messages"]] == ["Inbox", "Sent Items", "Deleted Items"]
    newest = o.get_conversation(wid(r.a, a), 2)
    assert newest["count"] == 3 and newest["returned"] == 2 and newest["has_more"] is True
    assert [m["subject"] for m in newest["messages"]] == ["RE: Plan", "RE: RE: Plan"]


def test_get_conversation_needs_an_id_and_works_without_a_topic(w):
    o, _, r = w
    with pytest.raises(OutlookError, match="no conversation id"):
        o.get_conversation(wid(r.a, r.m1), 10)
    r.m1.ConversationID = "C-9"
    res = o.get_conversation(wid(r.a, r.m1), 10)
    assert res["count"] == 1 and res["messages"][0]["id"] == wid(r.a, r.m1)


# -- folders: move and delete -------------------------------------------------------------


def test_move_folder_inside_an_account(w):
    o, _, r = w
    res = o.move_folder(wid(r.a, r.projects), wid(r.a, r.sent))
    assert res["path"] == "Sent Items/Projects" and res["id"] == wid(r.a, r.projects)
    assert r.projects.Parent is r.sent and r.projects not in r.inbox._subs
    assert not any(e[0] == "move" for e in w[1].log)  # folders move with MoveTo, not item Move


def test_move_folder_refusals(w):
    o, _, r = w
    sub = r.projects.add("Sub")
    with pytest.raises(OutlookError, match="inbox folder"):
        o.move_folder(wid(r.a, r.inbox), wid(r.a, r.sent))
    with pytest.raises(OutlookError, match="own account"):
        o.move_folder(wid(r.a, r.projects), wid(r.g, r.g_inbox))
    with pytest.raises(OutlookError, match="itself or one of its own"):
        o.move_folder(wid(r.a, r.projects), wid(r.a, sub))
    with pytest.raises(OutlookError, match="delete_folder"):
        o.move_folder(wid(r.a, r.projects), wid(r.a, r.deleted))
    with pytest.raises(OutlookError, match="delete_folder"):
        o.move_folder(wid(r.a, r.old), wid(r.a, r.inbox))
    r.sent.add("Projects")
    with pytest.raises(OutlookError, match="already has a folder named"):
        o.move_folder(wid(r.a, r.projects), wid(r.a, r.sent))
    assert r.projects.Parent is r.inbox


def test_delete_folder_moves_it_to_deleted_items(w):
    o, app, r = w
    res = o.delete_folder(wid(r.a, r.projects))
    assert res["deleted"] == wid(r.a, r.projects) and "permanently" in res["note"]
    assert r.projects.Parent is r.deleted and ("moveto", "Projects", "Deleted Items") in app.log
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_folder(wid(r.a, r.projects))  # already there now
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_folder(wid(r.a, r.old))
    with pytest.raises(OutlookError, match="inbox folder"):
        o.delete_folder(wid(r.a, r.inbox))
    with pytest.raises(OutlookError, match="deleted folder"):
        o.delete_folder(wid(r.a, r.deleted))
    assert r.inbox.Parent is r.a.root


def test_delete_folder_refuses_a_name_clash_and_stores_without_deleted_items(w):
    o, _, r = w
    r.deleted.add("Projects")
    with pytest.raises(OutlookError, match="already has a folder named"):
        o.delete_folder(wid(r.a, r.projects))
    assert r.projects.Parent is r.inbox
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_folder(wid(r.archive, r.old_mail))
    assert r.old_mail.Parent is r.archive.root


# -- out of office ------------------------------------------------------------------------


def test_get_auto_reply_reads_only_the_state(w):
    o, _, r = w
    res = o.get_auto_reply()
    assert res["account"] == ME and res["enabled"] is None and "File > Automatic Replies" in res["note"]
    r.a.PropertyAccessor._props[PR_OOF_STATE] = True
    res = o.get_auto_reply()
    assert res["enabled"] is True
    assert all(res[k] is None for k in ("scheduled", "start", "end", "internal_message", "external_message",
                                        "external_audience"))
    assert o.get_auto_reply("gmail")["account"] == GMAIL
    with pytest.raises(OutlookError, match="No account matches"):
        o.get_auto_reply("nobody")


def test_set_auto_reply_is_a_platform_gap(w):
    o, _, _ = w
    with pytest.raises(OutlookError, match="File > Automatic Replies"):
        o.set_auto_reply(enabled=True, internal_message="away")
