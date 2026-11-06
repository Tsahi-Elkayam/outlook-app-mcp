"""Google accounts: Outlook's dictionary has no account object for them, so their
folders arrive keyed by root folder and the address is inferred or configured."""

import base64
import json

import pytest
from conftest import FOLDERS, message_raw, rows

from outlook_app_mcp.mac.mail import infer_google_address
from outlook_app_mcp.mac.osa import FS, GS, RS, US, OutlookError, mailbox_key

GMAIL = [
    (159, "google:151", "", "Inbox", 5, 900, "inbox", "google"),
    (162, "google:151", "", "Sent Mail", 0, 40, "sent", "google"),
    (158, "google:151", "", "Drafts", 0, 1, "drafts", "google"),
    (157, "google:151", "", "Trash", 0, 3, "deleted", "google"),
    (154, "google:151", "", "All Mail", 0, 950, "all", "google"),
]
SCHOOL = [(211, "google:204", "", "Inbox", 1, 40, "inbox", "google"), (209, "google:204", "", "Trash", 0, 0, "deleted", "google")]


def folders_reply(sent_151=("me@gmail.example",) * 6, inbox_204=("list@school.example",) * 8):
    folders = FOLDERS + RS + rows(*GMAIL, *SCHOOL)
    google = RS.join([
        FS.join(["151", US.join(sent_151), ""]),
        FS.join(["204", "", US.join(inbox_204)]),
    ])
    return folders + GS + google


def test_inference_prefers_sent_mail_and_marks_inbox_as_guess():
    assert infer_google_address(["a@x"] * 4 + ["b@x"], []) == ("a@x", "sent mail")
    addr, source = infer_google_address([], ["list@x"] * 3 + ["me@x"] * 2)
    assert addr == "list@x" and "guess" in source
    assert infer_google_address(["a@x"] * 3, ["b@x"] * 3) is None  # too few to tell
    assert infer_google_address(["a@x", "b@x", "c@x", "d@x", "e@x"], []) is None  # no majority


def test_google_folders_get_an_address_and_kind(make):
    o, r = make({"read/folders": folders_reply()})
    by_id = {f.id: f for f in o.folders()}
    assert by_id["159"].account == "me@gmail.example" and by_id["159"].is_google
    assert by_id["159"].as_dict()["account_source"] == "sent mail"
    assert by_id["211"].account == "list@school.example" and "guess" in by_id["211"].account_source
    assert by_id["101"].kind == "exchange" and "account_source" not in by_id["101"].as_dict()
    assert r.last("read/folders") == []


def test_overrides_file_wins_and_known_roots_are_not_sampled_again(make, no_real_accounts_file):
    no_real_accounts_file.write_text(json.dumps({"204": "student@st.school.example"}))
    o, r = make({"read/folders": folders_reply()})
    school = o.folder(211)
    assert school.account == "student@st.school.example" and school.account_source == "config"
    o.folders(refresh=True)
    assert r.last("read/folders") == ["151", "204"]
    assert [f.id for f in o.folders("st.school")] == ["211", "209"]


def test_unresolvable_google_account_says_how_to_fix(make):
    o, _ = make({"read/folders": folders_reply(sent_151=(), inbox_204=())})
    f = o.folder(159)
    assert f.account == "google:151" and "set it in" in f.account_source


def test_bad_overrides_file(make, no_real_accounts_file):
    no_real_accounts_file.write_text("[1, 2]")
    o, _ = make({"read/folders": folders_reply()})
    with pytest.raises(OutlookError, match="JSON object"):
        o.folders()


def test_search_skips_all_mail_unless_asked(make):
    o, r = make({"read/folders": folders_reply(), "read/window": ""})
    o.search("x")
    assert "154" not in r.last("read/window")
    o.search("x", folder_id=154)
    assert r.last("read/window")[1:] == ["154"]


def test_cannot_write_as_a_google_account(make):
    o, r = make({"read/folders": folders_reply(), "write/compose": "1", "write/folder": "1"})
    with pytest.raises(OutlookError, match="Google account"):
        o.create_draft(["a@b.co"], "s", "b", account="gmail")
    with pytest.raises(OutlookError, match="Google account"):
        o.create_folder("x", account="gmail")
    assert not any(c[0] in ("write/compose", "write/folder") for c in r.calls)


def test_google_delete_moves_to_its_trash(make):
    work_trash = RS + rows((106, "me@work.example.com", "Work", "Deleted Items", 0, 0, "deleted"))
    folders, google = folders_reply().split(GS, 1)
    o, r = make({
        "read/folders": folders + work_trash + GS + google,
        "read/locate": rows((5, 159), (6, 101)),
        "write/move": "900",
        "write/delete": "1",
    })
    assert o.delete_messages([5, 6])["deleted"] == 2
    assert r.last("write/move") == ["157", "5"]
    assert r.last("write/delete")[1:] == ["6"]
    assert set(r.last("write/delete")[0].split(GS)) == {"106", "157", "209"}


def test_google_message_in_trash_is_not_deleted_again(make):
    o, r = make({"read/folders": folders_reply(), "read/locate": rows((5, 157))})
    with pytest.raises(OutlookError, match="permanently"):
        o.delete_messages([5])
    assert not any(c[0] in ("write/move", "write/delete") for c in r.calls)


def test_google_mail_cannot_be_replied_to_or_forwarded(make):
    o, r = make({"read/folders": folders_reply(), "read/locate": rows((5, 159)), "write/reply": "77"})
    for call in (
        lambda: o.reply_to_message(5, "thanks", send=True),
        lambda: o.reply_to_message(5, "thanks"),
        lambda: o.forward_message(5, ["x@y.co"]),
    ):
        with pytest.raises(OutlookError, match="can't"):
            call()
    assert not any(c[0] == "write/reply" for c in r.calls)


def test_google_draft_is_not_sent(make):
    o, r = make({"read/folders": folders_reply(), "read/message": message_raw(folder_id=158)})
    with pytest.raises(OutlookError, match="Google account"):
        o.send_draft(55)
    assert not any(c[0] == "write/send_draft" for c in r.calls)


def test_read_message_fills_google_account(make):
    raw = message_raw(folder_id=159).replace("me@work.example.com", "", 1)
    o, _ = make({"read/folders": folders_reply(), "read/message": raw})
    assert o.read_message(55)["account"] == "me@gmail.example"


# -- calendars of Google accounts ------------------------------------------------

def xid(mailbox: bytes, rest: bytes = b"\x00\x2e\x00") -> str:
    """An Exchange id: format, type, length, mailbox bytes, then the item part."""
    raw = bytes([1, 3, 36]) + mailbox + rest
    return base64.urlsafe_b64encode(raw).decode().rstrip("=")


# Real compressed mailbox GUIDs from Outlook (Gmail, SCHOOL): 36 characters each.
GMAIL_BOX = b"\x00abcd00\x02-1111-2222\x00-00\x02-00\x0a"
SCHOOL_BOX = b"\x00abcd00\x02-3333\x00-4444-00\x02-00\x0a"


def test_mailbox_key():
    assert mailbox_key(xid(GMAIL_BOX)) == "abcd0000-1111-2222-0000-000000000000"
    assert mailbox_key(xid(SCHOOL_BOX)) == "abcd0000-3333-4444-0000-000000000000"
    # the item part (folder 0x2e vs calendar 0x46) is not part of the key
    assert mailbox_key(xid(GMAIL_BOX)) == mailbox_key(xid(GMAIL_BOX, b"\x00\x46\x07\x01"))
    assert mailbox_key(xid(GMAIL_BOX)) != mailbox_key(xid(SCHOOL_BOX))
    assert mailbox_key("") is None and mailbox_key("!!!") is None and mailbox_key("AA") is None


def test_google_calendars_get_their_account_and_roots_are_hidden(make):
    folders = FOLDERS + RS + rows(*GMAIL, *SCHOOL)
    google = RS.join([
        FS.join(["151", US.join(("me@gmail.example",) * 6), "", xid(GMAIL_BOX)]),
        FS.join(["204", "", US.join(("student@st.school.example",) * 6), xid(SCHOOL_BOX)]),
    ])
    cals = rows(
        (102, "me@work.example.com", "", "false", ""),
        (105, "me@work.example.com", "לוח שנה", "true", "x"),
        (152, "", "", "false", ""),
        (168, "", "me@gmail.example", "false", xid(GMAIL_BOX, b"\x00\x46")),
        (214, "", "Calendar", "false", xid(SCHOOL_BOX, b"\x00\x46")),
        (13, "", "Calendar", "false", ""),
    )
    o, _ = make({"read/folders": folders + GS + google, "read/calendars": cals})
    got = {c["id"]: c["account"] for c in o.list_calendars()}
    assert got == {"105": "me@work.example.com", "168": "me@gmail.example", "214": "student@st.school.example", "13": ""}


def test_known_roots_keep_their_address_without_samples(make, no_real_accounts_file):
    no_real_accounts_file.write_text(json.dumps({"151": "me@gmail.example"}))
    google = FS.join(["151", "", "", xid(GMAIL_BOX)])
    o, _ = make({"read/folders": FOLDERS + RS + rows(*GMAIL) + GS + google})
    assert o.folder(159).account == "me@gmail.example"
    assert o.mailbox_account(xid(GMAIL_BOX, b"\x00\x46")) == "me@gmail.example"
