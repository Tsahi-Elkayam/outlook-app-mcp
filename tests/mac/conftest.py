import pytest

from outlook_app_mcp.mac import MacOutlook
from outlook_app_mcp.mac.osa import FS, GS, RS, US


def rows(*recs):
    return RS.join(FS.join(str(x) for x in r) for r in recs)


def a(name, address):
    return f"{name}{US}{address}"


FOLDERS = rows(
    (101, "me@work.example.com", "Work", "Inbox", 3, 10, "inbox"),
    (102, "me@work.example.com", "Work", "Inbox/Projects", 0, 4, ""),
    (103, "me@work.example.com", "Work", "Sent Items", 0, 7, "sent"),
    (104, "me@work.example.com", "Work", "Empty", 0, 0, ""),
    (105, "me@work.example.com", "Work", "Drafts", 0, 1, "drafts"),
)


class FakeRunner:
    def __init__(self, replies):
        self.replies = dict(replies)
        self.calls = []

    def __call__(self, script, args):
        self.calls.append((script, list(args)))
        reply = self.replies.get(script)
        if callable(reply):
            reply = reply(list(args))
        if isinstance(reply, list):
            reply = reply.pop(0)
        if isinstance(reply, Exception):
            raise reply
        if reply is None:
            raise AssertionError(f"unexpected script {script}")
        return reply

    def last(self, script):
        return [c[1] for c in self.calls if c[0] == script][-1]


@pytest.fixture(autouse=True)
def no_real_accounts_file(tmp_path, monkeypatch):
    """Never read the user's real overrides file in tests."""
    monkeypatch.setenv("OUTLOOK_APP_MCP_ACCOUNTS", str(tmp_path / "accounts.json"))
    return tmp_path / "accounts.json"


@pytest.fixture
def make():
    def _make(replies):
        runner = FakeRunner({"read/folders": FOLDERS, **replies})
        return MacOutlook(runner=runner), runner

    return _make


def message_raw(mid=55, folder_id=101, to=(("Lee", "lee@x"),), atts=(), body="hello", flag="«constant ****FlNF»"):
    header = [
        str(mid), str(folder_id), "Inbox", "me@work.example.com",
        "2026-11-4-36000", "2026-11-4-35900", "false", "true", "נושא",
        a("Rina", "rina@x"),
        GS.join(a(n, ad) for n, ad in to), "", "",
        GS.join(US.join(x) for x in atts),
        flag, GS.join(["Red", "Blue"]), "false",
    ]
    return FS.join(header) + RS + body
