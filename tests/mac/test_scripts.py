"""Static checks on the AppleScripts: they compile, read scripts change nothing,
and each write script uses only the Outlook commands it is meant to."""

import re
import shutil
import subprocess
from pathlib import Path

import pytest

from outlook_app_mcp.mac.osa import COMMON, SCRIPTS, script_source

READ = sorted((SCRIPTS / "read").glob("*.applescript"))
WRITE = sorted((SCRIPTS / "write").glob("*.applescript"))

ACTION = re.compile(
    r"\b(send meeting|send(?! to external\b| only to known\b)|delete|move|duplicate|save|make|forward|reply to|redirect|"
    r"accept meeting|accept tentatively meeting|decline meeting|accept invite|decline invite|"
    r"accept tentatively invite|cancel meeting|get occurrence of|get locked occurrence of|set autoreply data|"
    r"import|export|sync|empty|purge|permanently|"
    r"open|print|quit|do shell script|run script|load script|store script)\b"
    r"|\bset\s+(?!end\s+of\b)([\w ]+?) of \w+ to\b",
    re.IGNORECASE,
)

# Changes the date object it builds; touches nothing in Outlook.
COMMON_ALLOWED = {"set:day", "set:year", "set:month", "set:time"}

WRITE_ALLOWED = {
    "mark": {"set:is read", "set:todo flag", "set:due date"},
    "move": {"move"},
    "delete": {"delete"},
    "compose": {"make", "send", "set:priority"},
    "draft": {"delete", "make", "set:subject", "set:content", "set:priority"},
    "send_draft": {"send"},
    "reply": {"reply to", "forward", "make", "send", "move", "set:content"},
    "folder": {"make", "set:name", "move"},
    "save_attachment": {"save"},
    "event": {"make", "delete", "set:subject", "set:start time", "set:end time", "set:location", "set:content",
              "set:all day flag", "set:free busy status", "set:has reminder", "set:reminder time", "set:is private",
              "set:recurrence", "set:categories", "send meeting", "get occurrence of"},
    "cancel": {"cancel meeting", "get occurrence of", "set:content"},
    "respond": {"accept meeting", "accept tentatively meeting", "decline meeting", "accept invite",
                "accept tentatively invite", "decline invite", "get occurrence of"},
    "sync": {"sync"},
    "task": {"make", "delete", "set:name", "set:content", "set:due date", "set:start date", "set:reminder date time",
             "set:priority", "set:todo flag", "set:categories"},
    "note": {"make", "delete", "set:name", "set:content"},
    "category": {"make", "set:categories"},
    "autoreply": {"set autoreply data"},
    "contact": {"make", "delete", "set:first name", "set:last name", "set:company", "set:job title",
                "set:email addresses", "set:mobile number", "set:business phone number", "set:home phone number",
                "set:plain text note", "set:middle name", "set:nickname", "set:department", "set:office",
                "set:home street address", "set:home city", "set:home state", "set:home zip", "set:home country",
                "set:business street address", "set:business city", "set:business state", "set:business zip",
                "set:business country", "set:business web page", "set:birthday", "set:categories"},
    "group": {"make", "delete", "set:name", "set:members"},
}

NEVER = {"permanently", "empty", "purge", "do shell script", "run script", "load script", "store script",
         "quit", "sync", "import", "export", "open", "print", "duplicate", "redirect", "cancel meeting"}


def actions(text: str) -> set[str]:
    found = set()
    for line in text.splitlines():
        code = line.split("--", 1)[0]
        code = re.sub(r'"(?:[^"\\]|\\.)*"', '""', code)  # drop string literals
        for m in ACTION.finditer(code):
            found.add(f"set:{m.group(2).strip().lower()}" if m.group(2) else m.group(1).lower())
    return found


def test_guard_sees_writes_and_ignores_reads():
    assert actions("set is read of m to true") == {"set:is read"}
    assert actions("set end time of e to dv") == {"set:end time"}
    assert actions('delete m\nsend m\nmove m to f\nmake new outgoing message') == {"delete", "send", "move", "make"}
    assert actions('do shell script "x"') == {"do shell script"}
    assert actions("send meeting e\nset o to get occurrence of e at d") == {"send meeting", "get occurrence of"}
    # Out-of-office record fields, not the send command.
    assert actions("set x to send to external of r\nset y to send only to known of r") == set()
    quiet = 'set end of out to x\nset m to message id 3\nset ids to id of every message of f\nif mode is "forward" then\n-- delete m'
    assert actions(quiet) == set()


def test_common_only_builds_dates():
    assert actions(COMMON.read_text()) <= COMMON_ALLOWED


@pytest.mark.parametrize("path", READ, ids=lambda p: p.name)
def test_read_scripts_change_nothing(path):
    assert actions(path.read_text()) == set()


def test_every_write_script_is_declared():
    assert {p.stem for p in WRITE} == set(WRITE_ALLOWED)


@pytest.mark.parametrize("path", WRITE, ids=lambda p: p.name)
def test_write_scripts_use_only_their_commands(path):
    used = actions(path.read_text())
    assert used <= WRITE_ALLOWED[path.stem], used - WRITE_ALLOWED[path.stem]
    assert not (used & (NEVER - WRITE_ALLOWED[path.stem]))


def test_rare_commands_have_one_home():
    """Syncing and cancelling a meeting each happen in exactly one script."""
    for cmd, home in (("sync", "sync"), ("cancel meeting", "cancel"), ("set autoreply data", "autoreply")):
        assert {k for k, v in WRITE_ALLOWED.items() if cmd in v} == {home}


def test_never_list_holds_everything_else():
    for name, allowed in WRITE_ALLOWED.items():
        assert not (allowed & (NEVER - {"sync", "cancel meeting"})), name


# Compiling needs Outlook's scripting dictionary, so Outlook itself (CI runners have none).
OUTLOOK_APP = Path("/Applications/Microsoft Outlook.app")


@pytest.mark.skipif(shutil.which("osacompile") is None or not OUTLOOK_APP.exists(), reason="needs macOS with Outlook")
@pytest.mark.parametrize("path", READ + WRITE, ids=lambda p: f"{p.parent.name}/{p.name}")
def test_script_compiles_against_outlook(path, tmp_path):
    src = tmp_path / "s.applescript"
    src.write_text(script_source(f"{path.parent.name}/{path.stem}"), encoding="utf-8")
    proc = subprocess.run(["osacompile", "-o", str(tmp_path / "s.scpt"), str(src)], check=False, capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr
