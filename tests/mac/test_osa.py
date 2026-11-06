from datetime import UTC, datetime

import pytest

from outlook_app_mcp.mac import osa
from outlook_app_mcp.mac.osa import (
    OutlookError,
    check_ids,
    enum,
    flag,
    parse_stamp,
    to_stamp,
)


def test_stamp_roundtrip_is_exact_local_time():
    dt = parse_stamp("2026-11-10-36000")
    assert (dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second) == (2026, 11, 10, 10, 0, 0)
    assert to_stamp(dt) == "2026-11-10-36000"
    assert parse_stamp("") is None


def test_to_stamp_converts_aware_times_to_local():
    utc = datetime(2026, 11, 10, 8, 0, tzinfo=UTC)
    local = utc.astimezone()
    secs = local.hour * 3600 + local.minute * 60
    assert to_stamp(utc) == f"{local.year}-{local.month}-{local.day}-{secs}"




def test_enum_both_forms():
    assert enum("busy") == "busy"
    assert enum("«constant ****eSBu»") == "busy"
    assert enum("«constant FlageSBu»") == "busy"
    assert flag("«constant ****FlNC»") == "flagged"
    assert flag("completed") == "completed"
    assert flag("not flagged") is None






def test_mac_ids_are_numbers_as_text():
    assert check_ids([5, "6", " 7 "]) == ["5", "6", "7"]
    with pytest.raises(OutlookError, match="not an Outlook for Mac id"):
        check_ids(["AAMkAD="])
    with pytest.raises(OutlookError, match="empty"):
        check_ids([])


def test_explain_maps_known_errors():
    assert "Automation" in osa.explain("execution error: Not authorized (-1743)")
    assert "could not find" in osa.explain("Can't get message id 5. (-1728)")
    assert "respond_to_meeting" in osa.explain("HAS_ATTENDEES: event 1")


@pytest.mark.skipif(not __import__("shutil").which("osascript"), reason="needs macOS")
def test_real_osascript_passes_argv_untouched(tmp_path, monkeypatch):
    """Runs a real script through run_osascript: common is prepended, argv arrives verbatim,
    and an AppleScript-injection string stays plain data."""
    (tmp_path / "read").mkdir()
    (tmp_path / "read" / "echo.applescript").write_text(
        "on run argv\n my initSeps()\n return my joinList(argv, RS)\nend run\n", encoding="utf-8"
    )
    monkeypatch.setattr(osa, "SCRIPTS", tmp_path)
    evil = '" & (do shell script "echo pwned") & "'
    out = osa.run_osascript("read/echo", ["שלום", evil, "a\nb"])
    assert out.split(osa.RS) == ["שלום", evil, "a\nb"]


@pytest.mark.skipif(not __import__("shutil").which("osascript"), reason="needs macOS")
def test_real_osascript_error_is_explained(tmp_path, monkeypatch):
    (tmp_path / "read").mkdir()
    (tmp_path / "read" / "boom.applescript").write_text('on run argv\n error "nope" number -1728\nend run\n')
    monkeypatch.setattr(osa, "SCRIPTS", tmp_path)
    with pytest.raises(OutlookError, match="could not find"):
        osa.run_osascript("read/boom", [])


def test_unknown_script_name():
    with pytest.raises(OutlookError, match="No such script"):
        osa.script_source("read/nope")


def test_timeout_on_a_write_warns_it_may_have_happened(monkeypatch):
    """A killed osascript doesn't undo an Apple Event Outlook already got: a write that
    timed out must not be retried blindly (a second send, a second delete)."""
    import subprocess

    def slow(*a, **kw):
        raise subprocess.TimeoutExpired("osascript", 1)

    monkeypatch.setattr(subprocess, "run", slow)
    with pytest.raises(OutlookError, match="may still have happened"):
        osa.run_osascript("write/compose", [])
    with pytest.raises(OutlookError, match="Narrow the request"):
        osa.run_osascript("read/folders", [])


def test_undecodable_output_is_replaced_not_fatal(monkeypatch):
    import subprocess

    seen = {}

    def run(*a, **kw):
        seen.update(kw)
        return subprocess.CompletedProcess(a, 0, "ok\n", "")

    monkeypatch.setattr(subprocess, "run", run)
    assert osa.run_osascript("read/folders", []) == "ok"
    assert seen["errors"] == "replace"
