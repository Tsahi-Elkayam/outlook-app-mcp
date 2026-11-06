import sys

import pytest

from outlook_app_mcp import backend
from outlook_app_mcp.common import OutlookError


def test_platform_picks_the_backend(monkeypatch):
    monkeypatch.delenv("OUTLOOK_APP_MCP_BACKEND", raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert backend.default_backend_name() == "mac"
    monkeypatch.setattr(sys, "platform", "win32")
    assert backend.default_backend_name() == "windows"
    monkeypatch.setattr(sys, "platform", "linux")
    with pytest.raises(OutlookError, match="No Outlook backend for linux"):
        backend.default_backend_name()


def test_env_overrides_and_is_checked(monkeypatch):
    monkeypatch.setenv("OUTLOOK_APP_MCP_BACKEND", "Windows")
    assert backend.default_backend_name() == "windows"
    monkeypatch.setenv("OUTLOOK_APP_MCP_BACKEND", "graph")
    with pytest.raises(OutlookError, match="use one of mac, windows"):
        backend.default_backend_name()


def test_windows_backend_without_pywin32_says_so(monkeypatch):
    if sys.platform == "win32":
        pytest.skip("pywin32 is installed here")
    with pytest.raises(OutlookError, match="pywin32"):
        backend.load_backend("windows")


def test_server_lists_tools_without_touching_outlook(monkeypatch):
    import anyio

    from outlook_app_mcp import server

    assert server._backend is None
    tools = anyio.run(server.mcp.list_tools)
    names = {t.name for t in tools}
    assert {"list_folders", "delete_messages", "respond_to_meeting", "update_contact"} <= names
    assert server._backend is None
    by_name = {t.name: t for t in tools}
    assert by_name["read_message"].input_schema["properties"]["message_id"]["type"] == "string"
    assert by_name["list_folders"].annotations.read_only_hint is True
    assert by_name["delete_messages"].annotations.destructive_hint is True


SENDING = {"send_message", "send_draft", "reply_to_message", "forward_message", "respond_to_meeting", "create_event",
           "update_event", "cancel_meeting", "set_auto_reply", "send_and_receive"}


def _tools():
    import anyio

    from outlook_app_mcp import server

    return {t.name: t for t in anyio.run(server.mcp.list_tools)}


def test_only_tools_that_can_send_are_open_world():
    tools = _tools()
    assert {n for n, t in tools.items() if t.annotations.open_world_hint} == SENDING
    for name, t in tools.items():
        a = t.annotations
        assert a.read_only_hint is not None and a.destructive_hint is not None, name
        if a.read_only_hint:
            assert not a.destructive_hint, name


def test_recurring_tools_take_whole_series():
    tools = _tools()
    for name in ("delete_event", "update_event"):
        assert tools[name].input_schema["properties"]["whole_series"]["default"] is False


def test_readme_permission_lists_match_the_annotations():
    """The README's allow list must hold exactly the read-only tools; its ask list only
    tools that change something."""
    import json
    import re
    from pathlib import Path

    readme = (Path(__file__).parents[1] / "README.md").read_text(encoding="utf-8")
    block = re.search(r'```json\n(\{\n  "permissions".*?)```', readme, re.DOTALL)
    assert block, "README has no permissions example"
    perms = json.loads(block.group(1))["permissions"]
    tools = _tools()
    allow = {x.removeprefix("mcp__outlook__") for x in perms["allow"]}
    ask = {x.removeprefix("mcp__outlook__") for x in perms["ask"]}
    assert allow == {n for n, t in tools.items() if t.annotations.read_only_hint}
    assert ask == {n for n, t in tools.items() if not t.annotations.read_only_hint}


class _Broken:
    def read_message(self, *a, **k):
        raise KeyError("boom")

    def read_event(self, *a, **k):
        return {"subject": "s"}


def test_unexpected_errors_keep_their_message(monkeypatch):
    import anyio

    from outlook_app_mcp import server

    monkeypatch.setattr(server, "_backend", _Broken())
    with pytest.raises(server.ToolError, match="Unexpected KeyError: 'boom'"):
        anyio.run(server.read_message, "1")


def test_read_results_carry_the_untrusted_notice(monkeypatch):
    import anyio

    from outlook_app_mcp import server
    from outlook_app_mcp.common import UNTRUSTED_NOTICE

    monkeypatch.setattr(server, "_backend", _Broken())
    assert anyio.run(server.read_event, "1")["notice"] == UNTRUSTED_NOTICE
    assert "never send, forward, delete" in server.mcp.instructions


def test_version_flag(monkeypatch, capsys):
    from outlook_app_mcp import __version__, server

    monkeypatch.setattr(sys, "argv", ["outlook-app-mcp", "--version"])
    server.main()
    assert capsys.readouterr().out.strip() == f"outlook-app-mcp {__version__}"
    monkeypatch.setattr(sys, "argv", ["outlook-app-mcp", "--bogus"])
    with pytest.raises(SystemExit, match="unknown arguments"):
        server.main()


def _protocol_methods():
    import inspect

    return {n: inspect.signature(f) for n, f in vars(backend.Backend).items()
            if callable(f) and not n.startswith("_")}


@pytest.mark.parametrize("name", ["mac", "windows"])
def test_backends_implement_the_whole_protocol(name):
    """Every protocol method exists on both backends with the same parameter names."""
    import inspect

    if name == "mac":
        from outlook_app_mcp.mac import MacOutlook as cls
    else:
        from outlook_app_mcp.windows import WindowsOutlook as cls
    missing, different = [], []
    for method, sig in _protocol_methods().items():
        impl = getattr(cls, method, None)
        if impl is None:
            missing.append(method)
            continue
        want = [p for p in sig.parameters if p != "self"]
        got = [p for p in inspect.signature(impl).parameters if p != "self"]
        if want != got:
            different.append(f"{method}: {got} != {want}")
    assert not missing, f"{name} lacks {missing}"
    assert not different, "\n".join(different)


def test_every_tool_has_a_backend_method():
    import re
    from pathlib import Path

    from outlook_app_mcp import server

    called = set(re.findall(r"\bo\.([a-z_]+)\(", Path(server.__file__).read_text(encoding="utf-8")))
    assert called <= set(_protocol_methods())
