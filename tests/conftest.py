"""Shared by every test: files may be attached from and saved to the test's own
temporary folder only, never the real ~/Downloads, ~/Desktop or ~/Documents."""

import pytest


@pytest.fixture(autouse=True)
def file_dirs(tmp_path, monkeypatch):
    monkeypatch.setenv("OUTLOOK_APP_MCP_FILE_DIRS", str(tmp_path))
    return tmp_path
