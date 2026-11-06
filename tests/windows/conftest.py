"""Shared set-up for the Windows tests that go beyond the 0.2 feature set: the same fake Outlook as
test_windows.build(), plus Tasks and Notes folders, a few tasks and notes, a meeting request, a directory."""

from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace

import pytest
from fakecom import AddressEntry, MeetingRequest, Note, Task

from outlook_app_mcp.windows import WindowsOutlook


@pytest.fixture
def w(monkeypatch):
    from test_windows import build

    monkeypatch.setattr(WindowsOutlook, "_jet", staticmethod(lambda dt: dt.strftime("%d/%m/%Y %H:%M")))
    o, app, r = build()
    tasks = r.a.root.add("Tasks", kind=3)
    notes = r.a.root.add("Notes", kind=5)
    r.a.defaults[13] = tasks
    r.a.defaults[12] = notes
    rent = Task("Pay rent", due=datetime(2026, 11, 10))
    old = Task("Old chore", due=datetime(2026, 10, 1), complete=True)
    undated = Task("No date", importance=2)
    undated.Categories = "Home"
    request = Task("Task request")
    request.MessageClass = "IPM.TaskRequest"
    tasks.put(rent, old, undated, request)
    shopping = Note(Body="Shopping\r\nmilk\r\neggs")
    shopping.set_local("LastModificationTime", datetime(2026, 11, 1, 10, 0))
    idea = Note(Body="Idea")
    idea.set_local("LastModificationTime", datetime(2026, 11, 3, 9, 30))
    notes.put(shopping, idea)
    invite = MeetingRequest("Planning", sender=("Boss", "boss@x.co"), received=datetime.now(), appointment=r.meeting)
    cancelled = MeetingRequest("Cancelled: Planning", sender=("Boss", "boss@x.co"), received=datetime.now(),
                               message_class="IPM.Schedule.Meeting.Canceled")
    r.inbox.put(invite, cancelled)
    extra = SimpleNamespace(tasks=tasks, notes=notes, rent=rent, old_task=old, undated=undated, request=request,
                            shopping=shopping, idea=idea, invite=invite, cancelled=cancelled)
    refs = SimpleNamespace(**{**vars(r), **vars(extra)})
    return o, app, refs


@pytest.fixture
def directory(w):
    _, app, _ = w
    app.gal = [
        AddressEntry("Dana Levi", "dana@corp.example", "CTO", "R&D", "Tel Aviv", "03-5551234"),
        AddressEntry("Danny Cohen", "danny@corp.example", "Analyst", "Finance"),
        AddressEntry("All Staff", is_list=True),
    ]
    return app.gal
