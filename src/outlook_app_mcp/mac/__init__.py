"""macOS backend: Microsoft Outlook for Mac (Legacy mode), driven through AppleScript.

Mail, calendar and contacts, read and write. Every operation runs a fixed
script from ``scripts/read`` or ``scripts/write``; nothing runs free-form
AppleScript, nothing deletes permanently.
"""

from __future__ import annotations

from .calendar import CalendarMixin
from .contacts import ContactsMixin
from .mail import MailMixin
from .misc import MiscMixin
from .osa import Runner, run_osascript
from .tasks import TasksMixin


class MacOutlook(MailMixin, CalendarMixin, ContactsMixin, MiscMixin, TasksMixin):
    name = "mac"
    description = "Microsoft Outlook for Mac in Legacy Outlook mode, through AppleScript"

    def __init__(self, runner: Runner = run_osascript):
        self._run = runner
        self._folders = None
        self._folders_at = 0.0
        self._google = None
        self._mailboxes = {}
