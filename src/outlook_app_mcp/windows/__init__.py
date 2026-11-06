"""Windows backend: classic Outlook for Windows, driven through COM (pywin32).

Mail, calendar, contacts, tasks and notes, read and write, with the same safety rules as the
macOS backend: nothing deletes permanently, replies and forwards stay drafts
unless asked, meetings with attendees can't be deleted.
"""

from __future__ import annotations

from .calendar import WinCalendarMixin
from .com import com_error_type, open_outlook
from .contacts import WinContactsMixin
from .mail import WinMailMixin
from .misc import WinMiscMixin
from .notes import WinNotesMixin
from .tasks import WinTasksMixin


class WindowsOutlook(WinMailMixin, WinCalendarMixin, WinContactsMixin, WinMiscMixin, WinTasksMixin, WinNotesMixin):
    name = "windows"
    description = "classic Outlook for Windows, through COM"

    def __init__(self, app=None, com_error: type[BaseException] | None = None, connect=None):
        # connect() makes a new Outlook.Application; used again after Outlook restarts.
        self._connect = connect if connect is not None else (open_outlook if app is None else None)
        self._app = app if app is not None else self._connect()
        self._com_error = com_error or com_error_type()
        self._reset()

    def _reset(self) -> None:
        self._stores = None
        self._folders = None
        self._folders_at = 0.0
        self._accounts = {}

    def _reconnect(self) -> bool:
        """Connect to Outlook again after it was closed or restarted. False if this
        backend can't (a test double, or Outlook won't start)."""
        if self._connect is None:
            return False
        try:
            self._app = self._connect()
        except Exception:  # noqa: BLE001 - the original error is reported instead
            return False
        self._reset()
        return True

    @property
    def outlook_window_open(self) -> bool:
        """False when Outlook runs without a window (started by this server). Mail sent
        then can sit in the Outbox until Outlook is opened or a send/receive runs."""
        try:
            return self._app.Explorers.Count > 0
        except Exception:  # noqa: BLE001 - unknown: assume open
            return True
