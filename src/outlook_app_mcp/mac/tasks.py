"""macOS tasks and notes."""

from __future__ import annotations

from datetime import datetime

from ..common import (
    DEFAULT_LIMIT,
    DELETE_NOTE,
    OutlookError,
    check_categories,
    check_date_or_clear,
    check_priority,
    check_title,
    clamp_limit,
    parse_when,
    to_html,
)
from .osa import GS, RS, check_id, enum, priority, records, stamp_iso, to_stamp

NOT_EXCHANGE_DELETE = (
    "{what} {id} is in {where}, not in an Exchange account, and Outlook deletes such items for good. "
    "Nothing was deleted; delete it in Outlook if you mean to."
)


def _day(stamp: str) -> str | None:
    """A task date as YYYY-MM-DD (Outlook keeps tasks to the day)."""
    v = stamp_iso(stamp) if stamp else None
    return v[:10] if v else None


def _folders(rows: list[list[str]]) -> dict[str, dict]:
    return {f[1]: {"name": f[2], "account": f[3], "kind": f[4] if len(f) > 4 else ""} for f in rows if f[0] == "F"}


def _task(f: list[str], folders: dict[str, dict]) -> dict:
    folder = folders.get(f[8], {})
    done = enum(f[6]) == "completed"
    return {
        "id": f[1],
        "subject": f[2],
        "due": _day(f[3]),
        "start": _day(f[4]),
        "completed": done,
        "completed_on": _day(f[5]) if done else None,
        "priority": priority(f[7]) or "normal",
        "categories": f[9].split(GS) if len(f) > 9 and f[9] else [],
        "folder": folder.get("name"),
        "account": folder.get("account") or None,
    }


def _date_stamp(value: datetime) -> str:
    """Midnight of the day, local: Outlook keeps task dates as days."""
    local = value.astimezone()
    return to_stamp(local.replace(hour=0, minute=0, second=0, microsecond=0))


class TasksMixin:
    """Needs self._run, require_accounts and _ensure_categories."""

    # -- tasks -----------------------------------------------------------------------

    def _task_rows(self) -> tuple[dict[str, dict], list[list[str]]]:
        self.require_accounts()
        rows = records(self._run("read/tasks", []))
        return _folders(rows), [f for f in rows if f[0] == "T"]

    def list_tasks(self, include_completed: bool = False, query: str | None = None, due_before: str | None = None,
                   limit: int = DEFAULT_LIMIT) -> dict:
        limit = clamp_limit(limit)
        before = parse_when(due_before, "due_before").astimezone().date().isoformat() if due_before else None
        needle = (query or "").strip().casefold()
        folders, rows = self._task_rows()
        found = []
        for f in rows:
            t = _task(f, folders)
            if t["completed"] and not include_completed:
                continue
            if needle and needle not in t["subject"].casefold() and not any(
                    needle in c.casefold() for c in t["categories"]):
                continue
            if before and not (t["due"] and t["due"] < before):
                continue
            found.append(t)
        found.sort(key=lambda t: (t["due"] is None, t["due"] or "", t["subject"].casefold()))
        return {"count": len(found), "returned": min(limit, len(found)), "has_more": len(found) > limit,
                "tasks": found[:limit]}

    def read_task(self, task_id: str) -> dict:
        task_id = check_id(task_id, "task_id")
        folders, _ = self._task_rows()
        header, _, body = self._run("read/task", [task_id]).partition(RS)
        f = header.split("\x1f")
        task = _task(f, folders)
        task["body"] = body
        task["reminder"] = stamp_iso(f[10]) if len(f) > 10 and f[10] else None
        return task

    def _task_pairs(self, subject, body, due, start, reminder, remove_reminder, priority_, complete, cats) -> list[str]:
        pairs: list[str] = []
        if subject is not None:
            pairs += ["subject", check_title(subject)]
        if body is not None:
            pairs += ["html", to_html(body)]
        due_v = check_date_or_clear(due, "due")
        start_v = check_date_or_clear(start, "start")
        if isinstance(due_v, datetime) and isinstance(start_v, datetime) and start_v.date() > due_v.date():
            raise OutlookError("start must not be after due.")
        if due_v is not None:
            pairs += ["due", _date_stamp(due_v) if due_v else ""]
        if start_v is not None:
            pairs += ["start", _date_stamp(start_v) if start_v else ""]
        if remove_reminder:
            pairs += ["reminder", ""]
        elif reminder:
            pairs += ["reminder", to_stamp(parse_when(reminder, "reminder"))]
        if priority_ is not None:
            pairs += ["priority", check_priority(priority_)]
        if complete is not None:
            pairs += ["complete", "1" if complete else "0"]
        if cats is not None:
            pairs += ["categories", GS.join(self._ensure_categories(cats))]
        return pairs

    def _task_folder_for(self, account: str | None) -> str:
        if not account:
            return ""
        folders, _ = self._task_rows()
        want = account.strip().casefold()
        hits = [fid for fid, f in folders.items() if f["account"] and want in f["account"].casefold()]
        if not hits:
            raise OutlookError(f"No task folder of an account matching {account!r}. Accounts with tasks: "
                               f"{', '.join(sorted({f['account'] for f in folders.values() if f['account']}))}")
        return min(hits, key=int)

    def create_task(self, subject, body=None, due=None, start=None, reminder=None, priority=None, categories=None,
                    account=None) -> dict:
        subject = check_title(subject)
        cats = check_categories(categories)
        pairs = self._task_pairs(subject, body, due, start, reminder, False, priority, None, cats)
        folder = self._task_folder_for(account)
        self.require_accounts()
        tid = self._run("write/task", ["create", folder, *pairs]).strip()
        return self.read_task(tid)

    def update_task(self, task_id, subject=None, body=None, due=None, start=None, reminder=None,
                    remove_reminder=False, priority=None, complete=None, categories=None) -> dict:
        task_id = check_id(task_id, "task_id")
        cats = check_categories(categories)
        pairs = self._task_pairs(subject, body, due, start, reminder, remove_reminder, priority, complete, cats)
        if not pairs:
            raise OutlookError("Nothing to change.")
        current = self.read_task(task_id)
        new_due = current["due"] if due is None else (check_date_or_clear(due, "due") or None)
        new_start = current["start"] if start is None else (check_date_or_clear(start, "start") or None)
        day = lambda v: v.astimezone().date().isoformat() if isinstance(v, datetime) else v
        if new_due and new_start and day(new_start) > day(new_due):
            raise OutlookError("start must not be after due (counting the task's current dates).")
        self._run("write/task", ["update", task_id, *pairs])
        return self.read_task(task_id)

    def delete_task(self, task_id: str) -> dict:
        task_id = check_id(task_id, "task_id")
        folders, rows = self._task_rows()
        row = next((f for f in rows if f[1] == task_id), None)
        if row is None:
            raise OutlookError(f"No task with id {task_id}.")
        self._check_exchange_folder(folders.get(row[8], {}), "Task", task_id)
        self._run("write/task", ["delete", task_id])
        return {"deleted": task_id, "subject": row[2], "note": DELETE_NOTE}

    @staticmethod
    def _check_exchange_folder(folder: dict, what: str, item_id: str) -> None:
        if folder.get("kind") != "exchange":
            where = f"the folder {folder.get('name')!r} of {folder.get('account') or 'On My Computer'}"
            raise OutlookError(NOT_EXCHANGE_DELETE.format(what=what, id=item_id, where=where))

    # -- notes -----------------------------------------------------------------------

    def _note_rows(self) -> tuple[dict[str, dict], list[list[str]]]:
        self.require_accounts()
        rows = records(self._run("read/notes", []))
        return _folders(rows), [f for f in rows if f[0] == "N"]

    def list_notes(self, query: str | None = None, limit: int = DEFAULT_LIMIT) -> dict:
        limit = clamp_limit(limit)
        needle = (query or "").strip().casefold()
        folders, rows = self._note_rows()
        found = []
        for f in rows:
            folder = folders.get(f[4], {})
            note = {"id": f[1], "title": f[2], "modified": stamp_iso(f[3]) if f[3] else None,
                    "preview": f[5] if len(f) > 5 else "", "folder": folder.get("name"),
                    "account": folder.get("account") or None}
            if needle and needle not in (note["title"] + " " + note["preview"]).casefold():
                continue
            found.append(note)
        found.sort(key=lambda n: n["modified"] or "", reverse=True)
        return {"count": len(found), "returned": min(limit, len(found)), "has_more": len(found) > limit,
                "notes": found[:limit]}

    def read_note(self, note_id: str) -> dict:
        note_id = check_id(note_id, "note_id")
        self.require_accounts()
        header, _, body = self._run("read/note", [note_id]).partition(RS)
        f = header.split("\x1f")
        return {"id": f[0], "title": f[1], "body": body, "modified": stamp_iso(f[2]) if f[2] else None,
                "categories": f[4].split(GS) if len(f) > 4 and f[4] else []}

    def _note_folder_for(self, account: str | None) -> str:
        if not account:
            return ""
        folders, _ = self._note_rows()
        want = account.strip().casefold()
        hits = [fid for fid, f in folders.items() if f["account"] and want in f["account"].casefold()]
        if not hits:
            raise OutlookError(f"No notes folder of an account matching {account!r}.")
        return min(hits, key=int)

    def create_note(self, title: str, body: str = "", account: str | None = None) -> dict:
        title = check_title(title, "title")
        folder = self._note_folder_for(account)
        self.require_accounts()
        nid = self._run("write/note", ["create", folder, "title", title, "html", to_html(body or "")]).strip()
        return self.read_note(nid)

    def update_note(self, note_id: str, title: str | None = None, body: str | None = None) -> dict:
        note_id = check_id(note_id, "note_id")
        pairs: list[str] = []
        if title is not None:
            pairs += ["title", check_title(title, "title")]
        if body is not None:
            pairs += ["html", to_html(body)]
        if not pairs:
            raise OutlookError("Nothing to change.")
        self.read_note(note_id)
        self._run("write/note", ["update", note_id, *pairs])
        return self.read_note(note_id)

    def delete_note(self, note_id: str) -> dict:
        note_id = check_id(note_id, "note_id")
        folders, rows = self._note_rows()
        row = next((f for f in rows if f[1] == note_id), None)
        if row is None:
            raise OutlookError(f"No note with id {note_id}.")
        self._check_exchange_folder(folders.get(row[4], {}), "Note", note_id)
        self._run("write/note", ["delete", note_id])
        return {"deleted": note_id, "title": row[2], "note": DELETE_NOTE}
