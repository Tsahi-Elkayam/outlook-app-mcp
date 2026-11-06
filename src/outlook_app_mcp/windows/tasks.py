"""Windows tasks: Outlook's Tasks folders (not the flagged-mail To-Do list, and not task requests)."""

from __future__ import annotations

from datetime import date

from ..common import (
    ALREADY_DELETED,
    DEFAULT_LIMIT,
    DELETE_NOTE,
    OutlookError,
    check_categories,
    check_date_or_clear,
    check_priority,
    check_title,
    clamp_limit,
    parse_when,
)
from .com import (
    FOLDER_TASKS,
    IMPORTANCE,
    IMPORTANCE_CODES,
    ITEM_TASK,
    NO_DATE,
    com_errors,
    com_iso,
    date_com,
    day_to_com,
    from_com,
    make_id,
    prop,
    split_categories,
    to_com,
)
from .mail import table_rows

TASK_COLUMNS = ("EntryID", "MessageClass", "Subject", "DueDate", "StartDate", "Complete", "DateCompleted",
                "Importance", "Categories")


def is_task_class(message_class: str) -> bool:
    """A task itself; task requests and responses (IPM.TaskRequest...) arrive in the Inbox and are mail."""
    return (message_class or "").startswith("IPM.Task") and not (message_class or "").startswith("IPM.TaskRequest")


class WinTasksMixin:
    """Needs the store helpers of WinMailMixin and _kind_folders / _default_folder of WinMiscMixin."""

    @staticmethod
    def _task_summary(values: dict, key: str, folder_name: str, account: str) -> dict:
        return {
            "id": make_id(key, values["EntryID"]),
            "subject": values.get("Subject") or "",
            "due": date_com(values.get("DueDate")),
            "start": date_com(values.get("StartDate")),
            "completed": bool(values.get("Complete")),
            "completed_on": date_com(values.get("DateCompleted")) if values.get("Complete") else None,
            "priority": IMPORTANCE.get(values.get("Importance"), "normal"),
            "categories": split_categories(values.get("Categories")),
            "folder": folder_name,
            "account": account,
        }

    @com_errors
    def list_tasks(self, include_completed: bool = False, query: str | None = None, due_before: str | None = None,
                   limit: int = DEFAULT_LIMIT) -> dict:
        """Tasks from every Tasks folder, soonest due first (undated last). query matches the subject and the
        categories; due_before keeps tasks due on an earlier day (undated tasks are then left out)."""
        limit = clamp_limit(limit)
        before = parse_when(due_before, "due_before").astimezone().date() if due_before else None
        needle = (query or "").strip().casefold()
        found = []
        for store, folder in self._kind_folders(ITEM_TASK):
            table = folder.GetTable()
            cols = table.Columns
            cols.RemoveAll()
            for c in TASK_COLUMNS:
                cols.Add(c)
            for row in table_rows(table, TASK_COLUMNS, None):
                if not is_task_class(row.get("MessageClass")):
                    continue
                t = self._task_summary(row, store.key, prop(folder, "Name", "") or "", store.account)
                if t["completed"] and not include_completed:
                    continue
                if needle and needle not in t["subject"].casefold() and not any(
                        needle in c.casefold() for c in t["categories"]):
                    continue
                if before and not (t["due"] and date.fromisoformat(t["due"]) < before):
                    continue
                found.append(t)
        found.sort(key=lambda t: (t["due"] is None, t["due"] or "", t["subject"].casefold()))
        return {"count": len(found), "returned": min(limit, len(found)), "has_more": len(found) > limit,
                "tasks": found[:limit]}

    def _task_item(self, task_id: str, what: str = "task_id"):
        it, key = self._item(task_id, what)
        if not is_task_class(prop(it, "MessageClass", "") or ""):
            raise OutlookError(f"{task_id} is not a task ({prop(it, 'MessageClass', '?')}).")
        return it, key

    def _task_detail(self, it, key: str) -> dict:
        values = {c: prop(it, c) for c in TASK_COLUMNS}
        values["EntryID"] = it.EntryID
        parent = prop(it, "Parent")
        task = self._task_summary(values, key, prop(parent, "Name", "") or "", self._store(key).account)
        task["body"] = prop(it, "Body", "") or ""
        task["reminder"] = com_iso(prop(it, "ReminderTime")) if prop(it, "ReminderSet", False) else None
        return task

    @com_errors
    def read_task(self, task_id: str) -> dict:
        self.require_accounts()
        it, key = self._task_item(task_id)
        return self._task_detail(it, key)

    @staticmethod
    def _task_dates_ok(start, due) -> None:
        if start and due and day_to_com(start) > day_to_com(due):
            raise OutlookError("start must not be after due.")

    def _write_task(self, it, subject, body, due, start, reminder, remove_reminder, level, complete, cats) -> None:
        """Put the given fields on a task. due/start: None leaves, "" clears, a datetime sets the day."""
        if subject is not None:
            it.Subject = subject
        if body is not None:
            it.Body = body
        if due is not None:
            it.DueDate = NO_DATE if due == "" else day_to_com(due)
        if start is not None:
            it.StartDate = NO_DATE if start == "" else day_to_com(start)
        if remove_reminder:
            it.ReminderSet = False
        elif reminder is not None:
            it.ReminderSet = True
            it.ReminderTime = to_com(reminder)
        if level is not None:
            it.Importance = IMPORTANCE_CODES[level]
        if complete is not None:
            it.Complete = bool(complete)
        if cats is not None:
            self._apply_categories(it, cats)

    @com_errors
    def create_task(self, subject: str, body: str | None = None, due: str | None = None, start: str | None = None,
                    reminder: str | None = None, priority: str | None = None, categories=None,
                    account: str | None = None) -> dict:
        subject = check_title(subject)
        due_dt, start_dt = parse_when(due, "due") if due else None, parse_when(start, "start") if start else None
        remind = parse_when(reminder, "reminder") if reminder else None
        level, cats = check_priority(priority), check_categories(categories)
        self._task_dates_ok(start_dt, due_dt)
        folder, key = self._default_folder(FOLDER_TASKS, account, "Tasks")
        it = folder.Items.Add(ITEM_TASK)
        self._write_task(it, subject, body, due_dt, start_dt, remind, False, level, None, cats)
        it.Save()
        return self._task_detail(it, key)

    @com_errors
    def update_task(self, task_id: str, subject: str | None = None, body: str | None = None, due: str | None = None,
                    start: str | None = None, reminder: str | None = None, remove_reminder: bool = False,
                    priority: str | None = None, complete: bool | None = None, categories=None) -> dict:
        if subject is not None:
            subject = check_title(subject)
        due_c, start_c = check_date_or_clear(due, "due"), check_date_or_clear(start, "start")
        if reminder == "":
            remove_reminder, reminder = True, None
        if reminder and remove_reminder:
            raise OutlookError("Give reminder or remove_reminder, not both.")
        remind = parse_when(reminder, "reminder") if reminder else None
        level, cats = check_priority(priority), check_categories(categories)
        if all(v is None for v in (subject, body, due_c, start_c, remind, level, complete, cats)) \
                and not remove_reminder:
            raise OutlookError("Nothing to change.")
        it, key = self._task_item(task_id)
        if self._under_deleted(it, key):
            raise OutlookError("That task is in Deleted Items; restore it in Outlook before changing it.")
        new_start = start_c if start_c is not None else from_com(prop(it, "StartDate"))
        new_due = due_c if due_c is not None else from_com(prop(it, "DueDate"))
        self._task_dates_ok(new_start or None, new_due or None)
        self._write_task(it, subject, body, due_c, start_c, remind, remove_reminder, level, complete, cats)
        it.Save()
        return self._task_detail(it, key)

    @com_errors
    def delete_task(self, task_id: str) -> dict:
        it, key = self._task_item(task_id)
        if self._under_deleted(it, key):
            raise OutlookError(f"Task {task_id} {ALREADY_DELETED}")
        subject = prop(it, "Subject", "") or ""
        it.Move(self._deleted_folder_obj(key))  # Moved, never Delete(): Delete inside Deleted Items is permanent.
        return {"deleted": task_id, "subject": subject, "note": DELETE_NOTE}
