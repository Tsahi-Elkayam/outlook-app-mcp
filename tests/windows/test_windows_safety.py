"""Static checks on the Windows backend's source: reading functions change nothing in
Outlook, and each changing function uses only the object-model calls it is meant to.

Works on the syntax tree, like test_scripts.py does for the AppleScripts: every
reference to a changing method (called or not, so `item.Reply` in a lookup table
counts) and every property assignment on an Outlook object is collected per method.
"""

import ast
from pathlib import Path

import pytest

import outlook_app_mcp.windows as win

SOURCES = sorted(Path(win.__file__).parent.glob("*.py"))

CHANGERS = {
    "Save", "Send", "Move", "Delete", "PermanentDelete", "Add", "Remove", "RemoveAll", "Reply", "ReplyAll",
    "Forward", "Respond", "SaveAsFile", "ResolveAll", "Close", "Copy", "Display", "Execute", "ExecuteMethod",
    "ClearTaskFlag", "MarkAsTask", "Logoff", "Quit", "SetProperty", "DeleteProperty", "ShowCategoriesDialog",
    "Invoke", "SendAndReceive", "MoveTo", "AddMember", "RemoveMember",
}
# Collection methods: named with what they act on ("Items.Add" makes an item, "cols.Add" adds a table column).
QUALIFIED = {"Add", "Remove", "RemoveAll"}

# Shaping a table or collection view; touches no item.
VIEW = {"cols.Add", "cols.RemoveAll", "set:IncludeRecurrences"}

NEVER = {"Delete", "PermanentDelete", "Execute", "ExecuteMethod", "Copy", "Display", "Logoff", "Quit",
         "SetProperty", "DeleteProperty", "Remove"}

WRITE_ALLOWED = {
    "mark_messages": {"set:UnRead", "set:FlagStatus", "set:TaskStartDate", "set:TaskDueDate", "Save"},
    "move_messages": {"Move"},
    "delete_messages": {"Move"},
    "_compose": {"Items.Add", "set:Subject", "set:HTMLBody", "set:Importance", "Recipients.Add", "set:Type",
                 "ResolveAll", "Attachments.Add", "Send", "Save"},
    # Through the raw IDispatch (PROPERTYPUTREF); the plain assignment is the fallback for test doubles.
    "_set_send_account": {"Invoke", "set:SendUsingAccount"},
    "_after_send": {"SendAndReceive"},
    "send_and_receive": {"SendAndReceive"},
    "send_draft": {"Send"},
    "update_draft": {"set:Subject", "set:HTMLBody", "set:To", "set:CC", "set:BCC", "set:Importance", "ResolveAll",
                     "Attachments.Add", "Attachments.Remove", "Save"},
    "_respond": {"Reply", "ReplyAll", "Forward", "set:HTMLBody", "Recipients.Add", "set:Type", "Attachments.Add",
                 "ResolveAll", "Send", "Save", "Move"},
    "create_folder": {"Folders.Add"},
    "rename_folder": {"set:Name"},
    "move_folder": {"MoveTo"},
    "delete_folder": {"MoveTo"},
    "_save_one": {"SaveAsFile"},
    "_ensure_categories": {"Categories.Add"},
    "write_categories": {"set:Categories"},
    "categorize": {"Save"},
    "create_event": {"Items.Add", "set:Start", "set:End", "set:Sensitivity", "set:MeetingStatus", "Recipients.Add",
                     "set:Type", "ResolveAll", "Save", "Send"},
    "update_event": {"set:Start", "set:End", "set:Sensitivity", "set:RequiredAttendees", "set:OptionalAttendees",
                     "set:MeetingStatus", "ResolveAll", "Save", "Send"},
    "_set_fields": {"set:Subject", "set:Location", "set:Body", "set:AllDayEvent", "set:BusyStatus", "set:ReminderSet",
                    "set:ReminderMinutesBeforeStart"},
    "_apply_recurrence": {"set:RecurrenceType", "set:Interval", "set:DayOfWeekMask", "set:DayOfMonth",
                          "set:MonthOfYear", "set:PatternStartDate", "set:PatternEndDate", "set:Occurrences",
                          "set:NoEndDate"},
    "delete_event": {"Move"},
    # The cancellation is sent on purpose; the item is then filed in Deleted Items with Move, never Delete.
    "cancel_meeting": {"set:MeetingStatus", "set:Body", "Send", "Move"},
    "respond_to_meeting": {"Respond", "set:Body", "Send", "Close", "Save"},
    "create_task": {"Items.Add", "Save"},
    "update_task": {"Save"},
    "_write_task": {"set:Subject", "set:Body", "set:DueDate", "set:StartDate", "set:ReminderSet", "set:ReminderTime",
                    "set:Importance", "set:Complete"},
    "delete_task": {"Move"},
    "create_note": {"Items.Add", "set:Body", "Save"},
    "update_note": {"set:Body", "Save"},
    "delete_note": {"Move"},
    "create_contact": {"Items.Add", "Save"},
    "update_contact": {"Save"},
    "delete_contact": {"Move"},
    "_set": {"setattr", "set:Body", "set:Birthday"},
    "create_contact_group": {"Items.Add", "set:DLName", "AddMember", "Save"},
    "update_contact_group": {"set:DLName", "AddMember", "RemoveMember", "Save"},
    "delete_contact_group": {"Move"},
}

# The only functions that may send: sending is their purpose, and the caller has to have asked for it.
SENDERS = {"_compose", "send_draft", "_respond", "respond_to_meeting", "create_event", "update_event",
           "cancel_meeting"}
# Removing an attachment from a draft is a change to something the caller is editing; nothing else may.
ATTACHMENT_REMOVERS = {"update_draft"}
# Folders are moved, never deleted: MoveTo is the one call that does it.
FOLDER_MOVERS = {"move_folder", "delete_folder"}


def _tail(node: ast.expr) -> str:
    if isinstance(node, ast.Attribute):
        return node.attr
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Call):
        return _tail(node.func)
    return "?"


def ops_of(fn: ast.AST) -> set[str]:
    found = set()
    for node in ast.walk(fn):
        if isinstance(node, ast.Attribute) and isinstance(node.ctx, ast.Load) and node.attr in CHANGERS:
            found.add(f"{_tail(node.value)}.{node.attr}" if node.attr in QUALIFIED else node.attr)
        elif isinstance(node, (ast.Assign, ast.AugAssign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for t in targets:
                if isinstance(t, ast.Attribute) and not (isinstance(t.value, ast.Name) and t.value.id == "self"):
                    found.add(f"set:{t.attr}")
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in ("setattr", "delattr"):
            found.add(node.func.id)
    return found


def methods() -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for path in SOURCES:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                out.setdefault(node.name, set()).update(ops_of(node))
    return out


ALL = methods()


def test_guard_sees_changes():
    src = "def f(o, self):\n o.UnRead = True\n o.Save()\n x = o.Reply\n o.Items.Add(1)\n self.cache = 1\n setattr(o, 'A', 1)"
    assert ops_of(ast.parse(src)) == {"set:UnRead", "Save", "Reply", "Items.Add", "setattr"}


def test_every_declared_write_function_exists():
    assert set(WRITE_ALLOWED) <= set(ALL)


@pytest.mark.parametrize("name", sorted(ALL))
def test_functions_use_only_their_changes(name):
    used = ALL[name] - VIEW
    allowed = WRITE_ALLOWED.get(name, set())
    assert used <= allowed, f"{name} uses {used - allowed}"
    assert not (ALL[name] & NEVER)


def test_nothing_calls_delete():
    """Delete() on an item in Deleted Items is permanent; everything is moved instead."""
    users = {name for name, ops in ALL.items() if "Delete" in ops or "PermanentDelete" in ops}
    assert users == set()


def test_only_send_account_uses_raw_invoke():
    assert {name for name, ops in ALL.items() if "Invoke" in ops} == {"_set_send_account"}


def test_only_sending_functions_send():
    assert {name for name, ops in ALL.items() if "Send" in ops} == SENDERS


def test_attachments_are_removed_only_when_editing_a_draft():
    assert {name for name, ops in ALL.items() if "Attachments.Remove" in ops} == ATTACHMENT_REMOVERS


def test_folders_move_but_are_never_deleted():
    assert {name for name, ops in ALL.items() if "MoveTo" in ops} == FOLDER_MOVERS


def test_every_new_changing_function_is_listed():
    """A changing function nobody put in WRITE_ALLOWED would fail its own test; this names the rule."""
    changers = {name for name, ops in ALL.items() if ops - VIEW}
    assert changers <= set(WRITE_ALLOWED)
