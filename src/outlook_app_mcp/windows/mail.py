"""Windows mail: folders, messages, search, and the changes (mark, move, delete, compose, reply).

Reading uses Outlook's Table API: one filtered, sorted table per folder, read in
bulk, with exact counts. Deleting is always a move to the store's Deleted Items
folder; Outlook's own Delete is never called on mail, because it deletes for good
inside Deleted Items.
"""

from __future__ import annotations

import html as html_lib
import os
import re
import time
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

from ..common import (
    ALREADY_DELETED,
    DEFAULT_LIMIT,
    DEFAULT_MAX_CHARS,
    DEFAULT_SEARCH_DAYS,
    DELETE_NOTE,
    FOLDER_CACHE_SECONDS,
    OutlookError,
    attachment_destination,
    check_addresses,
    check_attachment_files,
    check_folder_name,
    check_ids,
    check_mark_action,
    check_max_chars,
    check_offset,
    check_priority,
    check_recipients,
    check_since_days,
    clamp_limit,
    matches,
    newest_first,
    parse_when,
    pick_attachment,
    to_html,
)
from .com import (
    EXCHANGE_NOT_EXCHANGE,
    EXCHANGE_PUBLIC_FOLDERS,
    FLAG_COMPLETE,
    FLAG_MARKED,
    FLAG_NONE,
    FORMAT_HTML,
    IMPORTANCE,
    IMPORTANCE_CODES,
    ITEM_MAIL,
    PR_ATTACH_MIME,
    PR_CONVERSATION_TOPIC,
    PR_FLAG_STATUS,
    PR_HASATTACH,
    PR_HIDDEN,
    PR_SENDER_SMTP,
    PR_TRANSPORT_HEADERS,
    RECIPIENT_BCC,
    RECIPIENT_CC,
    RECIPIENT_TO,
    ROLE_FOLDERS,
    categories,
    com_errors,
    com_iso,
    dasl_text,
    dasl_utc,
    day_to_com,
    each,
    flag_name,
    make_id,
    mapi,
    prop,
    recipient_address,
    sender_of,
    split_categories,
    split_id,
    store_key,
)

SEND_USING_ACCOUNT_DISPID = 64209
DISPATCH_PROPERTYPUTREF = 8
PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
PR_MESSAGE_CLASS = "http://schemas.microsoft.com/mapi/proptag/0x001A001F"
PR_CONTENT_COUNT = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
PR_CONTENT_UNREAD = "http://schemas.microsoft.com/mapi/proptag/0x36030003"
MSGFLAG_READ = 0x1

# UnRead is a computed property that a Table refuses as a column; the read state
# comes from PR_MESSAGE_FLAGS instead. ReceivedTime by its Outlook name is local time.
TABLE_COLUMNS = ("EntryID", "Subject", "ReceivedTime", PR_MESSAGE_FLAGS, "SenderName", "SenderEmailAddress",
                 PR_SENDER_SMTP, PR_FLAG_STATUS, PR_HASATTACH, "Categories")
MAX_SEARCH_ROWS = 20000  # per folder; a search that matches more should be narrowed
MAX_RECIPIENT_SCAN = 1000  # messages opened to check To/Cc addresses (the table has no address column)
MAX_THREAD_ROWS = 500  # messages of one subject looked at when finding a conversation
SIGNATURES_ENV = "OUTLOOK_APP_MCP_SIGNATURES_DIR"  # tests (and unusual setups) point this at another folder

# Message classes the mail tools work on. Appointments, contacts and tasks can sit in
# mail folders too (Deleted Items above all); the mail tools leave them alone.
MAIL_CLASSES = ("IPM.Note", "IPM.Schedule", "IPM.Post", "IPM.Sharing", "IPM.Outlook.Recall", "REPORT.")
MAIL_ONLY = "(" + " OR ".join(f'"{PR_MESSAGE_CLASS}" LIKE \'{c}%\'' for c in MAIL_CLASSES) + ")"


def is_mail_class(message_class: str) -> bool:
    return (message_class or "").startswith(MAIL_CLASSES)


@dataclass
class Store:
    key: str
    store_id: str
    display_name: str
    account: str  # SMTP address of the account that delivers here, else the store's name
    kind: str  # "exchange" or "other" (PST, IMAP, ...)
    roles: dict[str, str] = field(default_factory=dict)  # role -> folder EntryID
    has_account: bool = False  # an Outlook account delivers here (not an archive or other data file)
    public: bool = False  # Exchange Public Folders: not walked (can hold thousands of online folders)


@dataclass
class Folder:
    id: str
    account: str
    account_name: str
    path: str
    unread: int
    total: int
    role: str
    kind: str
    store: str  # store key
    entry_id: str

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "account": self.account,
            "account_name": self.account_name,
            "path": self.path,
            "unread": self.unread,
            "total": self.total,
            "role": self.role or None,
            "kind": self.kind,
        }


def table_rows(table, columns: Sequence[str], limit: int | None, skip: int = 0) -> list[dict]:
    """Rows of an Outlook Table as dicts, in table order, at most `limit`, after skipping `skip` rows."""
    while skip > 0 and not table.EndOfTable:
        skipped = table.GetArray(min(500, skip))
        if not skipped:
            break
        skip -= len(skipped)
    out: list[dict] = []
    while not table.EndOfTable and (limit is None or len(out) < limit):
        want = 500 if limit is None else min(500, limit - len(out))
        block = table.GetArray(want)
        if not block:
            break
        for row in block:
            out.append(dict(zip(columns, row)))
    return out


def summary(row: dict, folder_id_of) -> dict:
    sender = row.get("SenderEmailAddress") or ""
    if "@" not in sender or sender.startswith("/"):
        sender = row.get(PR_SENDER_SMTP) or sender
    return {
        "id": folder_id_of(row["EntryID"]),
        "received": com_iso(row.get("ReceivedTime")),
        "is_read": bool((row.get(PR_MESSAGE_FLAGS) or 0) & MSGFLAG_READ),
        "flag": flag_name(row.get(PR_FLAG_STATUS)),
        "subject": row.get("Subject") or "",
        "from": {"name": row.get("SenderName") or "", "address": sender or ""},
    }


def prepend_html(existing: str, block: str) -> str:
    """Put block at the top of an HTML body (after <body ...>), above the quoted original."""
    m = re.search(r"<body[^>]*>", existing or "", flags=re.IGNORECASE)
    if not m:
        return block + (existing or "")
    return existing[: m.end()] + block + existing[m.end():]


class WinMailMixin:
    """Needs self._app (Outlook.Application) and self._com_error."""

    _stores: dict[str, Store] | None
    _folders: list[Folder] | None
    _folders_at: float
    _accounts: dict[str, object]  # SMTP address (lower case) -> Account object

    # -- stores and folders --------------------------------------------------

    @property
    def _ns(self):
        return self._app.Session

    def _invalidate(self) -> None:
        self._folders = None

    def _load_stores(self) -> dict[str, Store]:
        accounts: dict[str, tuple[str, object]] = {}  # StoreID -> (smtp, account)
        self._accounts = {}
        for acc in each(self._ns.Accounts):
            smtp = prop(acc, "SmtpAddress", "") or ""
            if smtp:
                self._accounts[smtp.lower()] = acc
            try:
                sid = acc.DeliveryStore.StoreID
            except Exception:  # noqa: BLE001, S112 - accounts without a delivery store
                continue
            accounts.setdefault(sid, (smtp, acc))
        stores: dict[str, Store] = {}
        for st in each(self._ns.Stores):
            sid = st.StoreID
            name = prop(st, "DisplayName", "") or ""
            smtp = accounts.get(sid, ("", None))[0]
            store_type = prop(st, "ExchangeStoreType", EXCHANGE_NOT_EXCHANGE)
            kind = "other" if store_type == EXCHANGE_NOT_EXCHANGE else "exchange"
            roles = {}
            for role, const in ROLE_FOLDERS.items():
                try:
                    roles[role] = st.GetDefaultFolder(const).EntryID
                except Exception:  # noqa: BLE001, S110 - this store has no such folder
                    pass
            key = store_key(sid)
            stores[key] = Store(key, sid, name, smtp or name, kind, roles, has_account=bool(smtp),
                                public=store_type == EXCHANGE_PUBLIC_FOLDERS)
        return stores

    def _open_stores(self):
        """(COM store, Store) for every store worth walking: not Public Folders."""
        if self._stores is None:
            self._stores = self._load_stores()
        for st in each(self._ns.Stores):
            store = self._store(store_key(st.StoreID))
            if not store.public:
                yield st, store

    def _store(self, key: str) -> Store:
        if self._stores is None or key not in self._stores:
            self._stores = self._load_stores()
        if key not in self._stores:
            raise OutlookError("That id belongs to a mailbox or data file that is not open in Outlook.")
        return self._stores[key]

    def _load_folders(self) -> list[Folder]:
        self._stores = self._load_stores()
        out: list[Folder] = []
        for st, store in self._open_stores():
            role_of = {eid: role for role, eid in store.roles.items()}

            def walk(parent, prefix: str, store=store, role_of=role_of) -> None:
                for f in each(parent.Folders):
                    name = prop(f, "Name", "") or ""
                    path = f"{prefix}/{name}" if prefix else name
                    if prop(f, "DefaultItemType", None) == ITEM_MAIL and not mapi(f, PR_HIDDEN, False):
                        eid = f.EntryID
                        out.append(Folder(
                            id=make_id(store.key, eid),
                            account=store.account,
                            account_name=store.display_name,
                            path=path,
                            unread=int(mapi(f, PR_CONTENT_UNREAD) or prop(f, "UnReadItemCount", 0) or 0),
                            total=int(mapi(f, PR_CONTENT_COUNT) or f.Items.Count),
                            role=role_of.get(eid, ""),
                            kind=store.kind,
                            store=store.key,
                            entry_id=eid,
                        ))
                    walk(f, path)

            walk(st.GetRootFolder(), "")
        return out

    @com_errors
    def list_folders(self, account: str | None = None, refresh: bool = False) -> list[dict]:
        return [f.as_dict() for f in self.folders(account, refresh=refresh)]

    def folders(self, account: str | None = None, refresh: bool = False) -> list[Folder]:
        fresh = time.monotonic() - self._folders_at < FOLDER_CACHE_SECONDS
        if refresh or self._folders is None or not fresh:
            folders = self._load_folders()
            if not folders:
                raise OutlookError("Outlook shows no mail folders. Add an account in classic Outlook first.")
            self._folders = folders
            self._folders_at = time.monotonic()
        result = self._folders
        if account:
            want = account.casefold()
            result = [f for f in result if want in f.account.casefold() or want in f.account_name.casefold()]
            if not result:
                accounts = sorted({f.account for f in self._folders})
                raise OutlookError(f"No account matches {account!r}. Accounts: {', '.join(accounts)}")
        return result

    def require_accounts(self) -> None:
        self.folders()

    def folder(self, folder_id: str) -> Folder:
        split_id(folder_id, "folder_id")
        for refresh in (False, True):
            for f in self.folders(refresh=refresh):
                if f.id == folder_id.strip():
                    return f
        raise OutlookError(f"No mail folder with id {folder_id}. Use list_folders to see the ids.")

    def _folder_obj(self, folder: Folder):
        return self._ns.GetFolderFromID(folder.entry_id, self._store(folder.store).store_id)

    def _role_folder(self, role: str, account: str | None) -> Folder:
        found = [f for f in self.folders(account) if f.role == role]
        if not found:
            raise OutlookError(f"Could not find the {role} folder. Pass folder_id (see list_folders).")
        if len(found) > 1 and not account:
            # Archives and other data files often carry their own Inbox; prefer real
            # accounts, then Outlook's default store.
            with_account = [f for f in found if self._store(f.store).has_account]
            found = with_account or found
            if len(found) > 1:
                try:
                    default = store_key(self._ns.DefaultStore.StoreID)
                except Exception:  # noqa: BLE001 - no default store
                    default = ""
                found = [f for f in found if f.store == default] or found
        if len(found) > 1:
            names = ", ".join(f"{f.account} (id {f.id})" for f in found)
            raise OutlookError(f"More than one account has the {role} folder: {names}. Pass account or folder_id.")
        return found[0]

    def _role_folder_of(self, folder: Folder, role: str) -> Folder:
        for f in self.folders():
            if f.store == folder.store and f.role == role:
                return f
        raise OutlookError(f"{folder.account} has no {role} folder.")

    # -- items ---------------------------------------------------------------

    def _item(self, item_id: str, what: str = "message_id"):
        key, entry = split_id(item_id, what)
        return self._ns.GetItemFromID(entry, self._store(key).store_id), key

    def _mail_item(self, item_id: str, what: str = "message_id"):
        """An item the mail tools may act on: a message, not an appointment or contact."""
        item, key = self._item(item_id, what)
        if not is_mail_class(prop(item, "MessageClass", "")):
            raise OutlookError(f"{item_id} is not a mail message ({prop(item, 'MessageClass', '?')}).")
        return item, key

    def _under_deleted(self, item, key: str) -> bool:
        """True when the item is in its store's Deleted Items or any folder below it, or
        when that can't be ruled out (the store has no Deleted Items folder). Outlook
        deletes items in Deleted Items for good."""
        deleted = self._store(key).roles.get("deleted")
        if not deleted:
            return True
        node = prop(item, "Parent")
        for _ in range(64):
            eid = prop(node, "EntryID", None) if node is not None else None
            if not eid:
                return False
            if eid == deleted:
                return True
            node = prop(node, "Parent")
        return True

    def _deleted_folder_obj(self, key: str):
        store = self._store(key)
        if not store.roles.get("deleted"):
            raise OutlookError(f"{store.account} has no Deleted Items folder; nothing was deleted.")
        return self._ns.GetFolderFromID(store.roles["deleted"], store.store_id)

    def _folder_of_item(self, item, key: str) -> Folder:
        parent_eid = item.Parent.EntryID
        fid = make_id(key, parent_eid)
        for refresh in (False, True):
            for f in self.folders(refresh=refresh):
                if f.id == fid:
                    return f
        raise OutlookError("That item is not in a mail folder.")

    # -- reading -------------------------------------------------------------

    def _table(self, folder: Folder, since: datetime | None, unread_only: bool, extra: str = "",
               until: datetime | None = None):
        parts = []
        if since is not None:
            parts.append(f'"urn:schemas:httpmail:datereceived" > {dasl_text(dasl_utc(since))}')
        if until is not None:
            parts.append(f'"urn:schemas:httpmail:datereceived" < {dasl_text(dasl_utc(until))}')
        if unread_only:
            parts.append('"urn:schemas:httpmail:read" = 0')
        if extra:
            parts.append(extra)
        parts.append(MAIL_ONLY)
        table = self._folder_obj(folder).GetTable("@SQL=" + " AND ".join(parts), 0)
        cols = table.Columns
        cols.RemoveAll()
        for c in TABLE_COLUMNS:
            cols.Add(c)
        table.Sort("ReceivedTime", True)
        return table

    def _summaries(self, folder: Folder, rows: list[dict]) -> list[dict]:
        out = []
        for r in rows:
            m = summary(r, lambda eid: make_id(folder.store, eid))
            m["folder_id"] = folder.id
            m["folder"] = folder.path
            m["account"] = folder.account
            out.append(m)
        return out

    @staticmethod
    def _since(days: float | None) -> datetime | None:
        if check_since_days(days) is None:
            return None
        return datetime.now().astimezone() - timedelta(days=days)

    @com_errors
    def list_messages(self, folder_id=None, account=None, limit=DEFAULT_LIMIT, unread_only=False, since_days=None,
                      offset=0) -> dict:
        limit = clamp_limit(limit)
        offset = check_offset(offset)
        folder = self.folder(folder_id) if folder_id is not None else self._role_folder("inbox", account)
        table = self._table(folder, self._since(since_days), unread_only)
        matched = int(table.GetRowCount())
        msgs = self._summaries(folder, table_rows(table, TABLE_COLUMNS, limit, skip=offset))
        return {
            "folder": folder.as_dict(),
            "matched": matched,
            "offset": offset,
            "returned": len(msgs),
            "has_more": matched > offset + len(msgs),
            "messages": msgs,
        }

    @staticmethod
    def _row_passes(row: dict, sender: str, has_attachments, flagged, category: str) -> bool:
        """The filters a Table row can answer without opening the message."""
        if sender:
            hay = (row.get("SenderName"), row.get("SenderEmailAddress"), row.get(PR_SENDER_SMTP))
            if not any(sender in (h or "").casefold() for h in hay):
                return False
        if has_attachments is not None and bool(row.get(PR_HASATTACH)) != has_attachments:
            return False
        if flagged is not None and (flag_name(row.get(PR_FLAG_STATUS)) == "flagged") != flagged:
            return False
        mine = {c.casefold() for c in split_categories(row.get("Categories"))}
        return not category or category in mine

    @com_errors
    def search(self, query="", folder_id=None, account=None, since_days=DEFAULT_SEARCH_DAYS, include_body=False,
               limit=DEFAULT_LIMIT, from_address=None, to_address=None, has_attachments=None, flagged=None,
               category=None, after=None, before=None, offset=0) -> dict:
        """Search by text and/or filters. Text matches subject and sender (and the body with include_body).
        from_address, has_attachments, flagged, category, after and before are answered by the folder tables;
        to_address needs each candidate message opened (the table has no address column), so at most
        MAX_RECIPIENT_SCAN of the newest candidates are checked and the result says so when that cut in."""
        query = (query or "").strip()
        sender = (from_address or "").strip().casefold()
        to_needle = (to_address or "").strip().casefold()
        wanted_category = (category or "").strip().casefold()
        filtered = any([sender, to_needle, wanted_category, has_attachments is not None, flagged is not None,
                        after, before])
        if not query and not filtered:
            raise OutlookError("Give a query or at least one filter (from_address, to_address, has_attachments, "
                               "flagged, category, after, before).")
        limit = clamp_limit(limit)
        offset = check_offset(offset)
        after_dt = parse_when(after, "after") if after else None
        before_dt = parse_when(before, "before") if before else None
        if after_dt and before_dt and before_dt <= after_dt:
            raise OutlookError("before must be later than after.")
        since = after_dt or self._since(since_days)
        if folder_id is not None:
            scope = [self.folder(folder_id)]
        else:
            # Without an account, search the accounts' own mailboxes, not archives and other data files.
            scope = [f for f in self.folders(account) if f.total > 0 and (account or self._store(f.store).has_account)]
        needle = query.casefold()
        found: dict[str, dict] = {}
        for folder in scope:
            rows = table_rows(self._table(folder, since, False, until=before_dt), TABLE_COLUMNS, MAX_SEARCH_ROWS)
            for row, m in zip(rows, self._summaries(folder, rows), strict=True):
                if (not needle or matches(m, needle)) and self._row_passes(row, sender, has_attachments, flagged,
                                                                          wanted_category):
                    found[m["id"]] = m
            if include_body and query:
                # % is LIKE's wildcard; it is dropped from the text so it matches literally elsewhere.
                body = f'"urn:schemas:httpmail:textdescription" LIKE {dasl_text("%" + query.replace("%", "") + "%")}'
                rows = table_rows(self._table(folder, since, False, body, until=before_dt), TABLE_COLUMNS,
                                  MAX_SEARCH_ROWS)
                for row, m in zip(rows, self._summaries(folder, rows), strict=True):
                    if self._row_passes(row, sender, has_attachments, flagged, wanted_category):
                        found.setdefault(m["id"], m)
        msgs = newest_first(list(found.values()))
        capped = False
        if to_needle:
            kept = []
            for opened, m in enumerate(msgs):
                if opened >= MAX_RECIPIENT_SCAN:
                    capped = True
                    break
                item = self._item(m["id"])[0]
                recips = [r for r in each(item.Recipients) if prop(r, "Type", RECIPIENT_TO) in (RECIPIENT_TO, RECIPIENT_CC)]
                if any(to_needle in (recipient_address(r) or "").casefold() or to_needle in (prop(r, "Name", "") or "").casefold()
                       for r in recips):
                    kept.append(m)
            msgs = kept
        page = msgs[offset:offset + limit]
        result = {
            "query": query,
            "searched_folders": len(scope),
            "since_days": None if after_dt else since_days,
            "include_body": include_body,
            "matched": len(msgs),
            "offset": offset,
            "returned": len(page),
            "has_more": len(msgs) > offset + len(page),
            "messages": page,
        }
        if capped:
            result["note"] = (f"to_address was checked on the newest {MAX_RECIPIENT_SCAN} candidates only; narrow the "
                              "search (dates, folder, from_address) to be sure of older ones.")
        return result

    @com_errors
    def read_message(self, message_id: str, max_chars: int = DEFAULT_MAX_CHARS, include_headers: bool = False,
                     as_html: bool = False) -> dict:
        check_max_chars(max_chars)
        item, key = self._mail_item(message_id)
        folder = self._folder_of_item(item, key)
        to, cc, bcc = [], [], []
        for r in each(item.Recipients):
            entry = {"name": prop(r, "Name", "") or "", "address": recipient_address(r)}
            {RECIPIENT_TO: to, RECIPIENT_CC: cc, RECIPIENT_BCC: bcc}.get(prop(r, "Type", RECIPIENT_TO), to).append(entry)
        atts = []
        for i, a in enumerate(each(item.Attachments), 1):
            size = prop(a, "Size", None)
            atts.append({
                "index": i,
                "name": prop(a, "FileName", "") or prop(a, "DisplayName", "") or "",
                "size": int(size) if isinstance(size, int) and size >= 0 else None,
                "content_type": mapi(a, PR_ATTACH_MIME) or None,
            })
        # Read directly: a protected or not-downloaded body must fail, not look empty.
        body = (item.HTMLBody if as_html else item.Body) or ""
        status = prop(item, "FlagStatus", None)
        if status is None:
            status = mapi(item, PR_FLAG_STATUS)
        message_class = prop(item, "MessageClass", "") or ""
        result = {
            "id": make_id(key, item.EntryID),
            "folder_id": folder.id,
            "folder": folder.path,
            "account": folder.account,
            "received": com_iso(prop(item, "ReceivedTime")),
            "sent": com_iso(prop(item, "SentOn")),
            "is_read": not bool(prop(item, "UnRead", False)),
            "has_html": prop(item, "BodyFormat", None) == FORMAT_HTML,
            "subject": prop(item, "Subject", "") or "",
            "from": sender_of(item),
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "attachments": atts,
            "flag": flag_name(status),
            "categories": categories(item),
            "importance": IMPORTANCE.get(prop(item, "Importance", 1), "normal"),
            "conversation_id": prop(item, "ConversationID", None) or None,
            "is_meeting": message_class.startswith("IPM.Schedule.Meeting"),
            "meeting_event_id": self._meeting_event_id(item) if message_class.startswith("IPM.Schedule.Meeting") else None,
            "body_format": "html" if as_html else "text",
            "body": body[:max_chars],
            "body_truncated": len(body) > max_chars,
            "body_length": len(body),
        }
        if include_headers:
            result["headers"] = mapi(item, PR_TRANSPORT_HEADERS, "") or ""
        return result

    def _meeting_event_id(self, request) -> str | None:
        """The calendar event a meeting request, update or cancellation belongs to, or None when
        Outlook has none for it (not tracked yet, or already removed)."""
        try:
            appt = request.GetAssociatedAppointment(False)
            return make_id(self._key_of(appt), appt.EntryID) if appt is not None else None
        except Exception:  # noqa: BLE001 - no associated appointment: nothing to point at
            return None

    # -- changing ------------------------------------------------------------

    @com_errors
    def mark_messages(self, message_ids: Sequence[str], action: str, due_date: str | None = None) -> dict:
        check_mark_action(action)
        due = None
        if due_date:
            if action != "flag":
                raise OutlookError("due_date only goes with action='flag'.")
            due = day_to_com(parse_when(due_date, "due_date"))
        items = [self._mail_item(i)[0] for i in check_ids(message_ids)]
        for item in items:
            if action in ("read", "unread"):
                item.UnRead = action == "unread"
            else:
                item.FlagStatus = {"flag": FLAG_MARKED, "unflag": FLAG_NONE, "complete": FLAG_COMPLETE}[action]
                if due is not None:
                    item.TaskStartDate = due
                    item.TaskDueDate = due
            item.Save()
        self._invalidate()
        result = {"action": action, "changed": len(items)}
        if due is not None:
            result["due_date"] = due_date
        return result

    @com_errors
    def move_messages(self, message_ids: Sequence[str], folder_id: str) -> dict:
        ids = check_ids(message_ids)
        dest = self.folder(folder_id)
        items = [self._mail_item(i)[0] for i in ids]
        target = self._folder_obj(dest)
        new_ids: dict[str, str | None] = {}
        for old, item in zip(ids, items, strict=True):
            moved = item.Move(target)
            eid = prop(moved, "EntryID", "") if moved is not None else ""
            new_ids[old] = make_id(dest.store, eid) if eid else None
        self._invalidate()
        result = {"moved": len(items), "to": dest.as_dict(), "new_ids": new_ids}
        if None in new_ids.values():
            result["note"] = "Outlook did not report some new ids (null above); find them with list_messages."
        return result

    @com_errors
    def delete_messages(self, message_ids: Sequence[str]) -> dict:
        """Move messages to their store's Deleted Items. Never permanent: messages
        already in Deleted Items are refused, and Outlook's Delete is never called."""
        ids = check_ids(message_ids)
        located = []
        for i in ids:
            item, key = self._mail_item(i)
            located.append((i, item, key))
        already = [i for i, item, key in located if self._under_deleted(item, key)]
        if already:
            raise OutlookError(f"Message {', '.join(already)} {ALREADY_DELETED}")
        for _, item, key in located:
            item.Move(self._deleted_folder_obj(key))
        self._invalidate()
        return {"deleted": len(located), "note": DELETE_NOTE}

    def _send_account(self, account: str | None):
        """(Account object or None for Outlook's default, its address, its Drafts folder)."""
        if not account:
            return None, "", None
        matched = self.folders(account)
        emails = sorted({f.account for f in matched})
        if len(emails) > 1:
            raise OutlookError(f"{account!r} matches several accounts: {', '.join(emails)}.")
        acc = self._accounts.get(emails[0].lower())
        if acc is None:
            raise OutlookError(f"{emails[0]} is a data file, not an account Outlook can send from.")
        drafts = next((f for f in matched if f.role == "drafts"), None)
        return acc, emails[0], drafts

    @com_errors
    def _compose(self, mode, to, subject, body, cc, bcc, attachments, body_is_html, account, importance=None,
                 signature=None) -> dict:
        to_l, cc_l, bcc_l = check_recipients(to, cc, bcc, mode == "send")
        files = check_attachment_files(attachments)
        level = check_priority(importance)
        html_body = (body if body_is_html else to_html(body)) + self._signature_block(signature)
        acc, address, drafts = self._send_account(account)
        if drafts is not None:
            msg = self._folder_obj(drafts).Items.Add(ITEM_MAIL)
        else:
            msg = self._app.CreateItem(ITEM_MAIL)
        if acc is not None:
            self._set_send_account(msg, acc, address)
        msg.Subject = subject or ""
        msg.HTMLBody = html_body
        if level is not None:
            msg.Importance = IMPORTANCE_CODES[level]
        for addr_list, kind in ((to_l, RECIPIENT_TO), (cc_l, RECIPIENT_CC), (bcc_l, RECIPIENT_BCC)):
            for a in addr_list:
                msg.Recipients.Add(a).Type = kind
        resolved = msg.Recipients.ResolveAll() if (to_l or cc_l or bcc_l) else True
        for f in files:
            msg.Attachments.Add(f)
        self._invalidate()
        if mode == "send":
            self._check_resolved(msg, resolved)
            msg.Send()
            return {"id": None, "status": "sent", "account": address or "(default)", **self._after_send()}
        msg.Save()
        key = self._key_of(msg)
        return {"id": make_id(key, msg.EntryID), "status": "draft", "account": address or "(default)"}

    @staticmethod
    def _set_send_account(msg, acc, address: str) -> None:
        """Set the From account and prove it took. Through pywin32's late binding a plain
        assignment is a property put, which SendUsingAccount (an object reference)
        ignores or rejects; it needs DISPATCH_PROPERTYPUTREF on dispid 0xFAD1."""
        ole = getattr(msg, "_oleobj_", None)
        if ole is not None:
            ole.Invoke(SEND_USING_ACCOUNT_DISPID, 0, DISPATCH_PROPERTYPUTREF, 0, acc)
        else:
            msg.SendUsingAccount = acc
        actual = prop(prop(msg, "SendUsingAccount"), "SmtpAddress", "") or ""
        if actual.lower() != address.lower():
            raise OutlookError(
                f"Outlook did not accept {address} as the sending account (it reports {actual or 'none'}); "
                "nothing was saved or sent."
            )

    @staticmethod
    def _check_resolved(msg, resolved) -> None:
        if resolved:
            return
        bad = [prop(r, "Address", "") or prop(r, "Name", "") for r in each(msg.Recipients) if not prop(r, "Resolved", True)]
        raise OutlookError(f"Outlook could not resolve {', '.join(bad) or 'a recipient'}; nothing was sent.")

    def _after_send(self) -> dict:
        """When Outlook runs without a window, sent mail can wait in the Outbox. Start a
        send/receive and say so."""
        if self.outlook_window_open:
            return {}
        self._ns.SendAndReceive(False)
        return {"note": "Outlook is running without a window, so the message was queued in the Outbox and a "
                        "send/receive was started. Open Outlook to make sure it leaves the Outbox."}

    def _key_of(self, item) -> str:
        return store_key(item.Parent.Store.StoreID)

    def create_draft(self, to=(), subject="", body="", cc=(), bcc=(), attachments=(), body_is_html=False,
                     account=None, importance=None, signature=None) -> dict:
        return self._compose("draft", to, subject, body, cc, bcc, attachments, body_is_html, account, importance,
                             signature)

    def send_message(self, to, subject, body, cc=(), bcc=(), attachments=(), body_is_html=False, account=None,
                     importance=None, signature=None) -> dict:
        return self._compose("send", to, subject, body, cc, bcc, attachments, body_is_html, account, importance,
                             signature)

    # -- signatures ----------------------------------------------------------

    @staticmethod
    def _signatures_dir() -> Path:
        override = os.environ.get(SIGNATURES_ENV, "").strip()
        if override:
            return Path(override)
        appdata = os.environ.get("APPDATA", "").strip()
        if not appdata:
            raise OutlookError("Cannot find Outlook's Signatures folder: the APPDATA variable is not set.")
        return Path(appdata) / "Microsoft" / "Signatures"

    def _signature_files(self) -> dict[str, Path]:
        folder = self._signatures_dir()
        if not folder.is_dir():
            return {}
        return {p.stem: p for p in sorted(folder.glob("*.htm"), key=lambda p: p.stem.casefold())}

    @staticmethod
    def _read_signature_file(path: Path) -> str:
        """Outlook writes signatures as UTF-16 (with BOM), UTF-8 or the Windows code page."""
        raw = path.read_bytes()
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            return raw.decode("utf-16", errors="replace")
        try:
            return raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            return raw.decode("cp1252", errors="replace")

    def _signature_block(self, name: str | None) -> str:
        """HTML to append below a message body for the signature `name`, or ''."""
        if not name:
            return ""
        files = self._signature_files()
        match = next((p for n, p in files.items() if n.casefold() == name.strip().casefold()), None)
        if match is None:
            raise OutlookError(f"No signature named {name!r}. Signatures: {', '.join(files) or 'none'}.")
        text = self._read_signature_file(match)
        inner = re.search(r"<body[^>]*>(.*)</body>", text, flags=re.IGNORECASE | re.DOTALL)
        return "<div><br></div><div>" + (inner.group(1) if inner else text) + "</div>"

    @com_errors
    def list_signatures(self) -> dict:
        out = []
        for name, path in self._signature_files().items():
            twin = path.with_suffix(".txt")
            if twin.is_file():
                text = self._read_signature_file(twin)
            else:
                text = self._read_signature_file(path)
                text = re.sub(r"(?is)<(script|style|head)\b.*?</\1>", " ", text)
                text = html_lib.unescape(re.sub(r"<[^>]+>", " ", text))
            out.append({"name": name, "preview": " ".join(text.split())[:200]})
        return {"signatures": out}

    @com_errors
    def send_draft(self, draft_id: str) -> dict:
        msg = self.read_message(draft_id, max_chars=1)
        folder = self.folder(msg["folder_id"])
        if folder.role != "drafts":
            raise OutlookError(f"Message {draft_id} is not in a Drafts folder; only drafts can be sent.")
        if not (msg["to"] or msg["cc"] or msg["bcc"]):
            raise OutlookError(f"Draft {draft_id} has no recipients.")
        item, _ = self._mail_item(draft_id)
        item.Send()
        self._invalidate()
        return {"id": draft_id, "status": "sent", "subject": msg["subject"], **self._after_send()}

    @com_errors
    def _respond(self, message_id, mode, body, to, send, cc=(), attachments=(), signature=None) -> dict:
        to_l = check_addresses(to, "to")
        cc_l = check_addresses(cc, "cc")
        files = check_attachment_files(attachments)
        if mode == "forward" and send and not to_l:
            raise OutlookError("A forward to send needs at least one address in 'to'.")
        block = (to_html(body) if body else "") + self._signature_block(signature)
        item, key = self._mail_item(message_id)
        original = self._folder_of_item(item, key)
        drafts = self._role_folder_of(original, "drafts")
        r = {"reply": item.Reply, "replyall": item.ReplyAll, "forward": item.Forward}[mode]()
        if block:
            r.HTMLBody = prepend_html(r.HTMLBody, block)
        for addr_list, kind in ((to_l, RECIPIENT_TO), (cc_l, RECIPIENT_CC)):
            for a in addr_list:
                r.Recipients.Add(a).Type = kind
        for f in files:
            r.Attachments.Add(f)
        resolved = r.Recipients.ResolveAll() if (to_l or cc_l) else True
        self._invalidate()
        if send:
            self._check_resolved(r, resolved)
            r.Send()
            return {"id": None, "status": "sent", "in_reply_to": message_id, "mode": mode, **self._after_send()}
        r.Save()
        if r.Parent.EntryID != drafts.entry_id or self._key_of(r) != drafts.store:
            r = r.Move(self._folder_obj(drafts))
        return {"id": make_id(drafts.store, r.EntryID), "status": "draft", "in_reply_to": message_id, "mode": mode}

    def reply_to_message(self, message_id: str, body: str, reply_all: bool = False, send: bool = False, cc=(),
                         attachments=(), signature=None) -> dict:
        return self._respond(message_id, "replyall" if reply_all else "reply", body, (), send, cc, attachments,
                             signature)

    def forward_message(self, message_id: str, to: Sequence[str], body: str = "", send: bool = False, cc=(),
                        attachments=(), signature=None) -> dict:
        return self._respond(message_id, "forward", body, to, send, cc, attachments, signature)

    @com_errors
    def create_folder(self, name: str, parent_id: str | None = None, account: str | None = None) -> dict:
        name = check_folder_name(name)
        if parent_id is not None:
            parent = self._folder_obj(self.folder(parent_id))
            key = self.folder(parent_id).store
        else:
            inbox = self._role_folder("inbox", account)
            key = inbox.store
            parent = self._folder_obj(inbox).Store.GetRootFolder()
        new = parent.Folders.Add(name)
        self._invalidate()
        return {"id": make_id(key, new.EntryID), "name": name}

    @com_errors
    def rename_folder(self, folder_id: str, name: str) -> dict:
        name = check_folder_name(name)
        f = self.folder(folder_id)
        if f.role:
            raise OutlookError(f"{f.path} is the account's {f.role} folder; it can't be renamed.")
        self._folder_obj(f).Name = name
        self._invalidate()
        return {"id": f.id, "old_path": f.path, "name": name}

    @com_errors
    def update_draft(self, draft_id: str, to=None, cc=None, bcc=None, subject=None, body=None,
                     body_is_html: bool = False, add_attachments=(), remove_attachments=(), importance=None) -> dict:
        """Change a draft that is still unsent. A recipient list that is given replaces the old one: it is
        written as the address string (To, CC, BCC), which makes Outlook rebuild just that kind of recipient,
        so no recipient is ever deleted one by one. Attachments are removed by name or 1-based index."""
        lists = {name: None if value is None else check_addresses(value, name)
                 for name, value in (("to", to), ("cc", cc), ("bcc", bcc))}
        files = check_attachment_files(add_attachments)
        level = check_priority(importance)
        item, key = self._mail_item(draft_id, "draft_id")
        folder = self._folder_of_item(item, key)
        if folder.role != "drafts" or prop(item, "Sent", False):
            raise OutlookError(f"Message {draft_id} is not an unsent draft in a Drafts folder; only those can be changed.")
        current = [
            {"index": i, "name": prop(a, "FileName", "") or prop(a, "DisplayName", "") or ""}
            for i, a in enumerate(each(item.Attachments), 1)
        ]
        drop = sorted({pick_attachment(current, a, draft_id)["index"] for a in remove_attachments or []},
                      reverse=True)
        if all(v is None for v in lists.values()) and subject is None and body is None and level is None \
                and not files and not drop:
            raise OutlookError("Nothing to change.")
        if subject is not None:
            item.Subject = subject
        if body is not None:
            item.HTMLBody = body if body_is_html else to_html(body)
        if lists["to"] is not None:
            item.To = "; ".join(lists["to"])
        if lists["cc"] is not None:
            item.CC = "; ".join(lists["cc"])
        if lists["bcc"] is not None:
            item.BCC = "; ".join(lists["bcc"])
        if any(v is not None for v in lists.values()) and item.Recipients.Count:
            item.Recipients.ResolveAll()
        if level is not None:
            item.Importance = IMPORTANCE_CODES[level]
        for index in drop:  # highest first, so the numbers of the others stay right
            item.Attachments.Remove(index)
        for f in files:
            item.Attachments.Add(f)
        item.Save()
        self._invalidate()
        return {"id": draft_id, "status": "draft"}

    @com_errors
    def get_conversation(self, message_id: str, limit: int = DEFAULT_LIMIT) -> dict:
        """The messages of one conversation in every mail folder of the message's account, oldest first.

        Candidates are the messages with the same conversation topic (a cheap filter every store supports);
        each is opened and kept only if its ConversationID matches, so unrelated threads with the same subject
        drop out. When there are more than `limit`, the newest `limit` are returned."""
        limit = clamp_limit(limit)
        item, key = self._mail_item(message_id)
        cid = prop(item, "ConversationID", None)
        if not cid:
            raise OutlookError("Outlook gave this message no conversation id.")
        topic = prop(item, "ConversationTopic", "") or ""
        found = {make_id(key, item.EntryID): None}
        if topic:
            store_id = self._store(key).store_id
            clause = f'"{PR_CONVERSATION_TOPIC}" = {dasl_text(topic)}'
            for folder in [f for f in self.folders() if f.store == key]:
                for row in table_rows(self._table(folder, None, False, clause), TABLE_COLUMNS, MAX_THREAD_ROWS):
                    mid = make_id(key, row["EntryID"])
                    if mid in found:
                        continue
                    other = self._ns.GetItemFromID(row["EntryID"], store_id)
                    if prop(other, "ConversationID", None) == cid:
                        found[mid] = (folder, row)
        msgs = []
        for mid, hit in found.items():
            if hit is None:
                folder = self._folder_of_item(item, key)
                row = {
                    "EntryID": item.EntryID, "Subject": prop(item, "Subject", ""),
                    "ReceivedTime": prop(item, "ReceivedTime"), PR_MESSAGE_FLAGS: 0 if prop(item, "UnRead", False) else 1,
                    "SenderName": prop(item, "SenderName", ""), "SenderEmailAddress": prop(item, "SenderEmailAddress", ""),
                    PR_SENDER_SMTP: mapi(item, PR_SENDER_SMTP), PR_FLAG_STATUS: prop(item, "FlagStatus", 0),
                }
            else:
                folder, row = hit
            msgs.extend(self._summaries(folder, [row]))
        msgs.sort(key=lambda m: m["received"] or "")
        shown = msgs[-limit:]
        return {"conversation_id": cid, "count": len(msgs), "returned": len(shown), "has_more": len(msgs) > len(shown),
                "messages": shown}

    # -- folders: move and delete ----------------------------------------------

    @com_errors
    def move_folder(self, folder_id: str, parent_id: str) -> dict:
        """Move a mail folder below another folder of the same account. Special folders (Inbox, Sent ...) stay
        where they are, and nothing is moved into Deleted Items this way (that is delete_folder)."""
        f, parent = self.folder(folder_id), self.folder(parent_id)
        if f.role:
            raise OutlookError(f"{f.path} is the account's {f.role} folder; it can't be moved.")
        if f.store != parent.store:
            raise OutlookError("A folder can only be moved inside its own account; "
                               f"{parent.path} belongs to {parent.account}, not {f.account}.")
        if f.id == parent.id or parent.path.startswith(f.path + "/"):
            raise OutlookError("A folder can't be moved into itself or one of its own subfolders.")
        obj, target = self._folder_obj(f), self._folder_obj(parent)
        if self._under_deleted(obj, f.store) or self._under_deleted(target, f.store) \
                or parent.role == "deleted":
            raise OutlookError("Folders in or into Deleted Items are not moved here; use delete_folder to delete one, "
                               "and restore it in Outlook.")
        name = prop(obj, "Name", "") or ""
        if any((prop(x, "Name", "") or "").casefold() == name.casefold() for x in each(target.Folders)):
            raise OutlookError(f"{parent.path} already has a folder named {name!r}; rename one of them first.")
        obj.MoveTo(target)
        self._invalidate()
        return self.folder(make_id(f.store, obj.EntryID)).as_dict()

    @com_errors
    def delete_folder(self, folder_id: str) -> dict:
        """Move a folder, with everything in it, into its account's Deleted Items. Never permanent."""
        f = self.folder(folder_id)
        if f.role:
            raise OutlookError(f"{f.path} is the account's {f.role} folder; it can't be deleted.")
        obj = self._folder_obj(f)
        if self._under_deleted(obj, f.store):
            raise OutlookError(f"Folder {f.path} {ALREADY_DELETED}")
        trash = self._deleted_folder_obj(f.store)
        name = prop(obj, "Name", "") or ""
        if any((prop(x, "Name", "") or "").casefold() == name.casefold() for x in each(trash.Folders)):
            raise OutlookError(f"Deleted Items already has a folder named {name!r}; rename this one first. "
                               "Nothing was deleted.")
        obj.MoveTo(trash)
        self._invalidate()
        return {"deleted": folder_id, "path": f.path, "note": DELETE_NOTE}

    @com_errors
    def save_attachment(self, message_id: str, attachment: str | int, dest_dir: str = "~/Downloads") -> dict:
        msg = self.read_message(message_id, max_chars=1)
        item, _ = self._mail_item(message_id)
        if isinstance(attachment, str) and attachment.strip() == "all":
            if not msg["attachments"]:
                raise OutlookError(f"Message {message_id} has no attachments.")
            return {"saved": [self._save_one(item, a, dest_dir) for a in msg["attachments"]]}
        return self._save_one(item, pick_attachment(msg["attachments"], attachment, message_id), dest_dir)

    @staticmethod
    def _save_one(item, chosen: dict, dest_dir: str) -> dict:
        dest = attachment_destination(dest_dir, chosen["name"])
        item.Attachments.Item(chosen["index"]).SaveAsFile(str(dest))
        if not dest.exists():
            raise OutlookError(f"Outlook reported success but {dest} does not exist.")
        return {"path": str(dest), "name": chosen["name"], "size": dest.stat().st_size}

