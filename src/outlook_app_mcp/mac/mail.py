"""macOS mail: folders, messages, search, and the changes (mark, move, delete, compose, reply)."""

from __future__ import annotations

import json
import os
import time
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
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
from .osa import (
    GS,
    RS,
    US,
    FolderChanged,
    addr,
    addrs,
    as_bool,
    check_id,
    check_ids,
    entries,
    flag,
    mailbox_key,
    priority,
    records,
    stamp_iso,
    to_stamp,
)

NEW_OUTLOOK_HINT = (
    "Outlook returned no accounts. It is most likely in New Outlook mode, which "
    "does not support AppleScript. Switch it in the Outlook menu > Legacy Outlook."
)


GOOGLE_HINT = (
    "Outlook's scripting has no account object for Google accounts, so it cannot "
    "choose a Google account as the sender. Send from an Exchange account, or write "
    "the message in Outlook yourself."
)


def accounts_config_path() -> Path:
    return Path(os.environ.get("OUTLOOK_APP_MCP_ACCOUNTS", "~/.config/outlook-app-mcp/accounts.json")).expanduser()


@dataclass
class Folder:
    id: str
    account: str
    account_name: str
    path: str
    unread: int
    total: int
    role: str
    kind: str = "exchange"
    account_source: str = ""

    @property
    def is_google(self) -> bool:
        return self.kind == "google"

    def as_dict(self) -> dict:
        d = {
            "id": self.id,
            "account": self.account,
            "account_name": self.account_name,
            "path": self.path,
            "unread": self.unread,
            "total": self.total,
            "role": self.role or None,
            "kind": self.kind,
        }
        if self.account_source:
            d["account_source"] = self.account_source
        return d


def parse_folders(raw: str) -> list[Folder]:
    return [
        Folder(f[0], f[1], f[2], f[3], int(f[4]), int(f[5]), f[6] if len(f) > 6 else "", f[7] if len(f) > 7 else "exchange")
        for f in records(raw)
    ]


def _majority(addresses: list[str], share: float) -> str | None:
    """The address that makes up at least `share` of a sample of 5 or more, else None."""
    sample = [a.strip().lower() for a in addresses if a.strip()]
    if len(sample) < 5:
        return None
    top, n = Counter(sample).most_common(1)[0]
    return top if n / len(sample) >= share else None


def infer_google_address(sent: list[str], inbox: list[str]) -> tuple[str, str] | None:
    """An account sends as itself, so a clear majority in Sent Mail is reliable. Inbox
    recipients are only a guess: mailing lists can outnumber the account's own address."""
    a = _majority(sent, 0.6)
    if a:
        return a, "sent mail"
    a = _majority(inbox, 0.5)
    if a:
        return a, "inbox recipients (guess; set it in " + str(accounts_config_path()) + " if wrong)"
    return None


def load_account_overrides() -> dict[str, str]:
    p = accounts_config_path()
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return {str(int(k)): str(v).strip() for k, v in data.items()}
    except (ValueError, AttributeError) as e:
        raise OutlookError(f"{p} must be a JSON object of root folder id -> address: {e}") from e


def parse_index(raw: str) -> list[dict]:
    return [
        {
            "id": f[0],
            "folder_id": f[1],
            "received": stamp_iso(f[2]),
            "is_read": as_bool(f[3]),
            "flag": flag(f[4]),
            "subject": f[5],
            "from": addr(f[6]),
        }
        for f in records(raw)
    ]


def parse_message(raw: str, max_chars: int, as_html: bool = False, include_headers: bool = False) -> dict:
    header, _, body = raw.partition(RS)
    f = header.split("\x1f")
    if len(f) < 17:
        raise OutlookError(f"Unexpected reply from Outlook ({len(f)} header fields).")
    atts = []
    for i, e in enumerate(entries(f[13]), 1):
        size = int(e[1]) if len(e) > 1 and e[1].strip().isdigit() else None
        atts.append({"index": i, "name": e[0], "size": size, "content_type": (e[2] if len(e) > 2 else "") or None})
    return {
        "id": f[0],
        "folder_id": f[1],
        "folder": f[2],
        "account": f[3],
        "received": stamp_iso(f[4]),
        "sent": stamp_iso(f[5]),
        "is_read": as_bool(f[6]),
        "has_html": as_bool(f[7]),
        "subject": f[8],
        "from": addr(f[9]),
        "to": addrs(f[10]),
        "cc": addrs(f[11]),
        "bcc": addrs(f[12]),
        "attachments": atts,
        "flag": flag(f[14]),
        "categories": f[15].split(GS) if f[15] else [],
        "is_meeting": as_bool(f[16]),
        "importance": priority(f[17]) if len(f) > 17 else None,
        "conversation_id": (f[18] or None) if len(f) > 18 else None,
        "meeting_event_id": (f[19] or None) if len(f) > 19 else None,
        "body_format": "html" if as_html else "text",
        "body": body[:max_chars],
        "body_truncated": len(body) > max_chars,
        "body_length": len(body),
        **({"headers": f[20] if len(f) > 20 else ""} if include_headers else {}),
    }


def _days_to_seconds(days: float | None) -> str:
    if check_since_days(days) is None:
        return ""
    return str(int(days * 86400))


class MailMixin:
    """Needs self._run (runner)."""

    _folders: list[Folder] | None
    _folders_at: float
    _google: dict[str, tuple[str, str]] | None  # Google root folder id -> (address, source)
    _mailboxes: dict[str, str]  # mailbox key -> Google account address

    # -- folders -------------------------------------------------------------

    def _invalidate(self) -> None:
        self._folders = None

    def list_folders(self, account: str | None = None, refresh: bool = False) -> list[dict]:
        return [f.as_dict() for f in self.folders(account, refresh=refresh)]

    def folders(self, account: str | None = None, refresh: bool = False) -> list[Folder]:
        fresh = time.monotonic() - self._folders_at < FOLDER_CACHE_SECONDS
        if refresh or self._folders is None or not fresh:
            folders = self._load_folders()
            if not folders:
                raise OutlookError(NEW_OUTLOOK_HINT)
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

    def _load_folders(self) -> list[Folder]:
        """Folders of every account. Google accounts come back keyed by their root
        folder ("google:151"); their address is resolved once per process, from the
        overrides file first, then from a sample of their mail."""
        if self._google is None:
            self._google = {}
        overrides = load_account_overrides()
        known = sorted(set(self._google) | set(overrides), key=int)
        raw = self._run("read/folders", known)
        folder_part, _, google_part = raw.partition(GS)
        roots: dict[str, str] = {}  # root folder id -> mailbox key
        for g in records(google_part):
            root = g[0]
            key = mailbox_key(g[3]) if len(g) > 3 else None
            if key:
                roots[root] = key
            if root in self._google or root in overrides:
                continue
            sent = [x for x in (g[1].split(US) if len(g) > 1 else []) if x]
            inbox = [x for x in (g[2].split(US) if len(g) > 2 else []) if x]
            self._google[root] = infer_google_address(sent, inbox) or (f"google:{root}", "unknown; set it in " + str(accounts_config_path()))
        folders = parse_folders(folder_part)
        self._mailboxes = {}
        for f in folders:
            if f.is_google:
                root = f.account.split(":", 1)[1]
                if root in overrides:
                    f.account, f.account_source = overrides[root], "config"
                elif root in self._google:
                    f.account, f.account_source = self._google[root]
                f.account_name = f"Google account (root folder {root})"
                if root in roots:
                    self._mailboxes[roots[root]] = f.account
        return folders

    def mailbox_account(self, exchange_id: str) -> str | None:
        """The Google account address an Exchange id belongs to, if known."""
        self.folders()
        key = mailbox_key(exchange_id)
        return self._mailboxes.get(key) if key else None

    def require_accounts(self) -> None:
        """Fail with the Legacy-mode hint when Outlook shows no accounts. New Outlook
        exposes only an empty local store to AppleScript; without this check, calendar
        and contact tools would quietly read and write that invisible store."""
        self.folders()

    def folder(self, folder_id: str | int) -> Folder:
        folder_id = check_id(folder_id, "folder_id")
        for f in self.folders():
            if f.id == folder_id:
                return f
        for f in self.folders(refresh=True):
            if f.id == folder_id:
                return f
        raise OutlookError(f"No mail folder with id {folder_id}. Use list_folders to see the ids.")

    def _role_folder(self, role: str, account: str | None) -> Folder:
        found = [f for f in self.folders(account) if f.role == role]
        if not found:
            raise OutlookError(f"Could not find the {role} folder. Pass folder_id (see list_folders).")
        if len(found) > 1:
            names = ", ".join(f"{f.account} (id {f.id})" for f in found)
            raise OutlookError(f"More than one account has the {role} folder: {names}. Pass account or folder_id.")
        return found[0]

    def _account_email(self, account: str | None) -> str:
        """'' for Outlook's default account, else the exact address of one matching
        account that scripts can write from (not a Google account)."""
        if not account:
            return ""
        matched = self.folders(account)
        emails = sorted({f.account for f in matched})
        if len(emails) > 1:
            raise OutlookError(f"{account!r} matches several accounts: {', '.join(emails)}.")
        if matched[0].is_google:
            raise OutlookError(f"{emails[0]} is a Google account. {GOOGLE_HINT}")
        return emails[0]

    def _locate(self, message_ids: Sequence[str]) -> dict[str, Folder]:
        """The folder each message is in."""
        out = {}
        for f in records(self._run("read/locate", list(message_ids))):
            out[f[0]] = self.folder(f[1])
        return out

    def _in_deleted(self, folder: Folder) -> bool:
        """In a deleted-items folder or below one. Outlook deletes from there permanently."""
        for d in self.folders():
            same_account = d.role == "deleted" and d.account == folder.account and d.kind == folder.kind
            if same_account and (folder.id == d.id or folder.path.startswith(d.path + "/")):
                return True
        return False

    # -- reading -------------------------------------------------------------

    def _index(self, folder_ids: Sequence[str], since: str, unread_only: bool) -> list[dict]:
        if since and not unread_only:
            # Bulk-reads the top of each folder and proves it complete; fast on big folders.
            script, args = "read/window", [since, *folder_ids]
        else:
            script, args = "read/index", [since, "1" if unread_only else "0", *folder_ids]
        try:
            return parse_index(self._run(script, args))
        except FolderChanged:
            try:
                return parse_index(self._run(script, args))
            except FolderChanged as exc:
                raise OutlookError("A folder kept changing while it was read. Try again.") from exc

    def _decorate(self, msgs: list[dict]) -> list[dict]:
        by_id = {f.id: f for f in self.folders()}
        for m in msgs:
            f = by_id.get(m["folder_id"])
            m["folder"] = f.path if f else None
            m["account"] = f.account if f else None
        return msgs

    def list_messages(self, folder_id=None, account=None, limit=DEFAULT_LIMIT, unread_only=False, since_days=None,
                      offset=0) -> dict:
        limit = clamp_limit(limit)
        offset = check_offset(offset)
        folder = self.folder(folder_id) if folder_id is not None else self._role_folder("inbox", account)
        since = _days_to_seconds(since_days)
        if not unread_only:
            want = offset + limit
            fast = self._recent(folder.id, want, since)
            if fast is not None:
                block = newest_first(fast)
                # A block shorter than asked holds every message in range; a full
                # block may have more behind it.
                complete = len(block) < want or folder.total <= want
                msgs = block[offset:]
                return {
                    "folder": folder.as_dict(),
                    "matched": len(block) if complete else None,
                    "offset": offset,
                    "returned": len(msgs),
                    "has_more": not complete,
                    "messages": self._decorate(msgs),
                }
        matched, msgs = self._scan(folder.id, since, unread_only, limit, offset)
        return {
            "folder": folder.as_dict(),
            "matched": matched,
            "offset": offset,
            "returned": len(msgs),
            "has_more": matched > offset + len(msgs),
            "messages": self._decorate(msgs),
        }

    def _scan(self, folder_id: str, since: str, unread_only: bool, limit: int, offset: int = 0) -> tuple[int, list[dict]]:
        """Every matching message's id and date in two passes, then full summaries of
        only the newest `limit`. Returns (number matched, summaries newest first)."""
        args = [folder_id, since, "1" if unread_only else "0"]
        for attempt in range(2):
            try:
                stamps = records(self._run("read/stamps", args))
                break
            except FolderChanged as exc:
                if attempt:
                    raise OutlookError("A folder kept changing while it was read. Try again.") from exc
        dated = sorted(((stamp_iso(f[1]) or "", int(f[0])) for f in stamps), reverse=True)
        pick = [str(mid) for _, mid in dated[offset:offset + limit]]
        msgs = parse_index(self._run("read/summaries", pick)) if pick else []
        return len(dated), newest_first(msgs)

    def _recent(self, folder_id: str, limit: int, since: str) -> list[dict] | None:
        """The newest `limit` messages read off the top of the folder, or None when
        Outlook's order could not be proven (the caller then scans the whole folder)."""
        for _ in range(2):
            try:
                status, _, rest = self._run("read/recent", [folder_id, str(limit), since]).partition(RS)
            except FolderChanged:
                continue
            return parse_index(rest) if status == "OK" else None
        return None

    def search(self, query="", folder_id=None, account=None, since_days=DEFAULT_SEARCH_DAYS, include_body=False,
               limit=DEFAULT_LIMIT, from_address=None, to_address=None, has_attachments=None, flagged=None,
               category=None, after=None, before=None, offset=0) -> dict:
        query = (query or "").strip()
        filters = [from_address, to_address, has_attachments, flagged, category, after, before]
        if not query and all(x is None for x in filters):
            raise OutlookError("Give a query or at least one filter.")
        if include_body and not query:
            raise OutlookError("include_body needs a query.")
        limit = clamp_limit(limit)
        offset = check_offset(offset)
        after_dt = parse_when(after, "after") if after else None
        before_dt = parse_when(before, "before") if before else None
        scope = (
            [self.folder(folder_id)]
            if folder_id is not None
            # Gmail's All Mail repeats every other folder; search it only when asked.
            else [f for f in self.folders(account) if f.total > 0 and f.role != "all"]
        )
        ids = [f.id for f in scope]
        if after_dt is not None:
            since = str(max(1, int((datetime.now().astimezone() - after_dt).total_seconds()) + 1))
        else:
            since = _days_to_seconds(since_days)
        found: dict[str, dict] = {}
        if ids:
            needle = query.casefold()
            index = self._index(ids, since, False)
            for m in index:
                if not query or matches(m, needle):
                    found[m["id"]] = m
            if include_body:
                raw = self._run("read/body_search", [query, since, *ids])
                body_ids = {x.strip() for x in raw.split(RS) if x.strip()}
                for m in index:
                    if m["id"] in body_ids:
                        found.setdefault(m["id"], m)
        msgs = [m for m in found.values() if self._index_filter(m, from_address, flagged, after_dt, before_dt)]
        msgs = newest_first(msgs)
        if to_address is not None or has_attachments is not None or category is not None:
            msgs = self._facts_filter(msgs, ids, since, to_address, has_attachments, category)
        page = msgs[offset:offset + limit]
        return {
            "query": query,
            "searched_folders": len(ids),
            "since_days": None if after_dt else since_days,
            "include_body": include_body,
            "matched": len(msgs),
            "offset": offset,
            "returned": len(page),
            "has_more": len(msgs) > offset + len(page),
            "messages": self._decorate(page),
        }

    @staticmethod
    def _index_filter(m: dict, from_address, flagged, after_dt, before_dt) -> bool:
        if from_address and from_address.casefold() not in (m["from"].get("address") or "").casefold():
            return False
        if flagged is not None and (m["flag"] == "flagged") != flagged:
            return False
        received = parse_when(m["received"], "received") if m.get("received") else None
        if after_dt is not None and (received is None or received < after_dt):
            return False
        return not (before_dt is not None and (received is None or received >= before_dt))

    def _facts_filter(self, msgs, folder_ids, since, to_address, has_attachments, category) -> list[dict]:
        """Filters that need each message's recipients, attachments or categories."""
        if not msgs:
            return []
        needs = ("t" if to_address else "") + ("a" if has_attachments is not None else "") + ("c" if category else "")
        wanted = sorted({m["folder_id"] for m in msgs} & set(folder_ids))
        facts = {f[0]: f for f in records(self._run("read/facts", [needs, since, *wanted]))}
        out = []
        for m in msgs:
            f = facts.get(m["id"])
            if f is None:
                continue
            addrs_ = [x.casefold() for x in (f[1].split(GS) if f[1] else [])]
            cats = [x.casefold() for x in (f[3].split(GS) if len(f) > 3 and f[3] else [])]
            if to_address and not any(to_address.strip().casefold() in x for x in addrs_):
                continue
            if has_attachments is not None and (int(f[2] or 0) > 0) != has_attachments:
                continue
            if category and category.strip().casefold() not in cats:
                continue
            out.append(m)
        return out

    def read_message(self, message_id: str, max_chars: int = DEFAULT_MAX_CHARS, include_headers: bool = False,
                     as_html: bool = False) -> dict:
        check_max_chars(max_chars)
        args = [check_id(message_id, "message_id"), "1" if include_headers else "0", "1" if as_html else "0"]
        msg = parse_message(self._run("read/message", args), max_chars, as_html, include_headers)
        if not msg["account"]:
            try:
                msg["account"] = self.folder(msg["folder_id"]).account
            except OutlookError:
                pass
        return msg

    def get_conversation(self, message_id: str, limit: int = DEFAULT_LIMIT) -> dict:
        limit = clamp_limit(limit)
        msg = self.read_message(message_id, max_chars=1)
        conv = msg.get("conversation_id")
        if not conv:
            raise OutlookError(f"Message {message_id} has no conversation id.")
        account = msg["account"]
        folders = [f.id for f in self.folders() if f.account == account and f.total > 0 and f.role != "all"]
        msgs = parse_index(self._run("read/conversation", [conv, *folders])) if folders else []
        msgs.sort(key=lambda m: m["received"] or "")
        return {
            "conversation_id": conv,
            "count": len(msgs),
            "returned": min(limit, len(msgs)),
            "messages": self._decorate(msgs[:limit]),
        }

    # -- changing ------------------------------------------------------------

    def mark_messages(self, message_ids: Sequence[str], action: str, due_date: str | None = None) -> dict:
        check_mark_action(action)
        if due_date and action != "flag":
            raise OutlookError("due_date goes with action 'flag'.")
        due = to_stamp(parse_when(due_date, "due_date")) if due_date else ""
        n = int(self._run("write/mark", [action, due, *check_ids(message_ids)]))
        self._invalidate()
        return {"action": action, "changed": n}

    def move_messages(self, message_ids: Sequence[str], folder_id: str) -> dict:
        ids = check_ids(message_ids)
        dest = self.folder(folder_id)
        raw = self._run("write/move", [dest.id, *ids])
        new_ids = [x.strip() or None for x in raw.split(RS)] if raw else []
        self._invalidate()
        result = {"moved": len(new_ids), "to": dest.as_dict(), "new_ids": dict(zip(ids, new_ids))}
        if None in new_ids:
            result["note"] = "Outlook gave some messages a new id (null above); find them with list_messages."
        return result

    def delete_messages(self, message_ids: Sequence[str]) -> dict:
        """Move messages to their account's deleted-items folder. Never permanent:
        messages already in a deleted-items folder are refused, because Outlook
        deletes those for good."""
        ids = check_ids(message_ids)
        self.folders(refresh=True)  # a folder made under Deleted Items a moment ago must count
        where = self._locate(ids)
        missing = [i for i in ids if i not in where]
        if missing:
            raise OutlookError(f"No message with id {', '.join(missing)}.")
        already = [i for i in ids if self._in_deleted(where[i])]
        if already:
            raise OutlookError(f"Message {', '.join(already)} {ALREADY_DELETED}")
        for i in ids:
            # Without a known deleted-items folder (some IMAP and POP accounts) there is no
            # way to tell whether Outlook's delete would be permanent.
            self._role_folder_of(where[i], "deleted")
        google: dict[str, list[str]] = {}
        plain = []
        for i in ids:
            f = where[i]
            if f.is_google:
                trash = self._role_folder_of(f, "deleted")
                google.setdefault(trash.id, []).append(i)
            else:
                plain.append(i)
        n = 0
        if plain:
            protected = [d.id for d in self.folders() if d.role == "deleted"]
            n += int(self._run("write/delete", [GS.join(protected), *plain]))
        for trash_id, group in google.items():
            raw = self._run("write/move", [trash_id, *group])
            n += len(raw.split(RS)) if raw else 0
        self._invalidate()
        return {"deleted": n, "note": "Moved to the account's deleted-items folder; nothing was deleted permanently."}

    def _role_folder_of(self, folder: Folder, role: str) -> Folder:
        for f in self.folders():
            if f.role == role and f.account == folder.account and f.kind == folder.kind:
                return f
        raise OutlookError(f"{folder.account} has no {role} folder Outlook can name; nothing was changed.")

    def _compose(self, mode, to, subject, body, cc, bcc, attachments, body_is_html, account, importance=None,
                 signature=None) -> dict:
        to_l, cc_l, bcc_l = check_recipients(to, cc, bcc, mode == "send")
        files = check_attachment_files(attachments)
        importance = check_priority(importance)
        html_body = body if body_is_html else to_html(body)
        html_body += self._signature_html(signature)
        acct = self._account_email(account)
        args = [mode, acct, subject, html_body, GS.join(to_l), GS.join(cc_l), GS.join(bcc_l), GS.join(files),
                importance or ""]
        mid = self._run("write/compose", args).strip()
        self._invalidate()
        return {"id": mid, "status": "sent" if mode == "send" else "draft", "account": acct or "(default)"}

    def create_draft(self, to=(), subject="", body="", cc=(), bcc=(), attachments=(), body_is_html=False,
                     account=None, importance=None, signature=None) -> dict:
        return self._compose("draft", to, subject, body, cc, bcc, attachments, body_is_html, account, importance,
                             signature)

    def send_message(self, to, subject, body, cc=(), bcc=(), attachments=(), body_is_html=False, account=None,
                     importance=None, signature=None) -> dict:
        return self._compose("send", to, subject, body, cc, bcc, attachments, body_is_html, account, importance,
                             signature)

    def _signature_html(self, name: str | None) -> str:
        if not name:
            return ""
        names = [s["name"] for s in self.list_signatures()["signatures"]]
        match = next((n for n in names if n == name), None) or next(
            (n for n in names if n.casefold() == name.strip().casefold()), None)
        if match is None:
            raise OutlookError(f"No signature named {name!r}. Signatures: {', '.join(names) or 'none'}.")
        return "<br>" + self._run("read/signatures", [match])

    def list_signatures(self) -> dict:
        sigs = [{"name": f[0], "preview": f[1] if len(f) > 1 else ""} for f in records(self._run("read/signatures", []))]
        return {"signatures": sigs}

    def update_draft(self, draft_id: str, to=None, cc=None, bcc=None, subject=None, body=None, body_is_html=False,
                     add_attachments=(), remove_attachments=(), importance=None) -> dict:
        draft_id = check_id(draft_id, "draft_id")
        msg = self.read_message(draft_id, max_chars=1)
        folder = self.folder(msg["folder_id"])
        if folder.role != "drafts":
            raise OutlookError(f"Message {draft_id} is not in a Drafts folder; only drafts can be changed.")
        pairs: list[str] = []
        for key, value in (("to", to), ("cc", cc), ("bcc", bcc)):
            if value is not None:
                pairs += [key, GS.join(check_addresses(value, key))]
        if subject is not None:
            pairs += ["subject", subject]
        if body is not None:
            pairs += ["html", body if body_is_html else to_html(body)]
        files = check_attachment_files(add_attachments)
        if files:
            pairs += ["add", GS.join(files)]
        if remove_attachments:
            idx = sorted({pick_attachment(msg["attachments"], a, draft_id)["index"] for a in remove_attachments},
                         reverse=True)
            pairs += ["remove", GS.join(str(i) for i in idx)]
        if importance is not None:
            pairs += ["importance", check_priority(importance)]
        if not pairs:
            raise OutlookError("Nothing to change.")
        self._run("write/draft", [draft_id, *pairs])
        self._invalidate()
        return {"id": draft_id, "status": "draft"}

    def send_draft(self, draft_id: str) -> dict:
        draft_id = check_id(draft_id, "draft_id")
        msg = self.read_message(draft_id, max_chars=1)
        folder = self.folder(msg["folder_id"])
        if folder.role != "drafts":
            raise OutlookError(f"Message {draft_id} is not in a Drafts folder; only drafts can be sent.")
        if folder.is_google:
            raise OutlookError(f"Draft {draft_id} is in the Google account {folder.account}. {GOOGLE_HINT}")
        if not (msg["to"] or msg["cc"] or msg["bcc"]):
            raise OutlookError(f"Draft {draft_id} has no recipients.")
        self._run("write/send_draft", [draft_id])
        self._invalidate()
        return {"id": draft_id, "status": "sent", "subject": msg["subject"]}

    def _respond(self, message_id, mode, body, to, send, cc=(), attachments=(), signature=None) -> dict:
        to_l = check_addresses(to, "to")
        cc_l = check_addresses(cc, "cc")
        files = check_attachment_files(attachments)
        sig = self._signature_html(signature)
        if mode == "forward" and send and not to_l:
            raise OutlookError("A forward to send needs at least one address in 'to'.")
        mid = check_id(message_id, "message_id")
        where = self._locate([mid])
        if mid not in where:
            raise OutlookError(f"No message with id {mid}.")
        if where[mid].is_google:
            raise OutlookError(
                f"Message {mid} is in the Google account {where[mid].account}. Outlook's scripting can't "
                "reply to or forward Google mail: it can't send from a Google account, and it won't keep the "
                "draft in that account's Drafts. Reply in Outlook yourself."
            )
        drafts = self._role_folder_of(where[mid], "drafts")
        html = (to_html(body) if body else "") + sig
        args = [mid, mode, html, body or "", GS.join(to_l), "1" if send else "0", drafts.id, GS.join(cc_l),
                GS.join(files)]
        rid = self._run("write/reply", args).strip()
        self._invalidate()
        return {"id": rid, "status": "sent" if send else "draft", "in_reply_to": mid, "mode": mode}

    def reply_to_message(self, message_id: str, body: str, reply_all: bool = False, send: bool = False, cc=(),
                         attachments=(), signature=None) -> dict:
        return self._respond(message_id, "replyall" if reply_all else "reply", body, (), send, cc, attachments,
                             signature)

    def forward_message(self, message_id: str, to: Sequence[str], body: str = "", send: bool = False, cc=(),
                        attachments=(), signature=None) -> dict:
        return self._respond(message_id, "forward", body, to, send, cc, attachments, signature)

    def create_folder(self, name: str, parent_id: str | None = None, account: str | None = None) -> dict:
        name = check_folder_name(name)
        if parent_id is not None:
            parent = self.folder(parent_id)
            args = ["create", parent.id, "", name]
        else:
            acct = self._account_email(account) or self._only_account()
            args = ["create", "", acct, name]
        fid = self._run("write/folder", args).strip()
        self._invalidate()
        return {"id": fid, "name": name}

    def _only_account(self) -> str:
        emails = sorted({f.account for f in self.folders()})
        if len(emails) != 1:
            raise OutlookError(f"Several accounts ({', '.join(emails)}); pass account or parent_id.")
        return emails[0]

    def rename_folder(self, folder_id: str, name: str) -> dict:
        name = check_folder_name(name)
        f = self.folder(folder_id)
        if f.role:
            raise OutlookError(f"{f.path} is the account's {f.role} folder; it can't be renamed.")
        self._run("write/folder", ["rename", f.id, "", name])
        self._invalidate()
        return {"id": f.id, "old_path": f.path, "name": name}

    def _movable_folder(self, folder_id: str) -> Folder:
        f = self.folder(folder_id)
        if f.role:
            raise OutlookError(f"{f.path} is the account's {f.role} folder; it can't be moved or deleted.")
        if f.is_google:
            raise OutlookError(f"{f.path} is a Google label; manage labels in Gmail or Outlook itself.")
        return f

    def move_folder(self, folder_id: str, parent_id: str) -> dict:
        f = self._movable_folder(folder_id)
        parent = self.folder(parent_id)
        if parent.account != f.account or parent.kind != f.kind:
            raise OutlookError("A folder can only move within its own account.")
        if parent.id == f.id or parent.path.startswith(f.path + "/"):
            raise OutlookError("A folder can't move into itself.")
        if self._in_deleted(parent):
            raise OutlookError("To delete a folder use delete_folder.")
        self._run("write/folder", ["move", f.id, parent.id, ""])
        self._invalidate()
        return {"id": f.id, "old_path": f.path, "parent_id": parent.id}

    def delete_folder(self, folder_id: str) -> dict:
        self.folders(refresh=True)
        f = self._movable_folder(folder_id)
        if self._in_deleted(f):
            raise OutlookError(f"{f.path} {ALREADY_DELETED}")
        trash = self._role_folder_of(f, "deleted")
        self._run("write/folder", ["move", f.id, trash.id, ""])
        self._invalidate()
        return {"deleted": f.id, "path": f.path, "note": DELETE_NOTE}

    def save_attachment(self, message_id: str, attachment: str | int, dest_dir: str = "~/Downloads") -> dict:
        message_id = check_id(message_id, "message_id")
        msg = self.read_message(message_id, max_chars=1)
        if isinstance(attachment, str) and attachment.strip().lower() == "all":
            if not msg["attachments"]:
                raise OutlookError(f"Message {message_id} has no attachments.")
            return {"saved": [self._save_one(message_id, a, dest_dir) for a in msg["attachments"]]}
        return self._save_one(message_id, pick_attachment(msg["attachments"], attachment, message_id), dest_dir)

    def _save_one(self, message_id: str, chosen: dict, dest_dir: str) -> dict:
        dest = attachment_destination(dest_dir, chosen["name"])
        self._run("write/save_attachment", [message_id, str(chosen["index"]), str(dest)])
        if not dest.exists():
            raise OutlookError(f"Outlook reported success but {dest} does not exist.")
        return {"path": str(dest), "name": chosen["name"], "size": dest.stat().st_size}
