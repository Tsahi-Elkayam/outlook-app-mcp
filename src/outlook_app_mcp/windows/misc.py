"""Windows: accounts, send/receive, categories, out of office, free/busy and the organization directory."""

from __future__ import annotations

import math
from collections.abc import Sequence
from datetime import datetime, timedelta

from ..common import (
    DEFAULT_LIMIT,
    OutlookError,
    check_categories,
    check_category_mode,
    check_free_busy_request,
    check_ids,
    check_item_type,
    clamp_limit,
    iso,
    merge_categories,
)
from .com import (
    ACCOUNT_KINDS,
    CATEGORY_COLORS,
    PR_OOF_STATE,
    com_errors,
    each,
    explain_com,
    mapi,
    prop,
    store_key,
    to_com,
    walk_folders,
    write_categories,
)
from .com import (
    categories as category_list,
)

GAL_SCAN_CAP = 2000  # directory entries looked at by lookup_directory; a real directory can hold 100000
GOOGLE_DOMAINS = ("@gmail.com", "@googlemail.com")
AUTO_REPLY_NOTE = (
    "Classic Outlook's object model only shows whether automatic replies are on (read from the mailbox). The "
    "schedule, the messages and the audience are not available; see them in File > Automatic Replies."
)
NO_AUTO_REPLY_SETTER = (
    "Classic Outlook for Windows has no object-model call for automatic replies (out of office), so this "
    "server can't change them. In Outlook choose File > Automatic Replies, or use Outlook on the web."
)
NO_DIRECTORY = (
    "The organization directory needs an Exchange / Microsoft 365 account; Outlook has no address list of that "
    "kind here."
)


class WinMiscMixin:
    """Needs the store helpers of WinMailMixin, and the item finders of the calendar, tasks, notes and
    contacts mixins."""

    # -- accounts --------------------------------------------------------------

    @com_errors
    def list_accounts(self) -> dict:
        try:
            default_sid = self._ns.DefaultStore.StoreID
        except Exception:  # noqa: BLE001 - no default store: the first account stands in
            default_sid = ""
        out = []
        for acc in each(self._ns.Accounts):
            smtp = prop(acc, "SmtpAddress", "") or ""
            kind = ACCOUNT_KINDS.get(prop(acc, "AccountType", 5), "other")
            if kind == "imap" and smtp.lower().endswith(GOOGLE_DOMAINS):
                kind = "google"
            try:
                sid = acc.DeliveryStore.StoreID
            except Exception:  # noqa: BLE001 - accounts without a delivery store
                sid = None
            out.append({"address": smtp, "name": prop(acc, "DisplayName", "") or smtp, "kind": kind,
                        "is_default": bool(sid) and sid == default_sid})
        if out and not any(a["is_default"] for a in out):
            out[0]["is_default"] = True
        return {"accounts": out}

    @com_errors
    def send_and_receive(self) -> dict:
        """Start Outlook's send/receive for all accounts. Mail waiting in the Outbox leaves."""
        self._ns.SendAndReceive(False)
        return {"status": "started"}

    # -- categories ------------------------------------------------------------

    def _master_categories(self) -> dict[str, str]:
        """Outlook's category list: lower-case name -> name as spelled there."""
        return {(prop(c, "Name", "") or "").casefold(): prop(c, "Name", "") or "" for c in each(self._ns.Categories)}

    def _ensure_categories(self, names: Sequence[str]) -> list[str]:
        """The names as spelled in Outlook's category list, adding any that are not in it yet."""
        have = self._master_categories()
        out = []
        for n in names:
            if n.casefold() not in have:
                self._ns.Categories.Add(n)
                have[n.casefold()] = n
            out.append(have[n.casefold()])
        return out

    def _apply_categories(self, item, names: Sequence[str]) -> None:
        """Give an item exactly these categories (creating missing ones). Does not save."""
        write_categories(item, self._ensure_categories(names))

    @com_errors
    def list_categories(self) -> dict:
        out = []
        for c in each(self._ns.Categories):
            color = prop(c, "Color", 0)
            out.append({"name": prop(c, "Name", "") or "", "color": CATEGORY_COLORS.get(color, color)})
        return {"categories": out}

    @com_errors
    def create_category(self, name: str) -> dict:
        clean = check_categories([name], allow_empty=False)[0]
        return {"name": self._ensure_categories([clean])[0]}

    def _typed_item(self, item_id: str, kind: str):
        finder = {"message": self._mail_item, "event": self._event_item, "task": self._task_item,
                  "contact": self._contact_item, "note": self._note_item}[kind]
        return finder(item_id)[0]

    @com_errors
    def categorize(self, item_type: str, item_ids: Sequence[str], categories: Sequence[str],
                   mode: str = "add") -> dict:
        kind = check_item_type(item_type)
        mode = check_category_mode(mode)
        ids = check_ids(item_ids, "item_ids")
        names = check_categories(categories, allow_empty=mode == "set") or []
        items = [self._typed_item(i, kind) for i in ids]  # every id is checked before anything changes
        if mode != "remove":
            names = self._ensure_categories(names)
        for item in items:
            write_categories(item, merge_categories(category_list(item), names, mode))
            item.Save()
        self._invalidate()
        return {"changed": len(items), "categories": names}

    # -- folders of one kind of item -------------------------------------------

    def _kind_folders(self, item_type: int):
        """(Store, COM folder) for every folder of this item type in every open mailbox or data file."""
        self.require_accounts()
        return [(store, f) for st, store in self._open_stores() for f in walk_folders(st.GetRootFolder(), item_type)]

    def _default_folder(self, const: int, account: str | None, what: str):
        """(COM folder, store key) of the default Tasks / Notes folder of an account, or of Outlook's default
        mailbox when no account is given."""
        self.require_accounts()
        label = account or "Outlook's default mailbox"
        try:
            if account:
                folder = self._com_store_for(account)[0].GetDefaultFolder(const)
            else:
                folder = self._ns.GetDefaultFolder(const)
        except self._com_error as exc:
            raise OutlookError(f"{label} has no {what} folder (IMAP and POP accounts keep them in the default "
                               "data file; pass the account that has one).") from exc
        return folder, store_key(folder.Store.StoreID)

    # -- out of office ---------------------------------------------------------

    def _com_store_for(self, account: str | None):
        """(COM store, account label) for an account, or Outlook's default store."""
        if account:
            keys = {f.store for f in self.folders(account)}
            if len(keys) > 1:
                raise OutlookError(f"{account!r} matches several accounts; be more exact.")
            key = next(iter(keys))
        else:
            try:
                key = store_key(self._ns.DefaultStore.StoreID)
            except Exception as exc:
                raise OutlookError("Outlook has no default mailbox; pass account.") from exc
        for st in each(self._ns.Stores):
            if store_key(st.StoreID) == key:
                return st, self._store(key).account
        raise OutlookError("That account's mailbox is not open in Outlook.")

    @com_errors
    def get_auto_reply(self, account: str | None = None) -> dict:
        store, label = self._com_store_for(account)
        state = mapi(store, PR_OOF_STATE)
        result = {
            "account": label, "enabled": None if state is None else bool(state), "scheduled": None, "start": None,
            "end": None, "internal_message": None, "external_message": None, "external_audience": None,
            "note": AUTO_REPLY_NOTE,
        }
        if state is None:
            result["note"] = "Outlook did not report whether automatic replies are on for this mailbox. " + AUTO_REPLY_NOTE
        return result

    @com_errors
    def set_auto_reply(self, account=None, enabled=False, internal_message=None, external_message=None,
                       external_audience=None, start=None, end=None) -> dict:
        raise OutlookError(NO_AUTO_REPLY_SETTER)

    # -- free/busy -------------------------------------------------------------

    @com_errors
    def get_free_busy(self, addresses: Sequence[str], start: str, end: str, interval_minutes: int = 30) -> dict:
        """One digit per interval for each address, from the free/busy Outlook holds for them.

        Outlook answers with a string that begins at midnight of the day asked for, one character per
        interval, for as far ahead as the person's free/busy is published. The slice for [start, end) is
        cut out of it; intervals past its end are '?'."""
        people, s, e = check_free_busy_request(addresses, start, end, interval_minutes)
        slots = math.ceil((e - s) / timedelta(minutes=interval_minutes))
        local_start = s.astimezone()
        midnight = datetime(local_start.year, local_start.month, local_start.day).astimezone()
        first = int((local_start - midnight).total_seconds() // 60 // interval_minutes)
        out = []
        for address in people:
            row = {"address": address, "slots": "?" * slots, "error": None}
            out.append(row)
            rec = self._ns.CreateRecipient(address)
            if not rec.Resolve():
                row["error"] = "Outlook could not resolve this address."
                continue
            try:
                raw = rec.FreeBusy(to_com(midnight), interval_minutes, True) or ""
            except self._com_error as exc:
                row["error"] = explain_com(exc)
                continue
            if not raw:
                row["error"] = "No free/busy information is published for this person."
                continue
            row["slots"] = "".join(raw[first + i] if first + i < len(raw) else "?" for i in range(slots))
        return {"start": iso(s), "end": iso(e), "interval_minutes": interval_minutes, "people": out}

    # -- organization directory ---------------------------------------------------

    @staticmethod
    def _directory_person(entry) -> dict | None:
        try:
            user = entry.GetExchangeUser()
        except Exception:  # noqa: BLE001 - not a person (a distribution list, a one-off address)
            user = None
        if user is None:
            return None
        return {
            "name": prop(user, "Name", "") or prop(entry, "Name", "") or "",
            "address": prop(user, "PrimarySmtpAddress", "") or "",
            "title": prop(user, "JobTitle", "") or "",
            "department": prop(user, "Department", "") or "",
            "office": prop(user, "OfficeLocation", "") or "",
            "phone": prop(user, "BusinessTelephoneNumber", "") or "",
        }

    @com_errors
    def lookup_directory(self, query: str, limit: int = DEFAULT_LIMIT) -> dict:
        """People in the organization's address list (the GAL), by part of a name.

        Outlook has no search call for the GAL, so this resolves the text as a name (which finds an exact or
        unique match at once) and then walks the list, comparing names, for at most GAL_SCAN_CAP entries.
        When the list is longer than that the result says so; a more exact name finds the person."""
        q = (query or "").strip()
        if not q:
            raise OutlookError("query must not be empty.")
        limit = clamp_limit(limit)
        try:
            gal = self._ns.GetGlobalAddressList()
        except self._com_error:
            gal = None
        if gal is None:
            raise OutlookError(NO_DIRECTORY)
        people: list[dict] = []
        seen: set[str] = set()

        def add(entry) -> None:
            person = self._directory_person(entry)
            if person and person["address"].casefold() not in seen:
                seen.add(person["address"].casefold())
                people.append(person)

        rec = self._ns.CreateRecipient(q)
        if rec.Resolve():
            add(prop(rec, "AddressEntry"))
        needle = q.casefold()
        entries = gal.AddressEntries
        entry, scanned = entries.GetFirst(), 0
        while entry is not None and scanned < GAL_SCAN_CAP and len(people) < limit:
            scanned += 1
            if needle in (prop(entry, "Name", "") or "").casefold():
                add(entry)
            entry = entries.GetNext()
        result = {"query": q, "people": people[:limit]}
        if entry is not None and len(people) < limit:
            result["note"] = (f"Only the first {GAL_SCAN_CAP} entries of the directory were searched; "
                              "use a more exact name.")
        return result
