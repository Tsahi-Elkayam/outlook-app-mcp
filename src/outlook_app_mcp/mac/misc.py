"""macOS: accounts, send/receive, categories, automatic replies, free/busy, directory."""

from __future__ import annotations

import math
import re
from collections.abc import Sequence
from datetime import timedelta

from ..common import (
    DEFAULT_LIMIT,
    OutlookError,
    check_audience,
    check_categories,
    check_category_mode,
    check_free_busy_request,
    check_ids,
    check_item_type,
    clamp_limit,
    iso,
    merge_categories,
    parse_when,
    to_html,
)
from .osa import GS, US, as_bool, records, stamp_iso, to_stamp

FREE_BUSY_DIGITS = {"free": "0", "tentative": "1", "busy": "2", "out of office": "3", "working elsewhere": "4"}

EXCHANGE_ONLY = (
    "{what}: only available through an Exchange (Microsoft 365) account in Outlook for Mac; "
    "{account} is not one."
)


def _hex(color: str) -> str | None:
    """'65535,0,0' (16-bit channels) -> '#ff0000'."""
    try:
        r, g, b = (int(x) for x in color.split(","))
    except ValueError:
        return None
    return f"#{r >> 8:02x}{g >> 8:02x}{b >> 8:02x}"


_STAMPISH = re.compile(r"^\d{4}-\d{2}-\d{2}")


def parse_free_busy(raw: str, people: Sequence[str], slots: int) -> list[dict]:
    """Outlook's flat answer -> one row per person.

    Per person: address, start (a date text), then subject, location, status for each
    interval. A new person starts where an asked-for address is followed by a date."""
    items = raw.split(GS) if raw else []
    wanted = {p.casefold() for p in people}

    def starts_person(j: int) -> bool:
        return j + 1 < len(items) and items[j].casefold() in wanted and bool(_STAMPISH.match(items[j + 1]))

    by_address: dict[str, str] = {}
    i = 0
    while i < len(items):
        if not starts_person(i):
            i += 1
            continue
        address, j, statuses = items[i], i + 2, []
        while j + 2 < len(items) and not starts_person(j):
            statuses.append(items[j + 2])
            j += 3
        by_address[address.casefold()] = "".join(FREE_BUSY_DIGITS.get(x.strip().lower(), "?") for x in statuses)
        i = j
    out = []
    for p in people:
        got = by_address.get(p.casefold())
        if got is None:
            out.append({"address": p, "slots": "?" * slots, "error": "Outlook returned no free/busy for this address."})
        else:
            out.append({"address": p, "slots": (got + "?" * slots)[:slots], "error": None})
    return out


class MiscMixin:
    """Needs self._run and the folder helpers of MailMixin."""

    # -- accounts --------------------------------------------------------------

    def list_accounts(self) -> dict:
        self.require_accounts()
        out = []
        for f in records(self._run("read/accounts", [])):
            out.append({"address": f[0], "name": f[1] or f[0], "kind": f[2], "is_default": f[3] == "1"})
        google = sorted({f.account for f in self.folders() if f.is_google and "@" in f.account})
        out += [{"address": a, "name": a, "kind": "google", "is_default": False} for a in google]
        if out and not any(a["is_default"] for a in out):
            out[0]["is_default"] = True
        return {"accounts": out}

    def send_and_receive(self) -> dict:
        emails = [a["address"] for a in self.list_accounts()["accounts"] if a["kind"] != "google"]
        if not emails:
            raise OutlookError("Outlook has no account that can be synced from a script.")
        self._run("write/sync", emails)
        return {"status": "started", "accounts": emails}

    def _exchange_account(self, account: str | None, what: str) -> str:
        accounts = self.list_accounts()["accounts"]
        if account:
            want = account.strip().casefold()
            hits = [a for a in accounts if want in a["address"].casefold() or want in a["name"].casefold()]
            if not hits:
                raise OutlookError(f"No account matches {account!r}. Accounts: "
                                   f"{', '.join(a['address'] for a in accounts)}")
            if len(hits) > 1:
                raise OutlookError(f"{account!r} matches several accounts: {', '.join(a['address'] for a in hits)}.")
            if hits[0]["kind"] != "exchange":
                raise OutlookError(EXCHANGE_ONLY.format(what=what, account=hits[0]["address"]))
            return hits[0]["address"]
        exchange = [a for a in accounts if a["kind"] == "exchange"]
        default = next((a for a in exchange if a["is_default"]), None)
        if default:
            return default["address"]
        if len(exchange) == 1:
            return exchange[0]["address"]
        if not exchange:
            raise OutlookError(f"{what} needs an Exchange (Microsoft 365) account, and Outlook has none.")
        raise OutlookError(f"Several Exchange accounts ({', '.join(a['address'] for a in exchange)}); pass account.")

    # -- categories ------------------------------------------------------------

    def list_categories(self) -> dict:
        self.require_accounts()
        cats = [{"name": f[0], "color": _hex(f[1]) if len(f) > 1 and f[1] else None}
                for f in records(self._run("read/categories", []))]
        return {"categories": cats}

    def _ensure_categories(self, names: Sequence[str]) -> list[str]:
        have = {c["name"].casefold(): c["name"] for c in self.list_categories()["categories"]}
        out = []
        for n in names:
            if n.casefold() not in have:
                self._run("write/category", ["create", n])
                have[n.casefold()] = n
            out.append(have[n.casefold()])
        return out

    def create_category(self, name: str) -> dict:
        clean = check_categories([name], allow_empty=False)[0]
        return {"name": self._ensure_categories([clean])[0]}

    def categorize(self, item_type: str, item_ids: Sequence[str], categories: Sequence[str], mode: str = "add") -> dict:
        kind = check_item_type(item_type)
        mode = check_category_mode(mode)
        names = check_categories(categories, allow_empty=(mode == "set")) or []
        ids = check_ids(item_ids, "item_ids")
        self.require_accounts()
        current = {f[0]: (f[1].split(GS) if len(f) > 1 and f[1] else [])
                   for f in records(self._run("read/item_categories", [kind, *ids]))}
        missing = [i for i in ids if i not in current]
        if missing:
            raise OutlookError(f"No {kind} with id {', '.join(missing)}; nothing was changed.")
        names = self._ensure_categories(names) if mode != "remove" else names
        args = ["assign", kind]
        for i in ids:
            args += [i, GS.join(merge_categories(current[i], names, mode))]
        n = int(self._run("write/category", args) or 0)
        return {"changed": n, "categories": names, "mode": mode}

    # -- automatic replies -------------------------------------------------------

    def get_auto_reply(self, account: str | None = None) -> dict:
        email = self._exchange_account(account, "Automatic replies")
        f = self._run("read/autoreply", [email]).split("\x1f")
        to_external, only_known = as_bool(f[4]), as_bool(f[5])
        return {
            "account": email,
            "enabled": as_bool(f[0]),
            "scheduled": as_bool(f[1]),
            "start": stamp_iso(f[2]) if f[2] else None,
            "end": stamp_iso(f[3]) if f[3] else None,
            "internal_message": f[6] if len(f) > 6 else "",
            "external_message": f[7] if len(f) > 7 else "",
            "external_audience": "none" if not to_external else ("known" if only_known else "all"),
        }

    def set_auto_reply(self, account=None, enabled=False, internal_message=None, external_message=None,
                       external_audience=None, start=None, end=None) -> dict:
        audience = check_audience(external_audience)
        email = self._exchange_account(account, "Automatic replies")
        s = parse_when(start, "start") if start else None
        e = parse_when(end, "end") if end else None
        if (s is None) != (e is None):
            raise OutlookError("Give both start and end to schedule automatic replies, or neither.")
        if s and e and e <= s:
            raise OutlookError("end must be after start.")
        cur = self.get_auto_reply(email)
        internal = cur["internal_message"] if internal_message is None else to_html(internal_message)
        external = cur["external_message"] if external_message is None else to_html(external_message)
        audience = audience or cur["external_audience"]
        self._run("write/autoreply", [
            email, "1" if enabled else "0", "1" if s else "0", to_stamp(s) if s else "", to_stamp(e) if e else "",
            internal or "", external or "", "0" if audience == "none" else "1", "1" if audience == "known" else "0",
        ])
        return self.get_auto_reply(email)

    # -- free/busy ---------------------------------------------------------------

    def get_free_busy(self, addresses: Sequence[str], start: str, end: str, interval_minutes: int = 30) -> dict:
        people, s, e = check_free_busy_request(addresses, start, end, interval_minutes)
        email = self._exchange_account(None, "Free/busy")
        slots = math.ceil((e - s) / timedelta(minutes=interval_minutes))
        raw = self._run("read/freebusy", [email, to_stamp(s), to_stamp(e), str(interval_minutes), *people])
        return {"start": iso(s), "end": iso(e), "interval_minutes": interval_minutes,
                "people": parse_free_busy(raw, people, slots)}

    # -- organization directory ----------------------------------------------------

    def lookup_directory(self, query: str, limit: int = DEFAULT_LIMIT) -> dict:
        q = (query or "").strip()
        if not q:
            raise OutlookError("query must not be empty.")
        limit = clamp_limit(limit)
        email = self._exchange_account(None, "The organization directory")
        found = self._resolve(email, q)
        words = q.split()
        if not found and len(words) > 1:
            # Exchange matches the start of one name part; look up the first word, keep those with all words.
            found = [p for p in self._resolve(email, words[0])
                     if all(w.casefold() in (p["name"] + " " + p["address"]).casefold() for w in words[1:])]
        return {"query": q, "people": found[:limit], "account": email}

    def _resolve(self, email: str, text: str) -> list[dict]:
        people, seen = [], set()
        for f in records(self._run("read/directory", [email, text])):
            name, _, address = f[0].partition(US)
            if address and address.casefold() not in seen:
                seen.add(address.casefold())
                people.append({"name": name, "address": address, "title": "", "department": "", "office": "",
                               "phone": ""})
        return people
