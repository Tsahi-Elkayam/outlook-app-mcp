"""macOS contacts and contact groups in Outlook's address books (not the organization directory)."""

from __future__ import annotations

from collections.abc import Sequence

from ..common import (
    DEFAULT_LIMIT,
    DELETE_NOTE,
    OutlookError,
    check_addresses,
    check_contact_fields,
    check_title,
    parse_when,
)
from .osa import GS, RS, US, FolderChanged, check_id, records, stamp_iso, to_stamp

FIELDS = {  # tool argument -> script key
    "first_name": "first",
    "middle_name": "middle",
    "last_name": "last",
    "nickname": "nickname",
    "company": "company",
    "job_title": "jobtitle",
    "department": "department",
    "office": "office",
    "mobile": "mobile",
    "business_phone": "business",
    "home_phone": "home",
    "home_street": "hstreet",
    "home_city": "hcity",
    "home_state": "hstate",
    "home_zip": "hzip",
    "home_country": "hcountry",
    "business_street": "bstreet",
    "business_city": "bcity",
    "business_state": "bstate",
    "business_zip": "bzip",
    "business_country": "bcountry",
    "web_page": "web",
    "note": "note",
}

# read/contact field order after the id.
READ_ORDER = (
    "display_name", "first_name", "middle_name", "last_name", "nickname", "company", "job_title", "department",
    "office", "mobile", "business_phone", "home_phone", "home_street", "home_city", "home_state", "home_zip",
    "home_country", "business_street", "business_city", "business_state", "business_zip", "business_country",
    "web_page",
)

NOT_EXCHANGE_DELETE = (
    "{what} {id} is in {where}, not in an Exchange account's address book, and Outlook deletes such items for "
    "good. Nothing was deleted; delete it in Outlook if you mean to."
)


def parse_contacts(raw: str) -> list[dict]:
    return [
        {
            "id": f[0],
            "display_name": f[1],
            "first_name": f[2],
            "last_name": f[3],
            "company": f[4],
            "job_title": f[5],
            "emails": f[6].split(GS) if f[6] else [],
            "mobile": f[7],
            "business_phone": f[8],
            "home_phone": f[9],
            "categories": f[10].split(GS) if len(f) > 10 and f[10] else [],
        }
        for f in records(raw)
    ]


def parse_contact(raw: str) -> dict:
    header, _, note = raw.partition(RS)
    f = header.split("\x1f")
    c: dict = {"id": f[0]}
    for i, key in enumerate(READ_ORDER, start=1):
        c[key] = f[i] if i < len(f) else ""
    n = len(READ_ORDER) + 1
    rest = f[n:] + [""] * 6
    c["birthday"] = (stamp_iso(rest[0]) or "")[:10] or None if rest[0] else None
    c["emails"] = rest[1].split(GS) if rest[1] else []
    c["categories"] = rest[2].split(GS) if rest[2] else []
    c["note"] = note
    c["_address_book"], c["_account"], c["_kind"] = rest[3], rest[4], rest[5]
    return c


def _public(c: dict) -> dict:
    out = {k: v for k, v in c.items() if not k.startswith("_")}
    out["account"] = c.get("_account") or None
    return out


def _digits(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())


def _group(f: list[str]) -> dict:
    members = []
    for e in f[2].split(GS) if len(f) > 2 and f[2] else []:
        name, _, address = e.partition(US)
        members.append({"name": name, "address": address})
    return {"id": f[0], "name": f[1], "members": members, "account": (f[4] if len(f) > 4 else "") or None,
            "_kind": f[5] if len(f) > 5 else ""}


class ContactsMixin:
    """Needs self._run, require_accounts and _ensure_categories."""

    def _all_contacts(self) -> list[dict]:
        self.require_accounts()
        try:
            return parse_contacts(self._run("read/contacts", []))
        except FolderChanged:
            return parse_contacts(self._run("read/contacts", []))

    def search_contacts(self, query: str, limit: int = DEFAULT_LIMIT) -> dict:
        q = query.strip().casefold()
        if not q:
            raise OutlookError("query must not be empty.")
        qd = _digits(q)
        found = []
        for c in self._all_contacts():
            text = [c["display_name"], c["first_name"], c["last_name"], c["company"], c["job_title"], *c["emails"]]
            phones = [c["mobile"], c["business_phone"], c["home_phone"]]
            if any(q in (t or "").casefold() for t in text) or (len(qd) >= 4 and any(qd in _digits(p) for p in phones)):
                found.append(c)
        return {"query": query, "matched": len(found), "contacts": found[: max(1, limit)]}

    def _read_contact(self, contact_id: str) -> dict:
        contact_id = check_id(contact_id, "contact_id")
        self.require_accounts()
        return parse_contact(self._run("read/contact", [contact_id]))

    def read_contact(self, contact_id: str) -> dict:
        return _public(self._read_contact(contact_id))

    def _contact_pairs(self, emails: Sequence[str] | None, fields: dict) -> list[str]:
        fields = check_contact_fields(fields)
        pairs: list[str] = []
        for arg, key in FIELDS.items():
            v = fields.get(arg)
            if v is not None:
                pairs += [key, v]
        if "birthday" in fields:
            b = fields["birthday"]
            if not b:
                raise OutlookError("Outlook for Mac's scripting can set a birthday but not clear it; nothing was "
                                   "changed. Clear it in Outlook.")
            pairs += ["birthday", to_stamp(parse_when(f"{b}T12:00", "birthday"))]
        if emails is not None:
            pairs += ["emails", GS.join(check_addresses(emails, "emails"))]
        if fields.get("categories") is not None:
            pairs += ["categories", GS.join(self._ensure_categories(fields["categories"]))]
        return pairs

    def _address_book_for(self, account: str | None) -> str:
        if not account:
            return ""
        want = account.strip().casefold()
        for f in records(self._run("read/address_books", [])):
            if f[1] and want in f[1].casefold():
                return f[0]
        raise OutlookError(f"No address book of an account matching {account!r}.")

    def create_contact(self, emails: Sequence[str] | None = None, **fields) -> dict:
        account = fields.pop("account", None)
        pairs = self._contact_pairs(emails, fields)
        if not any(fields.get(k) for k in ("first_name", "last_name", "company")) and not emails:
            raise OutlookError("Give at least a name, a company or an email.")
        self.require_accounts()
        book = self._address_book_for(account)
        cid = self._run("write/contact", ["create", book, *pairs]).strip()
        return self.read_contact(cid)

    def update_contact(self, contact_id: str, emails: Sequence[str] | None = None, **fields) -> dict:
        pairs = self._contact_pairs(emails, fields)
        if not pairs:
            raise OutlookError("Nothing to change.")
        contact_id = self._read_contact(contact_id)["id"]
        self._run("write/contact", ["update", contact_id, *pairs])
        return self.read_contact(contact_id)

    def delete_contact(self, contact_id: str) -> dict:
        c = self._read_contact(contact_id)
        if c["_kind"] != "exchange":
            raise OutlookError(NOT_EXCHANGE_DELETE.format(
                what="Contact", id=c["id"], where=c["_account"] or "an address book on this Mac"))
        self._run("write/contact", ["delete", c["id"]])
        return {"deleted": c["id"], "name": c["display_name"], "note": DELETE_NOTE}

    # -- groups ------------------------------------------------------------------

    def _groups(self) -> list[dict]:
        self.require_accounts()
        return [_group(f) for f in records(self._run("read/groups", []))]

    def _group(self, group_id: str) -> dict:
        group_id = check_id(group_id, "group_id")
        for g in self._groups():
            if g["id"] == group_id:
                return g
        raise OutlookError(f"No contact group with id {group_id}.")

    @staticmethod
    def _public_group(g: dict) -> dict:
        return {k: v for k, v in g.items() if not k.startswith("_")}

    def list_contact_groups(self, query: str | None = None) -> dict:
        needle = (query or "").strip().casefold()
        groups = [self._public_group(g) for g in self._groups() if not needle or needle in g["name"].casefold()]
        return {"groups": groups}

    @staticmethod
    def _members_arg(members: list[dict]) -> str:
        return GS.join(f"{m['name'] or m['address']}{US}{m['address']}" for m in members)

    def create_contact_group(self, name: str, members: Sequence[str] = ()) -> dict:
        name = check_title(name, "name")
        addresses = list(dict.fromkeys(check_addresses(members, "members")))
        self.require_accounts()
        gid = self._run("write/group", ["create", "", name,
                                        self._members_arg([{"name": "", "address": a} for a in addresses])]).strip()
        return self._public_group(self._group(gid))

    def update_contact_group(self, group_id: str, name: str | None = None, add_members=(), remove_members=()) -> dict:
        g = self._group(group_id)
        new_name = check_title(name, "name") if name is not None else ""
        add = check_addresses(add_members, "add_members")
        remove = {a.casefold() for a in check_addresses(remove_members, "remove_members")}
        if not new_name and not add and not remove:
            raise OutlookError("Nothing to change.")
        missing = remove - {m["address"].casefold() for m in g["members"]}
        if missing:
            raise OutlookError(f"Not in the group: {', '.join(sorted(missing))}. Nothing was changed.")
        members = [m for m in g["members"] if m["address"].casefold() not in remove]
        have = {m["address"].casefold() for m in members}
        members += [{"name": "", "address": a} for a in dict.fromkeys(add) if a.casefold() not in have]
        changed_members = bool(add) or bool(remove)
        self._run("write/group", ["update", g["id"], new_name,
                                  self._members_arg(members) if changed_members else "-"])
        return self._public_group(self._group(g["id"]))

    def delete_contact_group(self, group_id: str) -> dict:
        g = self._group(group_id)
        if g["_kind"] != "exchange":
            raise OutlookError(NOT_EXCHANGE_DELETE.format(
                what="Group", id=g["id"], where=g["account"] or "an address book on this Mac"))
        self._run("write/group", ["delete", g["id"]])
        return {"deleted": g["id"], "name": g["name"], "note": DELETE_NOTE}
