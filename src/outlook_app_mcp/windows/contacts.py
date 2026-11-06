"""Windows contacts: every contacts folder Outlook has open (not the organization directory)."""

from __future__ import annotations

from collections.abc import Sequence

from ..common import (
    ALREADY_DELETED,
    DEFAULT_LIMIT,
    DELETE_NOTE,
    OutlookError,
    check_addresses,
    check_contact_fields,
    check_title,
    parse_when,
)
from .com import (
    FOLDER_CONTACTS,
    ITEM_CONTACT,
    ITEM_DISTLIST,
    NO_DATE,
    com_errors,
    date_com,
    day_to_com,
    each,
    make_id,
    prop,
    recipient_address,
    split_categories,
    store_key,
    walk_folders,
)
from .mail import table_rows

FIELDS = {  # tool argument -> ContactItem property
    "first_name": "FirstName",
    "middle_name": "MiddleName",
    "last_name": "LastName",
    "nickname": "NickName",
    "company": "CompanyName",
    "job_title": "JobTitle",
    "department": "Department",
    "office": "OfficeLocation",
    "mobile": "MobileTelephoneNumber",
    "business_phone": "BusinessTelephoneNumber",
    "home_phone": "HomeTelephoneNumber",
    "home_street": "HomeAddressStreet",
    "home_city": "HomeAddressCity",
    "home_state": "HomeAddressState",
    "home_zip": "HomeAddressPostalCode",
    "home_country": "HomeAddressCountry",
    "business_street": "BusinessAddressStreet",
    "business_city": "BusinessAddressCity",
    "business_state": "BusinessAddressState",
    "business_zip": "BusinessAddressPostalCode",
    "business_country": "BusinessAddressCountry",
    "web_page": "WebPage",
}
EMAILS = ("Email1Address", "Email2Address", "Email3Address")
# What a folder Table is asked for when searching: the fields search looks in, plus categories. The other
# fields are read from the opened item (read_contact), where every property is certain to exist.
SEARCH_FIELDS = ("first_name", "last_name", "company", "job_title", "mobile", "business_phone", "home_phone")
COLUMNS = ("EntryID", "MessageClass", "FullName", *(FIELDS[f] for f in SEARCH_FIELDS), *EMAILS, "Categories")


def _digits(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())


def _contact(values: dict, key: str) -> dict:
    return {
        "id": make_id(key, values["EntryID"]),
        "display_name": values.get("FullName") or "",
        "first_name": values.get("FirstName") or "",
        "last_name": values.get("LastName") or "",
        "company": values.get("CompanyName") or "",
        "job_title": values.get("JobTitle") or "",
        "emails": [e for e in (values.get(x) for x in EMAILS) if e],
        "mobile": values.get("MobileTelephoneNumber") or "",
        "business_phone": values.get("BusinessTelephoneNumber") or "",
        "home_phone": values.get("HomeTelephoneNumber") or "",
        "categories": split_categories(values.get("Categories")),
    }


class WinContactsMixin:
    """Needs self._ns and the store helpers of WinMailMixin."""

    def _all_contacts(self) -> list[dict]:
        self.require_accounts()
        out = []
        for st, store in self._open_stores():
            key = store.key
            for folder in walk_folders(st.GetRootFolder(), ITEM_CONTACT):
                table = folder.GetTable()
                cols = table.Columns
                cols.RemoveAll()
                for c in COLUMNS:
                    cols.Add(c)
                for row in table_rows(table, COLUMNS, None):
                    if (row.get("MessageClass") or "").startswith("IPM.Contact"):
                        out.append(_contact(row, key))
        return out

    @com_errors
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

    def _read_contact(self, item, key: str) -> dict:
        """The contact as search_contacts shows it, with every other field and the note."""
        values = {name: prop(item, name, "") for name in (*COLUMNS, *FIELDS.values())}
        values["EntryID"] = item.EntryID
        values["Categories"] = prop(item, "Categories", "")
        out = _contact(values, key)
        for arg, name in FIELDS.items():
            out[arg] = prop(item, name, "") or ""
        out["birthday"] = date_com(prop(item, "Birthday"))
        out["note"] = prop(item, "Body", "") or ""
        return out

    def _contact_item(self, contact_id: str, what: str = "contact_id"):
        item, key = self._item(contact_id, what)
        if not (prop(item, "MessageClass", "") or "").startswith("IPM.Contact"):
            raise OutlookError(f"{contact_id} is not a contact.")
        return item, key

    @com_errors
    def read_contact(self, contact_id: str) -> dict:
        self.require_accounts()
        item, key = self._contact_item(contact_id)
        return self._read_contact(item, key)

    @staticmethod
    def _check_emails(emails: Sequence[str] | None) -> list[str] | None:
        if emails is None:
            return None
        checked = check_addresses(emails, "emails")
        if len(checked) > len(EMAILS):
            raise OutlookError(f"Outlook for Windows keeps at most {len(EMAILS)} email addresses per contact.")
        return checked

    @staticmethod
    def _set(item, emails: list[str] | None, fields: dict) -> bool:
        changed = False
        for arg, name in FIELDS.items():
            v = fields.get(arg)
            if v is not None:
                setattr(item, name, v)
                changed = True
        if fields.get("note") is not None:
            item.Body = fields["note"]
            changed = True
        if fields.get("birthday") is not None:
            # "" clears it. Outlook also keeps a yearly entry in the Birthdays calendar in step with this.
            item.Birthday = NO_DATE if fields["birthday"] == "" else day_to_com(parse_when(fields["birthday"]))
            changed = True
        if emails is not None:
            for i, name in enumerate(EMAILS):
                setattr(item, name, emails[i] if i < len(emails) else "")
            changed = True
        return changed

    @com_errors
    def create_contact(self, emails: Sequence[str] | None = None, **fields) -> dict:
        fields = check_contact_fields(fields)
        checked = self._check_emails(emails)
        if not any(fields.get(k) for k in ("first_name", "last_name", "company")) and not checked:
            raise OutlookError("Give at least a name, a company or an email.")
        self.require_accounts()
        folder = self._ns.GetDefaultFolder(FOLDER_CONTACTS)
        item = folder.Items.Add(ITEM_CONTACT)
        self._set(item, checked, fields)
        if fields.get("categories"):
            self._apply_categories(item, fields["categories"])
        item.Save()
        return self._read_contact(item, store_key(folder.Store.StoreID))

    @com_errors
    def update_contact(self, contact_id: str, emails: Sequence[str] | None = None, **fields) -> dict:
        fields = check_contact_fields(fields)
        checked = self._check_emails(emails)
        if checked is None and not any(fields.get(k) is not None for k in (*FIELDS, "note", "birthday", "categories")):
            raise OutlookError("Nothing to change.")
        item, key = self._contact_item(contact_id)
        self._set(item, checked, fields)
        if fields.get("categories") is not None:
            self._apply_categories(item, fields["categories"])
        item.Save()
        return self._read_contact(item, key)

    @com_errors
    def delete_contact(self, contact_id: str) -> dict:
        item, key = self._contact_item(contact_id)
        if self._under_deleted(item, key):
            raise OutlookError(f"Contact {contact_id} {ALREADY_DELETED}")
        name = prop(item, "FullName", "") or ""
        item.Move(self._deleted_folder_obj(key))  # Moved, never Delete(): Delete inside Deleted Items is permanent.
        return {"deleted": contact_id, "name": name, "note": DELETE_NOTE}

    # -- contact groups (distribution lists) -------------------------------------

    def _group_item(self, group_id: str, what: str = "group_id"):
        item, key = self._item(group_id, what)
        if not (prop(item, "MessageClass", "") or "").startswith("IPM.DistList"):
            raise OutlookError(f"{group_id} is not a contact group.")
        return item, key

    @staticmethod
    def _group(item, key: str) -> dict:
        members = []
        for i in range(1, int(prop(item, "MemberCount", 0) or 0) + 1):
            m = item.GetMember(i)
            members.append({"name": prop(m, "Name", "") or "", "address": recipient_address(m)})
        return {"id": make_id(key, item.EntryID), "name": prop(item, "DLName", "") or "", "members": members}

    @com_errors
    def list_contact_groups(self, query: str | None = None) -> dict:
        self.require_accounts()
        needle = (query or "").strip().casefold()
        groups = []
        for st, store in self._open_stores():
            for folder in walk_folders(st.GetRootFolder(), ITEM_CONTACT):
                for it in each(folder.Items):
                    if not (prop(it, "MessageClass", "") or "").startswith("IPM.DistList"):
                        continue
                    if needle and needle not in (prop(it, "DLName", "") or "").casefold():
                        continue
                    groups.append(self._group(it, store.key))
        return {"groups": groups}

    def _resolved(self, address: str):
        rec = self._ns.CreateRecipient(address)
        if not rec.Resolve():
            raise OutlookError(f"Outlook could not resolve {address}; nothing was changed.")
        return rec

    @com_errors
    def create_contact_group(self, name: str, members: Sequence[str] = ()) -> dict:
        name = check_title(name, "name")
        addresses = list(dict.fromkeys(check_addresses(members, "members")))
        self.require_accounts()
        folder = self._ns.GetDefaultFolder(FOLDER_CONTACTS)
        item = folder.Items.Add(ITEM_DISTLIST)
        item.DLName = name
        for a in addresses:
            item.AddMember(self._resolved(a))
        item.Save()
        return self._group(item, store_key(folder.Store.StoreID))

    @com_errors
    def update_contact_group(self, group_id: str, name: str | None = None, add_members: Sequence[str] = (),
                             remove_members: Sequence[str] = ()) -> dict:
        new_name = check_title(name, "name") if name is not None else None
        to_add = list(dict.fromkeys(check_addresses(add_members, "add_members")))
        to_drop = list(dict.fromkeys(check_addresses(remove_members, "remove_members")))
        if new_name is None and not to_add and not to_drop:
            raise OutlookError("Nothing to change.")
        item, key = self._group_item(group_id)
        members = self._group(item, key)["members"]
        have = {m["address"].casefold() for m in members}
        missing = [a for a in to_drop if a.casefold() not in have]
        if missing:
            raise OutlookError(f"Not members of this group: {', '.join(missing)}. Nothing was changed.")
        recipients_add = [self._resolved(a) for a in to_add if a.casefold() not in have]
        recipients_drop = [self._resolved(a) for a in to_drop]
        if new_name is not None:
            item.DLName = new_name
        for rec in recipients_drop:
            item.RemoveMember(rec)
        for rec in recipients_add:
            item.AddMember(rec)
        item.Save()
        return self._group(item, key)

    @com_errors
    def delete_contact_group(self, group_id: str) -> dict:
        item, key = self._group_item(group_id)
        if self._under_deleted(item, key):
            raise OutlookError(f"Group {group_id} {ALREADY_DELETED}")
        name = prop(item, "DLName", "") or ""
        item.Move(self._deleted_folder_obj(key))  # Moved, never Delete(): Delete inside Deleted Items is permanent.
        return {"deleted": group_id, "name": name, "note": DELETE_NOTE}
