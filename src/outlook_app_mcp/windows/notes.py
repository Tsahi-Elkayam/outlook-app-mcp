"""Windows notes: Outlook's Notes folders. A note is one text; its title is the first line."""

from __future__ import annotations

import re

from ..common import (
    ALREADY_DELETED,
    DEFAULT_LIMIT,
    DELETE_NOTE,
    OutlookError,
    check_title,
    clamp_limit,
)
from .com import (
    FOLDER_NOTES,
    ITEM_NOTE,
    categories,
    com_errors,
    com_iso,
    each,
    make_id,
    prop,
)

PREVIEW_CHARS = 200


def is_note_class(message_class: str) -> bool:
    return (message_class or "").startswith("IPM.StickyNote")


def split_note(body: str) -> tuple[str, str]:
    """(title, rest): a note's title is its first line; Outlook separates lines with CR LF."""
    first, _, rest = re.sub(r"\r\n?", "\n", body or "").partition("\n")
    return first.strip(), rest


def join_note(title: str, rest: str) -> str:
    return title + ("\r\n" + rest.replace("\n", "\r\n") if rest else "")


class WinNotesMixin:
    """Needs the store helpers of WinMailMixin and _kind_folders / _default_folder of WinMiscMixin."""

    def _note_item(self, note_id: str, what: str = "note_id"):
        it, key = self._item(note_id, what)
        if not is_note_class(prop(it, "MessageClass", "") or ""):
            raise OutlookError(f"{note_id} is not a note ({prop(it, 'MessageClass', '?')}).")
        return it, key

    @com_errors
    def list_notes(self, query: str | None = None, limit: int = DEFAULT_LIMIT) -> dict:
        """Notes from every Notes folder, most recently changed first. query matches title and text."""
        limit = clamp_limit(limit)
        needle = (query or "").strip().casefold()
        found = []
        for store, folder in self._kind_folders(ITEM_NOTE):
            for it in each(folder.Items):
                if not is_note_class(prop(it, "MessageClass", "") or ""):
                    continue
                body = prop(it, "Body", "") or ""
                if needle and needle not in body.casefold():
                    continue
                title, rest = split_note(body)
                found.append({
                    "id": make_id(store.key, it.EntryID),
                    "title": title,
                    "modified": com_iso(prop(it, "LastModificationTime")),
                    "preview": " ".join(rest.split())[:PREVIEW_CHARS],
                    "folder": prop(folder, "Name", "") or "",
                    "account": store.account,
                })
        found.sort(key=lambda n: n["modified"] or "", reverse=True)
        return {"count": len(found), "returned": min(limit, len(found)), "has_more": len(found) > limit,
                "notes": found[:limit]}

    def _note_detail(self, it, key: str) -> dict:
        body = prop(it, "Body", "") or ""
        return {
            "id": make_id(key, it.EntryID),
            "title": split_note(body)[0],
            "body": split_note(body)[1],
            "modified": com_iso(prop(it, "LastModificationTime")),
            "categories": categories(it),
        }

    @com_errors
    def read_note(self, note_id: str) -> dict:
        self.require_accounts()
        it, key = self._note_item(note_id)
        return self._note_detail(it, key)

    @staticmethod
    def _note_title(title: str) -> str:
        title = check_title(title, "title")
        if re.search(r"[\r\n]", title):
            raise OutlookError("title must be a single line; put the rest in body.")
        return title

    @com_errors
    def create_note(self, title: str, body: str = "", account: str | None = None) -> dict:
        title = self._note_title(title)
        folder, key = self._default_folder(FOLDER_NOTES, account, "Notes")
        it = folder.Items.Add(ITEM_NOTE)
        it.Body = join_note(title, body or "")
        it.Save()
        return self._note_detail(it, key)

    @com_errors
    def update_note(self, note_id: str, title: str | None = None, body: str | None = None) -> dict:
        """Change the title (the first line) and/or the text below it. Outlook derives the title from the text,
        so both are written back as one body."""
        if title is None and body is None:
            raise OutlookError("Nothing to change.")
        new_title = self._note_title(title) if title is not None else None
        it, key = self._note_item(note_id)
        if self._under_deleted(it, key):
            raise OutlookError("That note is in Deleted Items; restore it in Outlook before changing it.")
        old_title, old_rest = split_note(prop(it, "Body", "") or "")
        it.Body = join_note(new_title if new_title is not None else old_title, body if body is not None else old_rest)
        it.Save()
        return self._note_detail(it, key)

    @com_errors
    def delete_note(self, note_id: str) -> dict:
        it, key = self._note_item(note_id)
        if self._under_deleted(it, key):
            raise OutlookError(f"Note {note_id} {ALREADY_DELETED}")
        title = split_note(prop(it, "Body", "") or "")[0]
        it.Move(self._deleted_folder_obj(key))  # Moved, never Delete(): Delete inside Deleted Items is permanent.
        return {"deleted": note_id, "title": title, "note": DELETE_NOTE}
