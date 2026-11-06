# Changelog

## 0.3.0 (unreleased)

First public release.

- Two backends behind one set of 56 tools: Outlook for Mac in Legacy mode
  (AppleScript) and classic Outlook for Windows (COM, experimental).
- Mail: list (with paging), search (text plus sender, recipient, attachment,
  flag, category and date filters), read (headers, HTML), whole conversations,
  mark (flag with a due date), move, delete (to Deleted Items only), drafts and
  editing them, send, reply, forward (cc, attachments, signature, importance),
  folders (create, rename, move, delete to Deleted Items), attachments (one or
  all), signatures, send/receive.
- Calendar: calendars, events with recurring series expanded, single
  occurrences (read, change), create with recurrence, attendees and
  invitations, update with attendees and updates, delete (to Deleted Items;
  meetings refused), cancel meetings you organize, respond to invitations
  (also from the invitation message), free/busy.
- Tasks and notes: list, read, create, update, complete, delete (to Deleted
  Items).
- Contacts: search, read every field, create, update, delete (to Deleted
  Items); contact groups; the organization directory.
- Accounts, categories (list, create, assign), automatic replies (macOS:
  read and set; Windows: whether they are on).
- Safety: nothing deletes permanently; nothing is sent unless the tool's
  purpose is sending and the caller asked; replies and forwards are drafts
  unless asked; recurring series change only with `whole_series=true`;
  attachments only from and to allowed folders; mail content is marked as
  untrusted; tool annotations mark read, destructive and sending tools.
- `--version`, `OUTLOOK_APP_MCP_DEBUG=1` logging.
