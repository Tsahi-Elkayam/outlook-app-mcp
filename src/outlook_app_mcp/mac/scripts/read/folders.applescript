-- Every mail folder of every account.
-- argv: ids of Google account roots whose address the caller already knows
--       (their mail is not sampled again).
-- Output, two sections joined by GS:
--   1. folder records (RS) of fields (FS):
--        folder id, account email, account name, path, unread count, message count,
--        role (inbox, sent, drafts, deleted, junk, outbox, all, or empty), kind
--      kind is exchange, imap, pop or google. For Google accounts the account email
--      is "google:<root folder id>"; the caller resolves it from section 2.
--   2. Google root records (RS) of fields (FS):
--        root folder id, senders of recent sent mail (US), recipients of recent
--        inbox mail (US), exchange id of one of its folders (names the mailbox,
--        which ties the account's calendars to it)
--      Roots listed in argv come back with empty samples.
--
-- Roles of Exchange/IMAP/POP folders come from the account's own properties, so
-- they work whatever the UI language. Google accounts (added through the
-- Microsoft cloud) have no account object in Outlook's dictionary: their root
-- folder has no name, account or container, and their roles come from Gmail's
-- folder names.

property SAMPLE : 30

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook"
		set exA to every exchange account
		set imA to every imap account
		set poA to every pop account
	end tell
	repeat with pair in {{exA, "exchange"}, {imA, "imap"}, {poA, "pop"}}
		repeat with a in (item 1 of pair)
			tell application "Microsoft Outlook"
				set aEmail to my clean(email address of a)
				set aName to my clean(name of a)
				set tops to mail folders of a
			end tell
			set roles to {}
			repeat with role in {"inbox", "sent", "drafts", "deleted", "junk", "outbox"}
				my addRole(a, role as text, roles)
			end repeat
			my walk(tops, "", aEmail, aName, roles, item 2 of pair, out)
		end repeat
	end repeat

	set known to {}
	repeat with t in argv
		set end of known to (t as text)
	end repeat
	set gOut to {}
	repeat with rId in my googleRoots()
		set rText to rId as text
		tell application "Microsoft Outlook"
			set tops to mail folders of mail folder id rId
		end tell
		set roles to my nameRoles(tops)
		my walk(tops, "", "google:" & rText, "", roles, "google", out)
		set sentL to {}
		set inboxL to {}
		if known does not contain rText then
			set sentL to my senders(my roleFolder(tops, roles, "sent"))
			set inboxL to my recipients(my roleFolder(tops, roles, "inbox"))
		end if
		set end of gOut to rText & FS & my joinList(sentL, US) & FS & my joinList(inboxL, US) & FS & my mailboxId(tops)
	end repeat
	return my joinList(out, RS) & GS & my joinList(gOut, RS)
end run

on mailboxId(tops)
	repeat with f in tops
		tell application "Microsoft Outlook"
			set x to exchange id of f
		end tell
		if x is not missing value then return my clean(x)
	end repeat
	return ""
end mailboxId

-- Root folders of Google accounts: no name, no account, no container.
on googleRoots()
	set found to {}
	tell application "Microsoft Outlook"
		set ids to id of every mail folder
		set nms to name of every mail folder
		repeat with i from 1 to count ids
			if item i of nms is missing value then
				set f to mail folder id (item i of ids)
				if (account of f) is missing value and (container of f) is missing value then set end of found to item i of ids
			end if
		end repeat
	end tell
	return found
end googleRoots

on nameRoles(tops)
	set roles to {}
	set table to {{"Inbox", "inbox"}, {"Sent Mail", "sent"}, {"Drafts", "drafts"}, {"Trash", "deleted"}, {"Spam", "junk"}, {"Outbox", "outbox"}, {"All Mail", "all"}}
	repeat with f in tops
		tell application "Microsoft Outlook"
			set fName to name of f
			set fId to (id of f) as text
		end tell
		repeat with r in table
			if fName is (item 1 of r) then set end of roles to {fId, item 2 of r}
		end repeat
	end repeat
	return roles
end nameRoles

on roleFolder(tops, roles, role)
	repeat with r in roles
		if item 2 of r is role then
			tell application "Microsoft Outlook" to return mail folder id ((item 1 of r) as integer)
		end if
	end repeat
	return missing value
end roleFolder

on senders(f)
	set out to {}
	if f is missing value then return out
	tell application "Microsoft Outlook"
		set n to count messages of f
		if n > SAMPLE then set n to SAMPLE
		repeat with i from 1 to n
			try
				set s to sender of (message i of f)
				set end of out to my clean(address of s)
			end try
		end repeat
	end tell
	return out
end senders

on recipients(f)
	set out to {}
	if f is missing value then return out
	tell application "Microsoft Outlook"
		set n to count messages of f
		if n > SAMPLE then set n to SAMPLE
		repeat with i from 1 to n
			set m to message i of f
			set rl to (to recipients of m) & (cc recipients of m)
			repeat with x in rl
				try
					set e to email address of x
					set end of out to my clean(address of e)
				end try
			end repeat
		end repeat
	end tell
	return out
end recipients

on addRole(a, role, roles)
	try
		tell application "Microsoft Outlook"
			if role is "inbox" then
				set f to inbox of a
			else if role is "sent" then
				set f to sent items of a
			else if role is "drafts" then
				set f to drafts of a
			else if role is "deleted" then
				set f to deleted items of a
			else if role is "junk" then
				set f to junk mail of a
			else
				set f to outbox of a
			end if
			set end of roles to {(id of f) as text, role}
		end tell
	end try
end addRole

on roleOf(fId, roles)
	repeat with r in roles
		if item 1 of r is fId then return item 2 of r
	end repeat
	return ""
end roleOf

on walk(fl, prefix, aEmail, aName, roles, kind, out)
	repeat with f in fl
		tell application "Microsoft Outlook"
			set fName to my clean(name of f)
			set fId to (id of f) as text
			set nUnread to (unread count of f) as text
			set nAll to (count messages of f) as text
			set subs to mail folders of f
		end tell
		set p to prefix & fName
		set end of out to fId & FS & aEmail & FS & aName & FS & p & FS & nUnread & FS & nAll & FS & my roleOf(fId, roles) & FS & kind
		my walk(subs, p & "/", aEmail, aName, roles, kind, out)
	end repeat
end walk
