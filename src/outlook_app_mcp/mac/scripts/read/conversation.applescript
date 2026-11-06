-- Messages of one conversation in the given folders.
-- argv: conversation id, folderId...
-- Output records (RS) as read/index gives them: message id, folder id, received,
--   is read, flag, subject, sender.

on run argv
	my initSeps()
	set conv to (item 1 of argv) as integer
	set out to {}
	repeat with t in (items 2 thru -1 of argv)
		tell application "Microsoft Outlook"
			set fld to mail folder id (t as integer)
			set ms to (every message of fld whose conversation id is conv)
		end tell
		repeat with m in ms
			tell application "Microsoft Outlook"
				set mid to (id of m) as text
				set tr to time received of m
				set readFlag to (is read of m) as text
				set fl to todo flag of m
				set sb to subject of m
				set s to sender of m
			end tell
			set end of out to mid & FS & (t as text) & FS & my stamp(tr) & FS & readFlag & FS & my enumText(fl) & FS & my clean(sb) & FS & my addrText(s)
		end repeat
	end repeat
	return my joinList(out, RS)
end run
