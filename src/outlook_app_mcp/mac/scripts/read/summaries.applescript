-- Index records for the given messages, read one message at a time.
-- argv: messageId...
-- Output records (RS) as read/index gives them: message id, folder id, received,
--   is read, flag, subject, sender. Messages that no longer exist are left out.

on run argv
	my initSeps()
	set out to {}
	repeat with t in argv
		try
			tell application "Microsoft Outlook"
				set m to message id (t as integer)
				set fId to (id of (folder of m)) as text
				set tr to time received of m
				set readFlag to (is read of m) as text
				set fl to todo flag of m
				set sb to subject of m
				set s to sender of m
			end tell
			set end of out to (t as text) & FS & fId & FS & my stamp(tr) & FS & readFlag & FS & my enumText(fl) & FS & my clean(sb) & FS & my addrText(s)
		end try
	end repeat
	return my joinList(out, RS)
end run
