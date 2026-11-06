-- Every calendar.
-- Output records (RS) of fields (FS): calendar id, account email, name, is default,
--   exchange id (Google calendars have no account; the caller matches this to a mailbox)

on run argv
	my initSeps()
	set defId to ""
	tell application "Microsoft Outlook"
		try
			set defId to (id of (default calendar of default account)) as text
		end try
		set cals to every calendar
	end tell
	set out to {}
	repeat with c in cals
		set aEmail to ""
		set xId to ""
		tell application "Microsoft Outlook"
			set cId to (id of c) as text
			set cName to name of c
			try
				set aEmail to email address of (account of c)
			end try
			try
				set xId to exchange id of c
			end try
		end tell
		set end of out to cId & FS & my clean(aEmail) & FS & my clean(cName) & FS & ((cId is defId) as text) & FS & my clean(xId)
	end repeat
	return my joinList(out, RS)
end run
