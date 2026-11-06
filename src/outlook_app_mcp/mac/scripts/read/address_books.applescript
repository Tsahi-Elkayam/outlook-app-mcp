-- Address books and the account each belongs to.
-- Output records (RS): id, account email ("" = On My Computer or Google), account kind

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook" to set abs to every address book
	repeat with ab in abs
		set aEmail to ""
		set aKind to ""
		tell application "Microsoft Outlook"
			set abId to (id of ab) as text
			try
				set a to account of ab
				set aEmail to email address of a
				set aKind to my accountKind(a)
			end try
		end tell
		set end of out to abId & FS & my clean(aEmail) & FS & aKind
	end repeat
	return my joinList(out, RS)
end run
