-- Every contact group.
-- Output records (RS): id, name, members (entries GS, each name US address),
--   address book id, account email, account kind

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook" to set allGroups to every group
	repeat with g in allGroups
		set memOut to {}
		set abId to ""
		set aEmail to ""
		set aKind to ""
		tell application "Microsoft Outlook"
			set gId to (id of g) as text
			set nm to name of g
			set mems to members of g
			repeat with m in mems
				set end of memOut to my addrText(contents of m)
			end repeat
			try
				set ab to address book of g
				set abId to (id of ab) as text
				set a to account of ab
				set aEmail to email address of a
				set aKind to my accountKind(a)
			end try
		end tell
		set end of out to my joinList({gId, my clean(nm), my joinList(memOut, GS), abId, my clean(aEmail), aKind}, FS)
	end repeat
	return my joinList(out, RS)
end run
