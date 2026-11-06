-- The categories of some items of one itemKind.
-- argv: itemKind (message|event|task|contact|note|group), id...
-- Output records (RS): id, categories (names GS). Items that don't exist are left out.

on run argv
	my initSeps()
	set itemKind to item 1 of argv
	set out to {}
	repeat with t in (rest of argv)
		try
			set x to my itemOf(itemKind, (t as integer))
			set catOut to {}
			tell application "Microsoft Outlook" to set cats to categories of x
			repeat with c in cats
				tell application "Microsoft Outlook" to set end of catOut to my clean(name of c)
			end repeat
			set end of out to (t as text) & FS & my joinList(catOut, GS)
		end try
	end repeat
	return my joinList(out, RS)
end run

on itemOf(itemKind, n)
	tell application "Microsoft Outlook"
		if itemKind is "message" then return message id n
		if itemKind is "event" then return calendar event id n
		if itemKind is "task" then return task id n
		if itemKind is "contact" then return contact id n
		if itemKind is "note" then return note id n
		if itemKind is "group" then return group id n
	end tell
	error "unknown itemKind " & itemKind number 9003
end itemOf
