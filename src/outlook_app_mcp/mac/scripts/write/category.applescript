-- Categories: create one, or give items their categories.
-- argv: "create", name
--   or: "assign", itemKind (message|event|task|contact|note), id, names (GS), id, names, ...
--       Every name must already exist in Outlook's category list.
-- Output: "create" -> the name; "assign" -> number of items changed.

on run argv
	my initSeps()
	set action to item 1 of argv
	if action is "create" then
		set nm to item 2 of argv
		tell application "Microsoft Outlook"
			if (count of (every category whose name is nm)) is 0 then make new category with properties {name:nm}
		end tell
		return nm
	else if action is "assign" then
		set itemKind to item 2 of argv
		set n to 0
		repeat with i from 3 to (count of argv) by 2
			set x to my itemOf(itemKind, ((item i of argv) as integer))
			set wanted to {}
			repeat with nm in my splitText(item (i + 1) of argv, GS)
				set nm to nm as text
				if nm is not "" then
					tell application "Microsoft Outlook" to set end of wanted to (first category whose name is nm)
				end if
			end repeat
			tell application "Microsoft Outlook" to set categories of x to wanted
			set n to n + 1
		end repeat
		return n as text
	end if
	error "unknown action " & action number 9003
end run

on itemOf(itemKind, n)
	tell application "Microsoft Outlook"
		if itemKind is "message" then return message id n
		if itemKind is "event" then return calendar event id n
		if itemKind is "task" then return task id n
		if itemKind is "contact" then return contact id n
		if itemKind is "note" then return note id n
	end tell
	error "unknown itemKind " & itemKind number 9003
end itemOf
