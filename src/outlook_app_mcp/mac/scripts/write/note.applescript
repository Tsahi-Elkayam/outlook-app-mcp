-- Create, update or delete a note.
-- argv: "create", note folder id ("" = Outlook's default), key, value...
--   or: "update", noteId, key, value...
--   or: "delete", noteId   (Python checks it is in an Exchange account's notes folder,
--                           where Outlook moves deleted items to Deleted Items)
-- Keys: title, html
-- Output: note id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	set pairs to {}
	if (count of argv) > 2 then set pairs to items 3 thru -1 of argv
	tell application "Microsoft Outlook"
		if action is "create" then
			if target is "" then
				set n to make new note with properties {name:"outlook-app-mcp"}
			else
				set n to make new note at (note folder id (target as integer)) with properties {name:"outlook-app-mcp"}
			end if
		else if action is "update" then
			set n to note id (target as integer)
		else if action is "delete" then
			delete (note id (target as integer))
			return target
		else
			error "unknown action " & action number 9003
		end if
	end tell
	repeat with i from 1 to (count of pairs) by 2
		set k to item i of pairs
		set v to item (i + 1) of pairs
		tell application "Microsoft Outlook"
			if k is "title" then
				set name of n to v
			else if k is "html" then
				set content of n to v
			else
				error "unknown key " & k number 9003
			end if
		end tell
	end repeat
	tell application "Microsoft Outlook" to set nid to (id of n) as text
	return nid
end run
