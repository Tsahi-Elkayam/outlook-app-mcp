-- Create, update or delete a contact group.
-- argv: "create", address book id ("" = Outlook's default), name, members
--   or: "update", groupId, name ("" = keep), members
--   or: "delete", groupId  (Python checks it is in an Exchange account's address book)
-- members: the whole new list, entries GS, each name US address; "-" = keep the list.
-- Output: group id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	tell application "Microsoft Outlook"
		if action is "create" then
			if target is "" then
				set g to make new group with properties {name:(item 3 of argv)}
			else
				set g to make new group at (address book id (target as integer)) with properties {name:(item 3 of argv)}
			end if
		else if action is "update" then
			set g to group id (target as integer)
			if (item 3 of argv) is not "" then set name of g to (item 3 of argv)
		else if action is "delete" then
			delete (group id (target as integer))
			return target
		else
			error "unknown action " & action number 9003
		end if
	end tell
	set memArg to item 4 of argv
	if memArg is not "-" then
		set recs to {}
		set entries to my splitText(memArg, GS)
		-- Built inside the tell block, so "name" and "address" are Outlook's record keys.
		tell application "Microsoft Outlook"
			repeat with e in entries
				set e to e as text
				if e is not "" then
					set p to my splitText(e, US)
					set end of recs to {name:(item 1 of p), address:(item 2 of p)}
				end if
			end repeat
			set members of g to recs
		end tell
	end if
	tell application "Microsoft Outlook" to set gid to (id of g) as text
	return gid
end run
