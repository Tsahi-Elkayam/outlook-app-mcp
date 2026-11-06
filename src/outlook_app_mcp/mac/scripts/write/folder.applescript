-- Create or rename a mail folder.
-- argv: "create", parentFolderId ("" = top level), account email (used when
--   parent is ""), name
--   or: "rename", folderId, "", new name
--   or: "move", folderId, destination folder id, ""  (delete = move into Deleted Items;
--       Python picks the destination and refuses special folders)
-- Output: folder id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	set acct to item 3 of argv
	set nm to item 4 of argv
	if action is "create" then
		if target is "" then
			set a to my findAccount(acct)
			tell application "Microsoft Outlook"
				set f to make new mail folder at a with properties {name:nm}
			end tell
		else
			tell application "Microsoft Outlook"
				set f to make new mail folder at (mail folder id (target as integer)) with properties {name:nm}
			end tell
		end if
	else if action is "rename" then
		tell application "Microsoft Outlook"
			set f to mail folder id (target as integer)
			set name of f to nm
		end tell
	else if action is "move" then
		tell application "Microsoft Outlook"
			set f to mail folder id (target as integer)
			move f to (mail folder id (acct as integer))
		end tell
	else
		error "unknown action " & action number 9003
	end if
	tell application "Microsoft Outlook" to set fid to (id of f) as text
	return fid
end run
