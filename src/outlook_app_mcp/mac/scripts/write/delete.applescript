-- Delete messages. Outlook moves them to the account's Deleted Items folder;
-- a message already in a deleted-items folder would be deleted permanently, so
-- every message is checked first and nothing is deleted if one fails.
-- argv: protected folder ids (GS-joined), messageId...
-- Output: number of messages deleted.

on run argv
	my initSeps()
	set protected to my splitText(item 1 of argv, GS)
	set msgs to {}
	repeat with t in (items 2 thru -1 of argv)
		tell application "Microsoft Outlook"
			set m to message id (t as integer)
			set fId to (id of (folder of m)) as text
		end tell
		if protected contains fId then error "ALREADY_DELETED: message " & t & " is already in a deleted-items folder" number 9005
		set end of msgs to m
	end repeat
	repeat with m in msgs
		tell application "Microsoft Outlook"
			delete m
		end tell
	end repeat
	return (count msgs) as text
end run
