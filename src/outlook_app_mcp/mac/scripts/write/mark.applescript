-- Change read state or flag of messages.
-- argv: action (read, unread, flag, unflag, complete), due date stamp ("" = none; flag only), messageId...
-- Output: number of messages changed.

on run argv
	my initSeps()
	set action to item 1 of argv
	set dueS to item 2 of argv
	set n to 0
	repeat with t in (items 3 thru -1 of argv)
		tell application "Microsoft Outlook"
			set m to message id (t as integer)
			if action is "read" then
				set is read of m to true
			else if action is "unread" then
				set is read of m to false
			else if action is "flag" then
				set todo flag of m to «constant ****FlNC»
				if dueS is not "" then set due date of m to my parseStamp(dueS)
			else if action is "unflag" then
				set todo flag of m to «constant ****FlNF»
			else if action is "complete" then
				set todo flag of m to «constant ****FlCo»
			else
				error "unknown action " & action number 9003
			end if
		end tell
		set n to n + 1
	end repeat
	return n as text
end run
