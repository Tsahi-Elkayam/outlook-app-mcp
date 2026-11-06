-- Create, update or delete a task.
-- argv: "create", task folder id ("" = Outlook's default), key, value...
--   or: "update", taskId, key, value...
--   or: "delete", taskId   (Python checks it is in an Exchange account's task folder,
--                           where Outlook moves deleted items to Deleted Items)
-- Keys: subject, html, due / start / reminder (stamp, or "" to clear),
--   priority (low, normal, high), complete ("1"/"0"), categories (names GS; must exist).
-- Output: task id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	set pairs to {}
	if (count of argv) > 2 then set pairs to items 3 thru -1 of argv
	tell application "Microsoft Outlook"
		if action is "create" then
			if target is "" then
				set t to make new task with properties {name:"outlook-app-mcp"}
			else
				set t to make new task at (task folder id (target as integer)) with properties {name:"outlook-app-mcp"}
			end if
		else if action is "update" then
			set t to task id (target as integer)
		else if action is "delete" then
			delete (task id (target as integer))
			return target
		else
			error "unknown action " & action number 9003
		end if
	end tell
	repeat with i from 1 to (count of pairs) by 2
		set k to item i of pairs
		set v to item (i + 1) of pairs
		if k is in {"due", "start", "reminder"} and v is not "" then set dv to my parseStamp(v)
		tell application "Microsoft Outlook"
			if k is "subject" then
				set name of t to v
			else if k is "html" then
				set content of t to v
			else if k is "due" then
				if v is "" then
					set due date of t to missing value
				else
					set due date of t to dv
				end if
			else if k is "start" then
				if v is "" then
					set start date of t to missing value
				else
					set start date of t to dv
				end if
			else if k is "reminder" then
				if v is "" then
					set reminder date time of t to missing value
				else
					set reminder date time of t to dv
				end if
			else if k is "priority" then
				if v is "high" then
					set priority of t to priority high
				else if v is "low" then
					set priority of t to priority low
				else
					set priority of t to priority normal
				end if
			else if k is "complete" then
				if v is "1" then
					set todo flag of t to completed
				else
					set todo flag of t to not completed
				end if
			else if k is "categories" then
				set wanted to {}
				repeat with nm in my splitText(v, GS)
					set nm to nm as text
					if nm is not "" then set end of wanted to (first category whose name is nm)
				end repeat
				set categories of t to wanted
			else
				error "unknown key " & k number 9003
			end if
		end tell
	end repeat
	tell application "Microsoft Outlook" to set tid to (id of t) as text
	return tid
end run
