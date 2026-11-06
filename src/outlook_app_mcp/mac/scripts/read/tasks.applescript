-- Every task, plus the task folders.
-- Output records (RS); the first field says which:
--   "F", folder id, folder name, account email ("" = On My Computer), account kind
--   "T", id, subject, due (stamp), start (stamp), completed (stamp), todo flag,
--        priority, folder id, categories (names GS)

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook" to set fls to every task folder
	repeat with f in fls
		set aEmail to ""
		set aKind to ""
		tell application "Microsoft Outlook"
			set fId to (id of f) as text
			set fName to name of f
			try
				set a to account of f
				set aEmail to email address of a
				set aKind to my accountKind(a)
			end try
		end tell
		set end of out to my joinList({"F", fId, my clean(fName), my clean(aEmail), aKind}, FS)
	end repeat
	tell application "Microsoft Outlook"
		set ids to id of every task
		set nms to name of every task
		set dues to due date of every task
		set sts to start date of every task
		set cps to completed date of every task
		set fgs to todo flag of every task
		set prs to priority of every task
		set fds to folder of every task
		set cts to categories of every task
	end tell
	set n to count of ids
	if (count of nms) is not n or (count of fds) is not n or (count of cts) is not n then
		error "FOLDER_CHANGED: tasks changed while they were being read" number 9001
	end if
	repeat with i from 1 to n
		set catOut to {}
		tell application "Microsoft Outlook"
			set fid to ""
			try
				set fid to (id of (item i of fds)) as text
			end try
			repeat with c in (item i of cts)
				set end of catOut to my clean(name of c)
			end repeat
		end tell
		set end of out to my joinList({"T", (item i of ids) as text, my clean(item i of nms), my stampOrEmpty(item i of dues), my stampOrEmpty(item i of sts), my stampOrEmpty(item i of cps), my enumText(item i of fgs), my enumText(item i of prs), fid, my joinList(catOut, GS)}, FS)
	end repeat
	return my joinList(out, RS)
end run
