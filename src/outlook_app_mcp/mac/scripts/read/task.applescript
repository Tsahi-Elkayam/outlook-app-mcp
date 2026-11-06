-- One task in full.
-- argv: taskId
-- Output: the "T" record of read/tasks (FS), then reminder (stamp), then RS, then the body text.

on run argv
	my initSeps()
	set tid to (item 1 of argv) as integer
	set catOut to {}
	tell application "Microsoft Outlook"
		set t to task id tid
		set nm to name of t
		set du to due date of t
		set sd to start date of t
		set cd to completed date of t
		set fg to todo flag of t
		set pr to priority of t
		set fid to ""
		try
			set fid to (id of (folder of t)) as text
		end try
		set cats to categories of t
		repeat with c in cats
			set end of catOut to my clean(name of (contents of c))
		end repeat
		set rm to reminder date time of t
		set body to plain text content of t
	end tell
	if body is missing value then set body to ""
	return my joinList({"T", tid as text, my clean(nm), my stampOrEmpty(du), my stampOrEmpty(sd), my stampOrEmpty(cd), my enumText(fg), my enumText(pr), fid, my joinList(catOut, GS), my stampOrEmpty(rm)}, FS) & RS & body
end run
