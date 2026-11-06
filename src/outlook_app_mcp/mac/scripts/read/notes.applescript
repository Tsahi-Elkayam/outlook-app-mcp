-- Every note, plus the note folders.
-- Output records (RS); the first field says which:
--   "F", folder id, folder name, account email ("" = On My Computer), account kind
--   "N", id, title, modified (stamp), folder id, preview (first 200 characters of the text)

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook" to set fls to every note folder
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
		set ids to id of every note
		set nms to name of every note
		set mds to modification date of every note
		set fds to folder of every note
	end tell
	set n to count of ids
	if (count of nms) is not n or (count of fds) is not n then
		error "FOLDER_CHANGED: notes changed while they were being read" number 9001
	end if
	repeat with i from 1 to n
		tell application "Microsoft Outlook"
			set fid to ""
			try
				set fid to (id of (item i of fds)) as text
			end try
			set pv to ""
			try
				set pv to plain text content of (note id (item i of ids))
				if pv is missing value then set pv to ""
				if (length of pv) > 200 then set pv to text 1 thru 200 of pv
			end try
		end tell
		set end of out to my joinList({"N", (item i of ids) as text, my clean(item i of nms), my stampOrEmpty(item i of mds), fid, my clean(pv)}, FS)
	end repeat
	return my joinList(out, RS)
end run
