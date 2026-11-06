-- Id and received time of the messages in one folder that match a filter.
-- Two scans of the folder, whatever its size; the caller picks which messages to
-- read in full (read/summaries).
-- argv: folderId, sinceSeconds ("" = no date limit), unreadOnly ("1"/"0")
-- Output records (RS) of fields (FS): message id, received (stamp)

on run argv
	my initSeps()
	set fid to (item 1 of argv) as integer
	set sinceArg to item 2 of argv
	set unreadOnly to ((item 3 of argv) is "1")
	if sinceArg is "" then
		set d to missing value
	else
		set d to (current date) - (sinceArg as integer)
	end if
	tell application "Microsoft Outlook"
		set f to mail folder id fid
		if d is missing value and not unreadOnly then
			set ids to id of every message of f
			set tms to time received of every message of f
		else if d is missing value then
			set ids to id of (every message of f whose is read is false)
			set tms to time received of (every message of f whose is read is false)
		else if not unreadOnly then
			set ids to id of (every message of f whose time received > d)
			set tms to time received of (every message of f whose time received > d)
		else
			set ids to id of (every message of f whose time received > d and is read is false)
			set tms to time received of (every message of f whose time received > d and is read is false)
		end if
	end tell
	if (count of ids) is not (count of tms) then
		error "FOLDER_CHANGED: folder " & fid & " changed while it was being read" number 9001
	end if
	set out to {}
	repeat with i from 1 to count of ids
		set end of out to ((item i of ids) as text) & FS & my stamp(item i of tms)
	end repeat
	return my joinList(out, RS)
end run
