-- Summaries of the messages in one or more mail folders.
-- argv: sinceSeconds ("" = no date limit), unreadOnly ("1"/"0"), folderId...
-- Output records (RS) of fields (FS):
--   message id, folder id, received (stamp), is read, flag, subject,
--   sender name, sender address

on run argv
	my initSeps()
	set sinceArg to item 1 of argv
	set unreadOnly to ((item 2 of argv) is "1")
	set folderIds to items 3 thru -1 of argv
	if sinceArg is "" then
		set d to missing value
	else
		set d to (current date) - (sinceArg as integer)
	end if
	set out to {}
	repeat with fidText in folderIds
		set fid to fidText as integer
		tell application "Microsoft Outlook"
			set f to mail folder id fid
			if d is missing value and not unreadOnly then
				set ids to id of every message of f
				set tms to time received of every message of f
				set rds to is read of every message of f
				set fls to todo flag of every message of f
				set sbs to subject of every message of f
				set snd to sender of every message of f
			else if d is missing value then
				set ids to id of (every message of f whose is read is false)
				set tms to time received of (every message of f whose is read is false)
				set rds to is read of (every message of f whose is read is false)
				set fls to todo flag of (every message of f whose is read is false)
				set sbs to subject of (every message of f whose is read is false)
				set snd to sender of (every message of f whose is read is false)
			else if not unreadOnly then
				set ids to id of (every message of f whose time received > d)
				set tms to time received of (every message of f whose time received > d)
				set rds to is read of (every message of f whose time received > d)
				set fls to todo flag of (every message of f whose time received > d)
				set sbs to subject of (every message of f whose time received > d)
				set snd to sender of (every message of f whose time received > d)
			else
				set ids to id of (every message of f whose time received > d and is read is false)
				set tms to time received of (every message of f whose time received > d and is read is false)
				set rds to is read of (every message of f whose time received > d and is read is false)
				set fls to todo flag of (every message of f whose time received > d and is read is false)
				set sbs to subject of (every message of f whose time received > d and is read is false)
				set snd to sender of (every message of f whose time received > d and is read is false)
			end if
		end tell
		set n to count of ids
		if (count of tms) is not n or (count of rds) is not n or (count of fls) is not n or (count of sbs) is not n or (count of snd) is not n then
			error "FOLDER_CHANGED: folder " & fid & " changed while it was being read" number 9001
		end if
		repeat with i from 1 to n
			set end of out to ((item i of ids) as text) & FS & (fid as text) & FS & my stamp(item i of tms) & FS & ((item i of rds) as text) & FS & my enumText(item i of fls) & FS & my clean(item i of sbs) & FS & my addrText(item i of snd)
		end repeat
	end repeat
	return my joinList(out, RS)
end run
