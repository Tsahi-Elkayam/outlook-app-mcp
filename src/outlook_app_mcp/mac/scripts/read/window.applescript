-- Summaries of the messages received after a cutoff, in one or more folders.
-- Same output as read/index, but fast on large folders: Outlook lists messages
-- newest first, so a binary search finds where the cutoff falls, the messages
-- above it are read in bulk as one range, and one count proves the range holds
-- every message after the cutoff. A folder whose proof fails is read with the
-- slow filtered scan instead, so the result never depends on the order.
-- argv: sinceSeconds (required), folderId...
-- Output records (RS) of fields (FS):
--   message id, folder id, received (stamp), is read, flag, subject, sender

on run argv
	my initSeps()
	set cutoff to (current date) - ((item 1 of argv) as integer)
	set out to {}
	repeat with fidText in (items 2 thru -1 of argv)
		set fid to fidText as integer
		set k to my rangeEnd(fid, cutoff)
		if k is missing value then
			my slowScan(fid, cutoff, out)
		else if k > 0 then
			tell application "Microsoft Outlook"
				set f to mail folder id fid
				set ids to id of messages 1 thru k of f
				set tms to time received of messages 1 thru k of f
				set rds to is read of messages 1 thru k of f
				set fls to todo flag of messages 1 thru k of f
				set sbs to subject of messages 1 thru k of f
				set snd to sender of messages 1 thru k of f
			end tell
			if (count of ids) is not k or (count of tms) is not k or (count of rds) is not k or (count of fls) is not k or (count of sbs) is not k or (count of snd) is not k then
				error "FOLDER_CHANGED: folder " & fid & " changed while it was being read" number 9001
			end if
			repeat with i from 1 to k
				set end of out to ((item i of ids) as text) & FS & (fid as text) & FS & my stamp(item i of tms) & FS & ((item i of rds) as text) & FS & my enumText(item i of fls) & FS & my clean(item i of sbs) & FS & my addrText(item i of snd)
			end repeat
		end if
	end repeat
	return my joinList(out, RS)
end run

-- How many messages at the top of the folder were received after the cutoff,
-- or missing value if the folder is not in newest-first order around it.
on rangeEnd(fid, cutoff)
	tell application "Microsoft Outlook"
		set f to mail folder id fid
		set n to count messages of f
		if n is 0 then return 0
		-- every message 1..lo is after the cutoff; every message hi+1..n is not
		set lo to 0
		set hi to n
		repeat while lo < hi
			set mid to (lo + hi + 1) div 2
			set t to time received of message mid of f
			if t is not missing value and t > cutoff then
				set lo to mid
			else
				set hi to mid - 1
			end if
		end repeat
		set inFolder to count (every message of f whose time received > cutoff)
	end tell
	if inFolder is not lo then return missing value
	-- The count matches; make sure the block itself is all after the cutoff.
	if lo > 0 then
		tell application "Microsoft Outlook"
			set tms to time received of messages 1 thru lo of f
		end tell
		repeat with t in tms
			set t to contents of t
			if t is missing value then return missing value
			if t ≤ cutoff then return missing value
		end repeat
	end if
	return lo
end rangeEnd

on slowScan(fid, d, out)
	tell application "Microsoft Outlook"
		set f to mail folder id fid
		set ids to id of (every message of f whose time received > d)
		set tms to time received of (every message of f whose time received > d)
		set rds to is read of (every message of f whose time received > d)
		set fls to todo flag of (every message of f whose time received > d)
		set sbs to subject of (every message of f whose time received > d)
		set snd to sender of (every message of f whose time received > d)
	end tell
	set n to count of ids
	if (count of tms) is not n or (count of rds) is not n or (count of fls) is not n or (count of sbs) is not n or (count of snd) is not n then
		error "FOLDER_CHANGED: folder " & fid & " changed while it was being read" number 9001
	end if
	repeat with i from 1 to n
		set end of out to ((item i of ids) as text) & FS & (fid as text) & FS & my stamp(item i of tms) & FS & ((item i of rds) as text) & FS & my enumText(item i of fls) & FS & my clean(item i of sbs) & FS & my addrText(item i of snd)
	end repeat
end slowScan
