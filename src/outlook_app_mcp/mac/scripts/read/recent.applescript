-- The newest messages of one folder, fast on folders of any size.
-- Outlook lists a folder's messages newest first, so this reads the first k in
-- bulk, then proves the block is complete: the count of messages in the whole
-- folder received from one second after the block's oldest time must equal the
-- number of block messages in that range. Outlook keeps fractions of a second but
-- hands AppleScript whole seconds, so messages within the oldest second are
-- treated as ties; any of them is a correct last item.
-- If the counts differ, the caller falls back to a full scan, so the result
-- never depends on the order being right.
-- argv: folderId, k, sinceSeconds ("" = no date limit)
-- Output: "OK" or "UNSORTED", then RS, then index records as read/index gives them
--   (message id, folder id, received, is read, flag, subject, sender).
--   With sinceSeconds, only messages received after the cutoff are returned.

on run argv
	my initSeps()
	set fid to (item 1 of argv) as integer
	set k to (item 2 of argv) as integer
	set sinceArg to item 3 of argv
	set cutoff to missing value
	if sinceArg is not "" then set cutoff to (current date) - (sinceArg as integer)
	tell application "Microsoft Outlook"
		set f to mail folder id fid
		set n to count messages of f
	end tell
	if k > n then set k to n
	if k is 0 then return "OK"
	tell application "Microsoft Outlook"
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

	set oldest to missing value
	repeat with t in tms
		set t to contents of t
		if t is not missing value then
			if oldest is missing value or t < oldest then set oldest to t
		end if
	end repeat
	set status to "OK"
	if oldest is not missing value then
		set boundary to oldest + 1
		set inBlock to 0
		repeat with t in tms
			set t to contents of t
			if t is not missing value then
				if t ≥ boundary then set inBlock to inBlock + 1
			end if
		end repeat
		tell application "Microsoft Outlook"
			set inFolder to count (every message of f whose time received ≥ boundary)
		end tell
		if inFolder is not inBlock then set status to "UNSORTED"
	end if

	set out to {status}
	repeat with i from 1 to k
		set t to item i of tms
		set keep to true
		if cutoff is not missing value then
			if t is missing value then
				set keep to false
			else if t ≤ cutoff then
				set keep to false
			end if
		end if
		if keep then
			set end of out to ((item i of ids) as text) & FS & (fid as text) & FS & my stamp(t) & FS & ((item i of rds) as text) & FS & my enumText(item i of fls) & FS & my clean(item i of sbs) & FS & my addrText(item i of snd)
		end if
	end repeat
	return my joinList(out, RS)
end run
