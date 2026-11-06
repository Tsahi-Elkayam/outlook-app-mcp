-- What the search filters need beyond the index, read a folder at a time.
-- argv: needs (letters: "t" to/cc addresses, "a" attachment count, "c" categories),
--   sinceSeconds ("" = no date limit), folderId...
-- Output records (RS): id, to and cc addresses (GS), attachment count,
--   categories (names GS); fields not asked for are "".

on run argv
	my initSeps()
	set needs to item 1 of argv
	set wantTo to (needs contains "t")
	set wantAtt to (needs contains "a")
	set wantCat to (needs contains "c")
	set sinceArg to item 2 of argv
	if sinceArg is "" then
		set d to (current date) - (200 * 365 * 86400)
	else
		set d to (current date) - (sinceArg as integer)
	end if
	set out to {}
	repeat with fidText in (items 3 thru -1 of argv)
		tell application "Microsoft Outlook"
			set f to mail folder id (fidText as integer)
			set ids to id of (every message of f whose time received > d)
			set n to count of ids
			set tos to {}
			set ccs to {}
			set ats to {}
			set cts to {}
			if wantTo then
				set tos to to recipients of (every message of f whose time received > d)
				set ccs to cc recipients of (every message of f whose time received > d)
			end if
			if wantAtt then set ats to attachments of (every message of f whose time received > d)
			if wantCat then set cts to categories of (every message of f whose time received > d)
		end tell
		if (wantTo and ((count of tos) is not n or (count of ccs) is not n)) or (wantAtt and (count of ats) is not n) or (wantCat and (count of cts) is not n) then
			error "FOLDER_CHANGED: folder " & fidText & " changed while it was being read" number 9001
		end if
		repeat with i from 1 to n
			set addrs to {}
			set catOut to {}
			set nAtt to ""
			tell application "Microsoft Outlook"
				if wantTo then
					repeat with r in ((item i of tos) & (item i of ccs))
						try
							set e to email address of r
							set end of addrs to my clean(address of e)
						end try
					end repeat
				end if
				if wantCat then
					repeat with c in (item i of cts)
						try
							set end of catOut to my clean(name of c)
						end try
					end repeat
				end if
			end tell
			if wantAtt then set nAtt to (count of (item i of ats)) as text
			set end of out to ((item i of ids) as text) & FS & my joinList(addrs, GS) & FS & nAtt & FS & my joinList(catOut, GS)
		end repeat
	end repeat
	return my joinList(out, RS)
end run
