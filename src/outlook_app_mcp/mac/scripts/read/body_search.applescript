-- Ids of messages whose body text contains a query.
-- argv: query, sinceSeconds ("" = no date limit), folderId...
-- Output: message ids joined by RS.
-- The query is only used as a value inside a "whose" filter.

on run argv
	my initSeps()
	set q to item 1 of argv
	set sinceArg to item 2 of argv
	set folderIds to items 3 thru -1 of argv
	if sinceArg is "" then
		set d to missing value
	else
		set d to (current date) - (sinceArg as integer)
	end if
	set out to {}
	repeat with fidText in folderIds
		tell application "Microsoft Outlook"
			set f to mail folder id (fidText as integer)
			if d is missing value then
				set ids to id of (every message of f whose plain text content contains q)
			else
				set ids to id of (every message of f whose time received > d and plain text content contains q)
			end if
		end tell
		repeat with i in ids
			set end of out to (i as text)
		end repeat
	end repeat
	return my joinList(out, RS)
end run
