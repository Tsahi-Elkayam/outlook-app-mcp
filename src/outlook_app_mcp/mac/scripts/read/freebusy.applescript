-- Free/busy of people, through an Exchange account.
-- argv: account email, start stamp, end stamp, interval minutes, address...
-- Output: Outlook's answer, items joined by GS (per person: address, start, then for
--   each interval subject, location, status; statuses as enum text).

on run argv
	my initSeps()
	set a to my findAccount(item 1 of argv)
	set d1 to my parseStamp(item 2 of argv)
	set d2 to my parseStamp(item 3 of argv)
	set iv to (item 4 of argv) as integer
	set people to {}
	repeat with p in (items 5 thru -1 of argv)
		set end of people to p as text
	end repeat
	tell application "Microsoft Outlook"
		if class of a is not exchange account then error "NOT_EXCHANGE: free/busy is only available through an Exchange account" number 9010
		set res to query freebusy a for attendees people range start time d1 range end time d2 interval iv
	end tell
	set out to {}
	repeat with x in res
		set x to contents of x
		if class of x is text then
			set end of out to my clean(x)
		else
			set end of out to my enumText(x)
		end if
	end repeat
	return my joinList(out, GS)
end run
