-- Cancel a meeting you organize: Outlook sends the cancellation to the attendees.
-- argv: eventId, occurrence start (stamp, or "" for the whole event or series),
--   comment (HTML; "" = none; becomes the body of the cancellation)
-- Output: event id.

on run argv
	my initSeps()
	set eid to (item 1 of argv) as integer
	set occS to item 2 of argv
	set cmt to item 3 of argv
	tell application "Microsoft Outlook"
		set e to calendar event id eid
		if occS is not "" then
			set e to get occurrence of e at my parseStamp(occS)
			if e is missing value then error "NO_OCCURRENCE: the series has no occurrence at that time" number 9011
		end if
		if cmt is not "" then set content of e to cmt
		with timeout of 590 seconds -- Outlook asks the user to confirm a scripted send
			cancel meeting e
		end timeout
	end tell
	return eid as text
end run
