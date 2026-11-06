-- Answer a meeting invitation.
-- argv: itemKind ("event" or "message"), id, occurrence start (stamp; "" = the event itself),
--   response (accept, tentative, decline), comment,
--   send response ("1" = tell the organizer, "0" = don't)
-- "message" answers a meeting request message in the inbox.
-- Output: the id.

on run argv
	my initSeps()
	set itemKind to item 1 of argv
	set n to (item 2 of argv) as integer
	set occS to item 3 of argv
	set resp to item 4 of argv
	set cmt to item 5 of argv
	set tell_ to ((item 6 of argv) is "1")
	tell application "Microsoft Outlook"
		with timeout of 590 seconds -- Outlook may ask the user to confirm sending the response
		if itemKind is "message" then
			set m to message id n
			if resp is "accept" then
				accept invite m comment cmt sending response tell_
			else if resp is "tentative" then
				accept tentatively invite m comment cmt sending response tell_
			else if resp is "decline" then
				decline invite m comment cmt sending response tell_
			else
				error "unknown response " & resp number 9003
			end if
		else
			set e to calendar event id n
			if occS is not "" then
				set e to get occurrence of e at my parseStamp(occS)
				if e is missing value then error "NO_OCCURRENCE: the series has no occurrence at that time" number 9011
			end if
			if resp is "accept" then
				accept meeting e comment cmt sending response tell_
			else if resp is "tentative" then
				accept tentatively meeting e comment cmt sending response tell_
			else if resp is "decline" then
				decline meeting e comment cmt sending response tell_
			else
				error "unknown response " & resp number 9003
			end if
		end if
		end timeout
	end tell
	return n as text
end run
