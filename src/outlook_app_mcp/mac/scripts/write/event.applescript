-- Create, update or delete a calendar event.
-- argv: "create", calendarId ("" = default calendar), start (stamp), end (stamp),
--         key, value, key, value...
--   or: "update", eventId, key, value...
--   or: "occurrence", seriesId, occurrence start (stamp), key, value...
--         (changes one occurrence; Outlook keeps it as an exception of the series)
--   or: "send", eventId, occurrence start (stamp, or "" for the event itself)
--         (sends the invitation or update to the attendees)
--   or: "delete", eventId
-- Keys: subject, start (stamp), end (stamp), location, html, allday ("1"/"0"),
--   freebusy (busy, free, tentative, oof), reminder (minutes, or "" for none),
--   private ("1"/"0"), addreq / addopt (addresses GS: attendees to add),
--   categories (names GS; each must exist; replaces the list),
--   recurrence (fields US: type daily|weekly|absolute monthly|absolute yearly, interval,
--     days (names GS), day of month, month, end type none|date|count, end data
--     (stamp or count), series start (stamp)).
-- "update" gets start and end from Python in an order that never puts start after end.
-- Delete refuses events that have attendees: those are meetings, and deleting
-- one would remove it without telling anyone. Use respond.applescript instead.
-- Output: event id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	if action is "create" then
		set startD to my parseStamp(item 3 of argv)
		set endD to my parseStamp(item 4 of argv)
		set pairs to {}
		if (count of argv) > 4 then set pairs to items 5 thru -1 of argv
		tell application "Microsoft Outlook"
			if target is "" then
				set c to default calendar of default account
			else
				set c to calendar id (target as integer)
			end if
			set e to make new calendar event at c with properties {subject:"", start time:startD, end time:endD}
		end tell
		-- Attendees go on last, after the event has synced: Outlook drops a meeting whose
		-- attendees were added before its first sync to Exchange and leaves a blank event.
		set plainPairs to {}
		set peoplePairs to {}
		repeat with i from 1 to (count of pairs) by 2
			if item i of pairs is in {"addreq", "addopt", "categories"} then
				set peoplePairs to peoplePairs & {item i of pairs, item (i + 1) of pairs}
			else
				set plainPairs to plainPairs & {item i of pairs, item (i + 1) of pairs}
			end if
		end repeat
		my applyPairs(e, plainPairs)
		-- On Exchange, changes made while the new event first syncs are lost; wait it out
		-- here, so that the caller's next change is safe too.
		set onExchange to false
		tell application "Microsoft Outlook"
			try
				set onExchange to ((class of (account of e)) is exchange account)
			end try
		end tell
		if onExchange then my waitForSync(e)
		if peoplePairs is not {} then my applyPairs(e, peoplePairs)
	else if action is "update" then
		set pairs to {}
		if (count of argv) > 2 then set pairs to items 3 thru -1 of argv
		tell application "Microsoft Outlook" to set e to calendar event id (target as integer)
		my applyPairs(e, pairs)
	else if action is "occurrence" then
		set occD to my parseStamp(item 3 of argv)
		set pairs to {}
		if (count of argv) > 3 then set pairs to items 4 thru -1 of argv
		tell application "Microsoft Outlook"
			set seriesEv to calendar event id (target as integer)
			set e to get occurrence of seriesEv at occD
		end tell
		if e is missing value then error "NO_OCCURRENCE: the series has no occurrence at that time" number 9011
		my applyPairs(e, pairs)
	else if action is "send" then
		tell application "Microsoft Outlook"
			set e to calendar event id (target as integer)
			if (item 3 of argv) is not "" then
				set e to get occurrence of e at my parseStamp(item 3 of argv)
				if e is missing value then error "NO_OCCURRENCE: the series has no occurrence at that time" number 9011
			end if
			with timeout of 590 seconds -- Outlook asks the user to confirm a scripted send
				send meeting e
			end timeout
		end tell
	else if action is "delete" then
		tell application "Microsoft Outlook"
			set e to calendar event id (target as integer)
			if (count of attendees of e) > 0 then error "HAS_ATTENDEES: event " & target & " is a meeting with attendees" number 9004
			delete e
		end tell
		return target
	else
		error "unknown action " & action number 9003
	end if
	tell application "Microsoft Outlook" to set eid to (id of e) as text
	return eid
end run

on applyPairs(e, pairs)
	repeat with i from 1 to (count of pairs) by 2
		set k to item i of pairs
		set v to item (i + 1) of pairs
		if k is "start" or k is "end" then set dv to my parseStamp(v)
		tell application "Microsoft Outlook"
			if k is "subject" then
				set subject of e to v
			else if k is "start" then
				set start time of e to dv
			else if k is "end" then
				set end time of e to dv
			else if k is "location" then
				set location of e to v
			else if k is "html" then
				set content of e to v
			else if k is "allday" then
				set all day flag of e to (v is "1")
			else if k is "freebusy" then
				if v is "busy" then
					set free busy status of e to «constant ****eSBu»
				else if v is "free" then
					set free busy status of e to «constant ****eSFr»
				else if v is "tentative" then
					set free busy status of e to «constant ****eSTe»
				else if v is "oof" then
					set free busy status of e to «constant ****eSOO»
				else
					error "unknown free busy " & v number 9003
				end if
			else if k is "private" then
				set is private of e to (v is "1")
			else if k is "addreq" or k is "addopt" then
				repeat with ad in my splitText(v, GS)
					set ad to ad as text
					if ad is not "" then
						if k is "addreq" then
							make new required attendee at e with properties {email address:{address:ad}}
						else
							make new optional attendee at e with properties {email address:{address:ad}}
						end if
					end if
				end repeat
			else if k is "categories" then
				set wanted to {}
				repeat with nm in my splitText(v, GS)
					set nm to nm as text
					if nm is not "" then set end of wanted to (first category whose name is nm)
				end repeat
				set categories of e to wanted
			else if k is "recurrence" then
				set recurrence of e to my recurrenceRecord(v)
			else if k is "reminder" then
				if v is "" then
					set has reminder of e to false
				else
					set has reminder of e to true
					set reminder time of e to (v as integer)
				end if
			else
				error "unknown key " & k number 9003
			end if
		end tell
	end repeat
end applyPairs

on recurrenceRecord(v)
	set f to my splitText(v, US)
	set itemKind to item 1 of f
	set startD to my parseStamp(item 8 of f)
	set endType to item 6 of f
	tell application "Microsoft Outlook"
		if endType is "date" then
			set ed to {end type:end date type, data:my parseStamp(item 7 of f)}
		else if endType is "count" then
			set ed to {end type:end numbered type, data:((item 7 of f) as integer)}
		else
			set ed to {end type:no end type}
		end if
		set r to {occurrence interval:((item 2 of f) as integer), start date:startD, end date:ed}
		if itemKind is "daily" then
			set r to r & {recurrence type:daily}
		else if itemKind is "weekly" then
			set dl to my splitText(item 3 of f, GS)
			set dw to {sunday:(dl contains "sunday"), monday:(dl contains "monday"), tuesday:(dl contains "tuesday"), wednesday:(dl contains "wednesday"), thursday:(dl contains "thursday"), friday:(dl contains "friday"), saturday:(dl contains "saturday")}
			set r to r & {recurrence type:weekly, days of week:dw}
		else if itemKind is "absolute monthly" then
			set r to r & {recurrence type:absolute monthly, day of month:((item 4 of f) as integer)}
		else if itemKind is "absolute yearly" then
			set r to r & {recurrence type:absolute yearly, day of month:((item 4 of f) as integer), month number:((item 5 of f) as integer)}
		else
			error "unknown recurrence " & itemKind number 9003
		end if
	end tell
	return r
end recurrenceRecord

on waitForSync(e)
	repeat 60 times
		tell application "Microsoft Outlook" to set x to exchange id of e
		if x is not missing value and x is not "" then
			delay 3 -- the id comes before Exchange is done with the event; changes in between are lost
			return
		end if
		delay 0.5
	end repeat
	tell application "Microsoft Outlook" to set eid to (id of e) as text
	error "NOT_SYNCED: event " & eid & " was saved without attendees; it did not reach Exchange within 30 seconds, and attendees added before that are lost" number 9012
end waitForSync
