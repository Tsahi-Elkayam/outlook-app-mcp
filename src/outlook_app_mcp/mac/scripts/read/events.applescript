-- Events overlapping a time range, plus recurring series that may recur into it.
-- argv: startStamp, endStamp, calendarId...
-- Output records (RS) of fields (FS):
--   kind ("single" or "series"), event id, calendar id, subject, start (stamp),
--   end (stamp), location, all day, organizer, free busy, icalendar data, time zone name
--   (series only; empty for single events), is occurrence, series id, original start
--   (stamp) -- the last three for changed occurrences (exceptions) of a series
-- Series are expanded into occurrences in Python from their iCalendar data,
-- because Outlook's own occurrence command creates exceptions (a change).

on run argv
	my initSeps()
	set startD to my parseStamp(item 1 of argv)
	set endD to my parseStamp(item 2 of argv)
	set calIds to items 3 thru -1 of argv
	set out to {}
	repeat with cidText in calIds
		set cid to cidText as integer
		tell application "Microsoft Outlook"
			set c to calendar id cid
			set ev to (every calendar event of c whose start time < endD and end time > startD and is recurring is false)
			set ser to (every calendar event of c whose is recurring is true and start time < endD)
		end tell
		repeat with e in ev
			set end of out to my row("single", e, cid, false)
		end repeat
		repeat with e in ser
			set end of out to my row("series", e, cid, true)
		end repeat
	end repeat
	return my joinList(out, RS)
end run

on row(kind, e, cid, withIcs)
	set ics to ""
	set org to ""
	set isOcc to "false"
	set serId to ""
	set origS to ""
	tell application "Microsoft Outlook"
		set eId to (id of e) as text
		set sb to subject of e
		set startT to start time of e
		set endT to end time of e
		set loc to location of e
		set ad to (all day flag of e) as text
		try
			set org to organizer of e
		end try
		set fb to free busy status of e
		if withIcs then set ics to icalendar data of e
		try
			set isOcc to (is occurrence of e) as text
		end try
		if isOcc is "true" then
			try
				set serId to (id of (master of e)) as text
			end try
			try
				set origS to my stamp(recurrence id of e)
			end try
		end if
	end tell
	return kind & FS & eId & FS & (cid as text) & FS & my clean(sb) & FS & my stamp(startT) & FS & my stamp(endT) & FS & my clean(loc) & FS & ad & FS & my clean(org) & FS & my enumText(fb) & FS & my clean(ics) & FS & my tzName(e) & FS & isOcc & FS & serId & FS & origS
end row
