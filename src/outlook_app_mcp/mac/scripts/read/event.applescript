-- Full read of one calendar event.
-- argv: eventId
-- Output: header fields (FS), then RS, then the plain-text description:
--   id, calendar id, subject, start (stamp), end (stamp), location, all day,
--   is recurring, organizer, free busy, is private, has reminder,
--   reminder minutes, attendees (entries GS, each name US address US type US status),
--   categories (names GS), time zone name, is occurrence, request responses,
--   recurrence: type (enum text), interval, days of week (names GS), day of month,
--   month, ordinal, end type (enum text), end data (count, or a date stamp),
--   iCalendar data (recurring series only; FS/RS-free)

on run argv
	my initSeps()
	set eid to (item 1 of argv) as integer
	set org to ""
	set calId to ""
	tell application "Microsoft Outlook"
		set e to calendar event id eid
		try
			set calId to (id of (calendar of e)) as text
		end try
		set sb to subject of e
		set startT to start time of e
		set endT to end time of e
		set loc to location of e
		set ad to (all day flag of e) as text
		set rec to (is recurring of e) as text
		try
			set org to organizer of e
		end try
		set fb to free busy status of e
		set pv to (is private of e) as text
		set hr to (has reminder of e) as text
		set rm to (reminder time of e) as text
		set atts to attendees of e
		set cats to categories of e
		set body to plain text content of e
		set isOcc to (is occurrence of e) as text
		set reqResp to ""
		try
			set reqResp to (request responses of e) as text
		end try
		set rType to ""
		set rIv to ""
		set rDays to {}
		set rDom to ""
		set rMon to ""
		set rOrd to ""
		set rEnd to ""
		set rData to ""
		set ics to ""
		if rec is "true" then
			try
				set ics to icalendar data of e
			end try
			try
				set r to recurrence of e
				set rType to my enumText(recurrence type of r)
				set rIv to (occurrence interval of r) as text
				try
					set dw to days of week of r
					if sunday of dw then set end of rDays to "sunday"
					if monday of dw then set end of rDays to "monday"
					if tuesday of dw then set end of rDays to "tuesday"
					if wednesday of dw then set end of rDays to "wednesday"
					if thursday of dw then set end of rDays to "thursday"
					if friday of dw then set end of rDays to "friday"
					if saturday of dw then set end of rDays to "saturday"
				end try
				try
					set rDom to (day of month of r) as text
				end try
				try
					set rMon to (month number of r) as text
				end try
				try
					set rOrd to (ordinal of r) as text
				end try
				try
					set ed to end date of r
					set rEnd to my enumText(end type of ed)
					try
						set dv to data of ed
						if class of dv is date then
							set rData to my stamp(dv)
						else
							set rData to dv as text
						end if
					end try
				end try
			end try
		end if
	end tell
	set attOut to {}
	repeat with a in atts
		tell application "Microsoft Outlook"
			set em to email address of a
			set aType to type of a
			set aStatus to status of a
		end tell
		set end of attOut to my addrText(em) & US & my enumText(aType) & US & my enumText(aStatus)
	end repeat
	set catOut to {}
	repeat with c in cats
		tell application "Microsoft Outlook" to set cName to name of c
		set end of catOut to my clean(cName)
	end repeat
	set hdr to {eid as text, calId, my clean(sb), my stamp(startT), my stamp(endT), my clean(loc), ad, rec, my clean(org), my enumText(fb), pv, hr, rm, my joinList(attOut, GS), my joinList(catOut, GS), my tzName(e), isOcc, reqResp, rType, rIv, my joinList(rDays, GS), rDom, rMon, rOrd, rEnd, rData, my clean(ics)}
	if body is missing value then set body to ""
	return my joinList(hdr, FS) & RS & body
end run
