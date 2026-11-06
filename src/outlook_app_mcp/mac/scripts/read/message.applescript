-- Full read of one message. Does not change "is read".
-- argv: messageId, include headers ("1"/"0"), body as HTML ("1"/"0")
-- Output: header fields (FS), then RS, then the body (plain text, or HTML source):
--   id, folder id, folder name, account email, received (stamp), sent (stamp),
--   is read, has html, subject, sender (name US address),
--   to, cc, bcc (entries GS, each name US address),
--   attachments (entries GS, each name US size US content type),
--   flag, categories (names GS), is meeting, priority, conversation id,
--   meeting event id (meeting requests: the calendar event), headers (or "")

on run argv
	my initSeps()
	set mid to (item 1 of argv) as integer
	set wantHeaders to ((item 2 of argv) is "1")
	set wantHtml to ((item 3 of argv) is "1")
	set evId to ""
	set hdrs to ""
	tell application "Microsoft Outlook"
		set m to message id mid
		set fld to folder of m
		set fId to (id of fld) as text
		set fName to name of fld
		set aEmail to ""
		try
			set aEmail to email address of (account of m)
		end try
		set tr to time received of m
		set ts to time sent of m
		set readFlag to (is read of m) as text
		set htmlFlag to (has html of m) as text
		set meetingFlag to (is meeting of m) as text
		set sb to subject of m
		set s to sender of m
		set toR to to recipients of m
		set ccR to cc recipients of m
		set bccR to bcc recipients of m
		set atts to attachments of m
		set fl to todo flag of m
		set cats to categories of m
		set pr to priority of m
		set conv to ""
		try
			set conv to (conversation id of m) as text
		end try
		if meetingFlag is "true" then
			try
				set ev to meeting of m
				if ev is not missing value then set evId to (id of ev) as text
			end try
		end if
		if wantHeaders then set hdrs to headers of m
		if wantHtml then
			set body to content of m
		else
			set body to plain text content of m
		end if
	end tell
	set attOut to {}
	repeat with a in atts
		set aSize to ""
		set aType to ""
		tell application "Microsoft Outlook"
			set aName to name of a
			try
				set aSize to (file size of a) as text
			end try
			try
				set aType to content type of a
			end try
		end tell
		set end of attOut to my clean(aName) & US & aSize & US & my clean(aType)
	end repeat
	set catOut to {}
	repeat with c in cats
		tell application "Microsoft Outlook" to set cName to name of c
		set end of catOut to my clean(cName)
	end repeat
	set hdr to {mid as text, fId, my clean(fName), my clean(aEmail), my stamp(tr), my stamp(ts), readFlag, htmlFlag, my clean(sb), my addrText(s), my recipText(toR), my recipText(ccR), my recipText(bccR), my joinList(attOut, GS), my enumText(fl), my joinList(catOut, GS), meetingFlag, my enumText(pr), conv, evId, my clean(hdrs)}
	if body is missing value then set body to ""
	return my joinList(hdr, FS) & RS & body
end run

on recipText(recips)
	set out to {}
	repeat with r in recips
		tell application "Microsoft Outlook" to set e to email address of r
		set end of out to my addrText(e)
	end repeat
	return my joinList(out, GS)
end recipText
