-- One contact in full.
-- argv: contactId
-- Output fields (FS): id, display name, first, middle, last, nickname, company, job title,
--   department, office, mobile, business phone, home phone, home street, home city,
--   home state, home zip, home country, business street, business city, business state,
--   business zip, business country, web page, birthday (stamp), emails (GS),
--   categories (GS), address book id, account email, account kind; then RS, then the note.

on run argv
	my initSeps()
	set cid to (item 1 of argv) as integer
	set emOut to {}
	set catOut to {}
	set abId to ""
	set aEmail to ""
	set aKind to ""
	tell application "Microsoft Outlook"
		set c to contact id cid
		set vals to {display name of c, first name of c, middle name of c, last name of c, nickname of c, company of c, job title of c, department of c, office of c, mobile number of c, business phone number of c, home phone number of c, home street address of c, home city of c, home state of c, home zip of c, home country of c, business street address of c, business city of c, business state of c, business zip of c, business country of c, business web page of c}
		set bd to birthday of c
		set ems to email addresses of c
		repeat with e in ems
			set end of emOut to my clean(address of (contents of e))
		end repeat
		set cats to categories of c
		repeat with x in cats
			set end of catOut to my clean(name of (contents of x))
		end repeat
		try
			set ab to address book of c
			set abId to (id of ab) as text
			set a to account of ab
			set aEmail to email address of a
			set aKind to my accountKind(a)
		end try
		set nt to plain text note of c
	end tell
	set out to {cid as text}
	repeat with v in vals
		set end of out to my clean(contents of v)
	end repeat
	set end of out to my stampOrEmpty(bd)
	set end of out to my joinList(emOut, GS)
	set end of out to my joinList(catOut, GS)
	set end of out to abId
	set end of out to my clean(aEmail)
	set end of out to aKind
	if nt is missing value then set nt to ""
	return my joinList(out, FS) & RS & nt
end run
