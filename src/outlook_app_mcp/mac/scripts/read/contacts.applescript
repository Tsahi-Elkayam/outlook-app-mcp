-- Every contact in Outlook's local address books (not the organization directory).
-- Output records (RS) of fields (FS):
--   id, display name, first name, last name, company, job title,
--   emails (addresses GS), mobile, business phone, home phone, categories (names GS)

on run argv
	my initSeps()
	tell application "Microsoft Outlook"
		set ids to id of every contact
		set dns to display name of every contact
		set fns to first name of every contact
		set lns to last name of every contact
		set cos to company of every contact
		set jts to job title of every contact
		set ems to email addresses of every contact
		set mos to mobile number of every contact
		set bps to business phone number of every contact
		set hps to home phone number of every contact
		set cts to categories of every contact
	end tell
	set n to count of ids
	if (count of dns) is not n or (count of ems) is not n or (count of hps) is not n or (count of cts) is not n then
		error "FOLDER_CHANGED: contacts changed while they were being read" number 9001
	end if
	set out to {}
	repeat with i from 1 to n
		set emOut to {}
		repeat with e in (item i of ems)
			tell application "Microsoft Outlook" to set addr to address of e
			set end of emOut to my clean(addr)
		end repeat
		set catOut to {}
		repeat with c in (item i of cts)
			tell application "Microsoft Outlook" to set cName to name of (contents of c)
			set end of catOut to my clean(cName)
		end repeat
		set end of out to my joinList({(item i of ids) as text, my clean(item i of dns), my clean(item i of fns), my clean(item i of lns), my clean(item i of cos), my clean(item i of jts), my joinList(emOut, GS), my clean(item i of mos), my clean(item i of bps), my clean(item i of hps), my joinList(catOut, GS)}, FS)
	end repeat
	return my joinList(out, RS)
end run
