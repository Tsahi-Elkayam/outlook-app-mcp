-- Shared handlers. Python prepends this file to every script before running it.
-- Separators: FS between fields, RS between records, GS between list entries
-- inside a field, US between sub-fields of a list entry.

global FS, RS, GS, US

on initSeps()
	set FS to character id 31
	set RS to character id 30
	set GS to character id 29
	set US to character id 28
end initSeps

-- Text with separator characters replaced by spaces; missing value -> "".
on clean(t)
	if t is missing value then return ""
	set t to t as text
	set saved to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {FS, RS, GS, US}
	set parts to text items of t
	set AppleScript's text item delimiters to " "
	set t to parts as text
	set AppleScript's text item delimiters to saved
	return t
end clean

on joinList(lst, sep)
	set saved to AppleScript's text item delimiters
	set AppleScript's text item delimiters to sep
	set res to lst as text
	set AppleScript's text item delimiters to saved
	return res
end joinList

on splitText(t, sep)
	if t is "" then return {}
	set saved to AppleScript's text item delimiters
	set AppleScript's text item delimiters to sep
	set parts to text items of t
	set AppleScript's text item delimiters to saved
	return parts
end splitText

-- Local date as "Y-M-D-S" (S = seconds since midnight). Locale-independent.
on stamp(d)
	if d is missing value then return ""
	return ((year of d) as text) & "-" & (((month of d) as integer) as text) & "-" & ((day of d) as text) & "-" & ((time of d) as text)
end stamp

-- Inverse of stamp().
on parseStamp(t)
	set p to my splitText(t, "-")
	set d to current date
	set day of d to 1
	set year of d to (item 1 of p) as integer
	set month of d to (item 2 of p) as integer
	set day of d to (item 3 of p) as integer
	set time of d to (item 4 of p) as integer
	return d
end parseStamp

-- "name<US>address" from an Outlook email address record.
on addrText(e)
	set nm to ""
	set ad to ""
	-- "name" and "address" must resolve to Outlook's record keys.
	tell application "Microsoft Outlook"
		try
			set nm to name of e
		end try
		try
			set ad to «class radd» of e -- address
		end try
	end tell
	return my clean(nm) & US & my clean(ad)
end addrText

-- The account whose email address is acct (case-insensitive).
-- This file uses raw codes for Outlook's terms so it compiles without Outlook
-- installed (CI); the scripts appended to it need Outlook's dictionary anyway.
on findAccount(acct)
	tell application "Microsoft Outlook"
		-- exchange, imap and pop accounts
		set accts to (every «class Eact») & (every «class iact») & (every «class pact»)
		repeat with a in accts
			if («class emad» of a) is acct then return contents of a -- email address
		end repeat
	end tell
	error "NO_ACCOUNT: no account with address " & acct number 9002
end findAccount

-- An Outlook enumeration value as text. Depending on context this is the term
-- ("not completed") or the raw form ("«constant FlagFlNC»"); Python maps both.
on enumText(v)
	if v is missing value then return ""
	tell application "Microsoft Outlook"
		return v as text
	end tell
end enumText

-- Name of an event's time zone ("Asia/Jerusalem", "Etc/GMT"), or "".
on tzName(e)
	set n to ""
	tell application "Microsoft Outlook"
		try
			-- Read the record first: "name of (timezone of e)" fails as one specifier.
			set tzr to «class rTmZ» of e -- timezone
			set n to name of tzr
		end try
	end tell
	return my clean(n)
end tzName

-- A date as a stamp, or "" for missing value.
on stampOrEmpty(d)
	if d is missing value then return ""
	try
		return my stamp(d)
	on error
		return ""
	end try
end stampOrEmpty

-- exchange, imap, pop or other. Raw class codes (exchange account, imap account,
-- pop account), so this file compiles without Outlook's dictionary.
on accountKind(a)
	tell application "Microsoft Outlook" to set c to class of a
	if c is «class Eact» then return "exchange"
	if c is «class iact» then return "imap"
	if c is «class pact» then return "pop"
	return "other"
end accountKind
