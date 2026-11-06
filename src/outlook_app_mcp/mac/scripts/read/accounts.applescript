-- The Exchange, IMAP and POP accounts (Google accounts have no account object).
-- Output records (RS): email, name, kind (exchange|imap|pop), is default ("1"/"0")

on run argv
	my initSeps()
	set defEmail to ""
	set out to {}
	tell application "Microsoft Outlook"
		try
			set defEmail to email address of default account
		end try
		set exA to every exchange account
		set imA to every imap account
		set poA to every pop account
	end tell
	repeat with pair in {{exA, "exchange"}, {imA, "imap"}, {poA, "pop"}}
		repeat with a in (item 1 of pair)
			tell application "Microsoft Outlook"
				set aEmail to email address of a
				set aName to name of a
			end tell
			if aEmail is defEmail then
				set isDef to "1"
			else
				set isDef to "0"
			end if
			set end of out to my joinList({my clean(aEmail), my clean(aName), item 2 of pair, isDef}, FS)
		end repeat
	end repeat
	return my joinList(out, RS)
end run
