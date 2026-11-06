-- Automatic replies (out of office) of an Exchange account.
-- argv: account email
-- Output fields (FS): enabled, scheduled, start (stamp), end (stamp), send to external,
--   send only to known, internal message (RS-free), external message

on run argv
	my initSeps()
	set a to my findAccount(item 1 of argv)
	tell application "Microsoft Outlook"
		if class of a is not exchange account then error "NOT_EXCHANGE: automatic replies are only available for Exchange accounts" number 9010
		set r to query autoreply data a
		set vOn to (enabled of r) as text
		set vSched to (scheduled of r) as text
		set vStart to ""
		set vEnd to ""
		try
			set vStart to my stamp(start on of r)
		end try
		try
			set vEnd to my stamp(end on of r)
		end try
		set vExt to (send to external of r) as text
		set vKnown to (send only to known of r) as text
		set vInt to ""
		set vExtMsg to ""
		try
			set vInt to autoreply message of r
		end try
		try
			set vExtMsg to external message of r
		end try
	end tell
	if vInt is missing value then set vInt to ""
	if vExtMsg is missing value then set vExtMsg to ""
	return my joinList({vOn, vSched, vStart, vEnd, vExt, vKnown, my clean(vInt), my clean(vExtMsg)}, FS)
end run
