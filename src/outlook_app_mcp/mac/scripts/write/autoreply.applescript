-- Set the automatic replies (out of office) of an Exchange account.
-- argv: account email, enabled, scheduled, start stamp ("" = keep), end stamp,
--   internal message (HTML), external message (HTML), send to external, send only to known
--   (flags "1"/"0")
-- Output: "ok"

on run argv
	my initSeps()
	set a to my findAccount(item 1 of argv)
	tell application "Microsoft Outlook"
		if class of a is not exchange account then error "NOT_EXCHANGE: automatic replies are only available for Exchange accounts" number 9010
		set rec to {enabled:((item 2 of argv) is "1"), scheduled:((item 3 of argv) is "1"), autoreply message:(item 6 of argv), external message:(item 7 of argv), send to external:((item 8 of argv) is "1"), send only to known:((item 9 of argv) is "1")}
		if (item 4 of argv) is not "" then set rec to rec & {start on:my parseStamp(item 4 of argv)}
		if (item 5 of argv) is not "" then set rec to rec & {end on:my parseStamp(item 5 of argv)}
		set autoreply data a new data rec
	end tell
	return "ok"
end run
