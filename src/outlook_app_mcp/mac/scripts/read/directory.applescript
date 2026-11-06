-- People in the organization directory (GAL) of an Exchange account.
-- argv: account email, text to look up
-- Output records (RS): name, address

on run argv
	my initSeps()
	set a to my findAccount(item 1 of argv)
	set q to item 2 of argv
	tell application "Microsoft Outlook"
		if class of a is not exchange account then error "NOT_EXCHANGE: the directory is only available through an Exchange account" number 9010
		set res to resolve names a for name q
	end tell
	set out to {}
	repeat with r in res
		set end of out to my addrText(contents of r)
	end repeat
	return my joinList(out, RS)
end run
