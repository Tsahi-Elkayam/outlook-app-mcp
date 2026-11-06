-- Send and receive: sync each named account (sends what waits in its Outbox).
-- argv: account email...
-- Output: number of accounts synced.

on run argv
	my initSeps()
	set n to 0
	repeat with e in argv
		set a to my findAccount(e as text)
		tell application "Microsoft Outlook" to sync a
		set n to n + 1
	end repeat
	return n as text
end run
