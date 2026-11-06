-- Save one attachment of a message to a file. Python checks the path first and
-- never passes one that already exists.
-- argv: messageId, attachment index (1-based), destination POSIX path
-- Output: the destination path.

on run argv
	my initSeps()
	set dest to POSIX file (item 3 of argv)
	tell application "Microsoft Outlook"
		set m to message id ((item 1 of argv) as integer)
		set a to attachment ((item 2 of argv) as integer) of m
		save a in dest
	end tell
	return item 3 of argv
end run
