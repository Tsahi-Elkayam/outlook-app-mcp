-- Send an existing draft.
-- argv: messageId
-- Output: the id that was sent.

on run argv
	my initSeps()
	tell application "Microsoft Outlook"
		set m to message id ((item 1 of argv) as integer)
		with timeout of 590 seconds -- Outlook asks the user to confirm a scripted send
			send m
		end timeout
	end tell
	return item 1 of argv
end run
