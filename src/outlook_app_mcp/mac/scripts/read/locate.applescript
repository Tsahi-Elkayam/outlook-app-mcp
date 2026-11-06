-- The folder of each message.
-- argv: messageId...
-- Output records (RS) of fields (FS): message id, folder id. Unknown ids are left out.

on run argv
	my initSeps()
	set out to {}
	repeat with t in argv
		try
			tell application "Microsoft Outlook"
				set fId to (id of (folder of (message id (t as integer)))) as text
			end tell
			set end of out to (t as text) & FS & fId
		end try
	end repeat
	return my joinList(out, RS)
end run
