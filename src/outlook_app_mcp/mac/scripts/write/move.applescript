-- Move messages to another mail folder.
-- argv: destinationFolderId, messageId...
-- Output: message ids after the move (RS), in the same order; "" where the id
-- changed. Outlook's move command returns nothing. A message that was in a folder
-- keeps its id; an unsaved one (Temporary Items) gets a new id, which is not
-- reported. The old id counts only if it now names a message in the destination.

on run argv
	my initSeps()
	set out to {}
	set destId to (item 1 of argv) as integer
	tell application "Microsoft Outlook"
		set dest to mail folder id destId
		repeat with t in (items 2 thru -1 of argv)
			move (message id (t as integer)) to dest
			set newId to ""
			try
				if ((id of folder of (message id (t as integer))) as integer) is destId then set newId to t as text
			end try
			set end of out to newId
		end repeat
	end tell
	return my joinList(out, RS)
end run
