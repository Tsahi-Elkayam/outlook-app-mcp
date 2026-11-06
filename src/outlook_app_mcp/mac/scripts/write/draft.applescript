-- Change a draft in place. Python checks first that it is in a Drafts folder.
-- argv: messageId, then pairs of key, value:
--   to / cc / bcc (addresses GS; replaces that list), subject, html (new body),
--   add (attachment POSIX paths GS), remove (1-based attachment indexes GS,
--   highest first), importance (low, normal, high)
-- Recipients and attachments are elements of the draft: removing them changes only
-- this unsent draft.
-- Output: message id.

on run argv
	my initSeps()
	set mid to (item 1 of argv) as integer
	tell application "Microsoft Outlook" to set m to message id mid
	repeat with i from 2 to (count of argv) by 2
		set k to item i of argv
		set v to item (i + 1) of argv
		tell application "Microsoft Outlook"
			if k is "to" then
				delete (every to recipient of m)
				repeat with ad in my splitText(v, GS)
					make new to recipient at m with properties {email address:{address:(ad as text)}}
				end repeat
			else if k is "cc" then
				delete (every cc recipient of m)
				repeat with ad in my splitText(v, GS)
					make new cc recipient at m with properties {email address:{address:(ad as text)}}
				end repeat
			else if k is "bcc" then
				delete (every bcc recipient of m)
				repeat with ad in my splitText(v, GS)
					make new bcc recipient at m with properties {email address:{address:(ad as text)}}
				end repeat
			else if k is "subject" then
				set subject of m to v
			else if k is "html" then
				set content of m to v
			else if k is "add" then
				repeat with p in my splitText(v, GS)
					make new attachment at m with properties {file:(POSIX file (p as text))}
				end repeat
			else if k is "remove" then
				repeat with ix in my splitText(v, GS)
					delete attachment ((ix as text) as integer) of m
				end repeat
			else if k is "importance" then
				if v is "high" then
					set priority of m to priority high
				else if v is "low" then
					set priority of m to priority low
				else
					set priority of m to priority normal
				end if
			else
				error "unknown key " & k number 9003
			end if
		end tell
	end repeat
	return mid as text
end run
