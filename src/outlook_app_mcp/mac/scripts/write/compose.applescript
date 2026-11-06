-- Create a new message, and send it if asked.
-- argv: mode ("draft" or "send"), account email ("" = default account),
--   subject, html body, to (GS), cc (GS), bcc (GS), attachment POSIX paths (GS),
--   importance ("", "low", "normal", "high")
-- Output: message id.

on run argv
	my initSeps()
	set mode to item 1 of argv
	set acct to item 2 of argv
	set subj to item 3 of argv
	set html to item 4 of argv
	set toL to my splitText(item 5 of argv, GS)
	set ccL to my splitText(item 6 of argv, GS)
	set bccL to my splitText(item 7 of argv, GS)
	set pr to item 9 of argv
	set attFiles to {}
	repeat with p in my splitText(item 8 of argv, GS)
		set end of attFiles to (POSIX file (p as text))
	end repeat
	if acct is not "" then set a to my findAccount(acct)
	tell application "Microsoft Outlook"
		if acct is "" then
			set m to make new outgoing message with properties {subject:subj, content:html}
		else
			set m to make new outgoing message with properties {subject:subj, content:html, account:a}
		end if
		repeat with ad in toL
			make new to recipient at m with properties {email address:{address:(ad as text)}}
		end repeat
		repeat with ad in ccL
			make new cc recipient at m with properties {email address:{address:(ad as text)}}
		end repeat
		repeat with ad in bccL
			make new bcc recipient at m with properties {email address:{address:(ad as text)}}
		end repeat
		repeat with f in attFiles
			make new attachment at m with properties {file:(contents of f)}
		end repeat
		if pr is "high" then
			set priority of m to priority high
		else if pr is "low" then
			set priority of m to priority low
		else if pr is "normal" then
			set priority of m to priority normal
		end if
		set mid to (id of m) as text
		if mode is "send" then
			with timeout of 590 seconds -- Outlook asks the user to confirm a scripted send
				send m
			end timeout
		end if
	end tell
	return mid
end run
