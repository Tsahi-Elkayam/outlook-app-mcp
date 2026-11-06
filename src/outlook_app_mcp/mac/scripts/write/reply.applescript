-- Reply, reply all or forward, with new text above the quoted message.
-- argv: messageId, mode ("reply", "replyall", "forward"), html text,
--   plain text, extra to-addresses (GS), send ("1" = send now, "0" = keep as draft),
--   drafts folder id (where a kept draft goes), extra cc addresses (GS),
--   attachment POSIX paths (GS)
-- Output: id of the new message.
-- Outlook makes a reply without a window as an unsaved message in "Temporary
-- Items"; a draft is moved to the account's Drafts folder so it is kept.

on run argv
	my initSeps()
	set mid to (item 1 of argv) as integer
	set mode to item 2 of argv
	set html to item 3 of argv
	set plainText to item 4 of argv
	set toL to my splitText(item 5 of argv, GS)
	set sendNow to ((item 6 of argv) is "1")
	set draftsId to (item 7 of argv) as integer
	set ccL to my splitText(item 8 of argv, GS)
	set attFiles to {}
	repeat with p in my splitText(item 9 of argv, GS)
		set end of attFiles to (POSIX file (p as text))
	end repeat
	tell application "Microsoft Outlook"
		set m to message id mid
		if mode is "forward" then
			set r to forward m opening window false
		else if mode is "replyall" then
			set r to reply to m opening window false reply to all true
		else
			set r to reply to m opening window false reply to all false
		end if
		repeat with ad in toL
			make new to recipient at r with properties {email address:{address:(ad as text)}}
		end repeat
		repeat with ad in ccL
			make new cc recipient at r with properties {email address:{address:(ad as text)}}
		end repeat
		repeat with f in attFiles
			make new attachment at r with properties {file:(contents of f)}
		end repeat
		set old to content of r
		set isHtml to has html of r
	end tell
	if isHtml then
		set newContent to my insertAfterBody(old, html)
	else
		if old is missing value then set old to ""
		set newContent to plainText & return & return & old
	end if
	tell application "Microsoft Outlook"
		set content of r to newContent
		if sendNow then
			set rid to (id of r) as text
			with timeout of 590 seconds -- Outlook asks the user to confirm a scripted send
				send r
			end timeout
		else
			-- The moved draft gets a new id and move returns nothing, so the new id is
			-- the one that appears in Drafts. Another draft can appear at the same moment
			-- (Outlook autosaving one the user is typing), so new ids are matched on the
			-- subject too, and two candidates are an error rather than a guess.
			set dest to mail folder id draftsId
			set subj to subject of r
			set idsBefore to id of every message of dest
			move r to dest
			set idsAfter to id of every message of dest
			set fresh to {}
			repeat with x in idsAfter
				if idsBefore does not contain (contents of x) then set end of fresh to (contents of x)
			end repeat
			if (count of fresh) > 1 then
				set same to {}
				repeat with x in fresh
					if subject of (message id x) is subj then set end of same to (contents of x)
				end repeat
				set fresh to same
			end if
			if (count of fresh) is 0 then error "DRAFT_LOST: the reply was moved to Drafts but could not be found there" number 9006
			if (count of fresh) > 1 then error "DRAFT_LOST: several new drafts with this subject appeared in Drafts at once; check Drafts in Outlook" number 9006
			set rid to (item 1 of fresh) as text
		end if
	end tell
	return rid
end run

-- Put h right after the opening <body ...> tag of c (or in front if there is none).
on insertAfterBody(c, h)
	if c is missing value then return h
	set i to offset of "<body" in c
	if i is 0 then return h & c
	set j to offset of ">" in (text i thru -1 of c)
	if j is 0 then return h & c
	set cut to i + j - 1
	if cut is (length of c) then return c & h
	return (text 1 thru cut of c) & h & (text (cut + 1) thru -1 of c)
end insertAfterBody
