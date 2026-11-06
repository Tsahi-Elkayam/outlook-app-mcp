-- Signatures. argv: none (list them) or a signature name (its HTML).
-- Output without argv: records (RS) of name, plain-text preview (first 200 characters).
-- Output with a name: the signature's HTML content.

on run argv
	my initSeps()
	if (count of argv) > 0 then
		set want to item 1 of argv
		tell application "Microsoft Outlook"
			repeat with s in signatures
				if (name of s) is want then return content of s
			end repeat
		end tell
		error "NO_SIGNATURE: no signature named " & want number 9007
	end if
	set out to {}
	tell application "Microsoft Outlook"
		repeat with s in signatures
			set nm to name of s
			set pt to plain text content of s
			if pt is missing value then set pt to ""
			if (length of pt) > 200 then set pt to text 1 thru 200 of pt
			set end of out to my clean(nm) & FS & my clean(pt)
		end repeat
	end tell
	return my joinList(out, RS)
end run
