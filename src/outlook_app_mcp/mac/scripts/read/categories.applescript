-- Outlook's category list.
-- Output records (RS): name, color (red,green,blue on 0-65535, or "")

on run argv
	my initSeps()
	set out to {}
	tell application "Microsoft Outlook" to set cats to every category
	repeat with c in cats
		set col to ""
		tell application "Microsoft Outlook"
			set nm to name of c
			try
				set rgb to color of c
				set col to ((item 1 of rgb) as text) & "," & ((item 2 of rgb) as text) & "," & ((item 3 of rgb) as text)
			end try
		end tell
		set end of out to my clean(nm) & FS & col
	end repeat
	return my joinList(out, RS)
end run
