-- One note in full.
-- argv: noteId
-- Output: id, title, modified (stamp), folder id, categories (names GS), then RS, then the text.

on run argv
	my initSeps()
	set nid to (item 1 of argv) as integer
	set catOut to {}
	tell application "Microsoft Outlook"
		set n to note id nid
		set nm to name of n
		set md to modification date of n
		set fid to ""
		try
			set fid to (id of (folder of n)) as text
		end try
		set cats to categories of n
		repeat with c in cats
			set end of catOut to my clean(name of (contents of c))
		end repeat
		set body to plain text content of n
	end tell
	if body is missing value then set body to ""
	return my joinList({nid as text, my clean(nm), my stampOrEmpty(md), fid, my joinList(catOut, GS)}, FS) & RS & body
end run
