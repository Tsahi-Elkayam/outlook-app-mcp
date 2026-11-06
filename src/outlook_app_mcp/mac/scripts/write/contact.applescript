-- Create or update a contact in Outlook's address book.
-- argv: "create", address book id ("" = Outlook's default), key, value...
--   or: "update", contactId, key, value...
--   or: "delete", contactId  (Python checks it is in an Exchange account's address book,
--                             where Outlook moves deleted items to Deleted Items)
-- Keys: first, middle, last, nickname, company, jobtitle, department, office, emails
--   (addresses joined by GS; replaces the list), mobile, business, home, hstreet, hcity,
--   hstate, hzip, hcountry, bstreet, bcity, bstate, bzip, bcountry, web, birthday (stamp,
--   or "" to clear), note, categories (names GS; must exist).
-- Output: contact id.

on run argv
	my initSeps()
	set action to item 1 of argv
	set target to item 2 of argv
	set pairs to {}
	if (count of argv) > 2 then set pairs to items 3 thru -1 of argv
	tell application "Microsoft Outlook"
		if action is "create" then
			if target is "" then
				set c to make new contact with properties {first name:""}
			else
				set c to make new contact at (address book id (target as integer)) with properties {first name:""}
			end if
		else if action is "update" then
			set c to contact id (target as integer)
		else if action is "delete" then
			delete (contact id (target as integer))
			return target
		else
			error "unknown action " & action number 9003
		end if
	end tell
	repeat with i from 1 to (count of pairs) by 2
		set k to item i of pairs
		set v to item (i + 1) of pairs
		if k is "birthday" and v is not "" then set dv to my parseStamp(v)
		if k is "emails" then
			set addrList to my splitText(v, GS)
			set recs to {}
			-- Built inside the tell block, so "address" and "type" are Outlook's record keys.
			tell application "Microsoft Outlook"
				repeat with ad in addrList
					set end of recs to {address:(ad as text), type:«constant ****eWrk»}
				end repeat
			end tell
		end if
		tell application "Microsoft Outlook"
			if k is "first" then
				set first name of c to v
			else if k is "last" then
				set last name of c to v
			else if k is "company" then
				set company of c to v
			else if k is "jobtitle" then
				set job title of c to v
			else if k is "emails" then
				set email addresses of c to recs
			else if k is "mobile" then
				set mobile number of c to v
			else if k is "business" then
				set business phone number of c to v
			else if k is "home" then
				set home phone number of c to v
			else if k is "note" then
				set plain text note of c to v
			else if k is "middle" then
				set middle name of c to v
			else if k is "nickname" then
				set nickname of c to v
			else if k is "department" then
				set department of c to v
			else if k is "office" then
				set office of c to v
			else if k is "hstreet" then
				set home street address of c to v
			else if k is "hcity" then
				set home city of c to v
			else if k is "hstate" then
				set home state of c to v
			else if k is "hzip" then
				set home zip of c to v
			else if k is "hcountry" then
				set home country of c to v
			else if k is "bstreet" then
				set business street address of c to v
			else if k is "bcity" then
				set business city of c to v
			else if k is "bstate" then
				set business state of c to v
			else if k is "bzip" then
				set business zip of c to v
			else if k is "bcountry" then
				set business country of c to v
			else if k is "web" then
				set business web page of c to v
			else if k is "birthday" then
				if v is "" then
					set birthday of c to missing value
				else
					set birthday of c to dv
				end if
			else if k is "categories" then
				set wanted to {}
				repeat with nm in my splitText(v, GS)
					set nm to nm as text
					if nm is not "" then set end of wanted to (first category whose name is nm)
				end repeat
				set categories of c to wanted
			else
				error "unknown key " & k number 9003
			end if
		end tell
	end repeat
	tell application "Microsoft Outlook" to set cid to (id of c) as text
	return cid
end run
