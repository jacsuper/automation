#set saveToFolder to POSIX path of (choose folder with prompt "Choose the destination folder")
set saveToFolder to "/Users/jacsteyn/Dropbox (Facebook)/Attachments/"
set ctr to 0
tell application "Microsoft Outlook"
	set selectedMessages to current messages
	repeat with msg in selectedMessages
		
		set sentstamp to time sent of msg
		
		set y to year of sentstamp
		set m to month of sentstamp
		set d to day of sentstamp
		set rdate to y & "-" & m & "-" & d
		set ctr to ctr + 1
		
		set attFiles to attachments of msg
		set actr to 0
		repeat with f in attFiles
			set attName to (get the name of f)
			log attName
			set saveAsName to saveToFolder & attName
			
			set actr to actr + 1
			
			save f in POSIX file saveAsName
		end repeat
	end repeat
end tell

display dialog "" & ctr & " messages were processed" buttons {"OK"} default button 1
return ctr