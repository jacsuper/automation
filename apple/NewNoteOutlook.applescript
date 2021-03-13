tell application "Microsoft Outlook"
	
	set theContent to ""
	
	set theMessage to make new outgoing message with properties {subject:" Notes", content:theContent}
	make new recipient with properties {email address:{address:"jacsteyn@fb.com", name:"Jac Steyn"}} at end of to recipients of theMessage
	open theMessage -- for further editing
	
end tell

tell application "System Events"
	if quit delay ≠ 0 then set quit delay to 0
	
	
	tell application process "Microsoft Outlook"
		set frontmost to true
		
		delay 1
		
		tell (first window whose subrole is "AXStandardWindow")
			
			keystroke (ASCII character 9)
			keystroke (ASCII character 9)
			keystroke (ASCII character 9)
			keystroke (ASCII character 9)
			key code 123
			
			
		end tell
		
		
	end tell
end tell