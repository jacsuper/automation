tell application "Microsoft Outlook"
	-- get the currently selected message or messages
	-- NOTE in this version it fails if the Outlook Reminder window is open, even if you select a message in the main window.
	set selectedMessages to current messages
	
	-- if there are no messages selected, warn the user and then quit
	if selectedMessages is {} then
		display dialog "Please select a message first and then run this script." with icon 1
		return
	end if
	
	set aMessage to item 1 of selectedMessages
	set emailAcct to account of aMessage
	set inBoxFolder to folder "Inbox" of emailAcct
	set ArchiveFolder to folder "Archive" of emailAcct
	
	repeat with theMessage in selectedMessages
		move theMessage to ArchiveFolder
	end repeat
end tell