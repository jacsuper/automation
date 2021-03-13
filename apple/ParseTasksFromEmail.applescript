use framework "Foundation"
use framework "AppKit"

set starter to {"*"}
set starter2 to {"**"}
set theaction to ""

property NSString : class "NSString"
property NSAttributedString : class "NSAttributedString"
property NSNumber : class "NSNumber"
property NSDictionary : class "NSDictionary"
property NSUTF8StringEncoding : a reference to 4
property NSCharacterEncodingDocumentOption : a reference to current application's NSCharacterEncodingDocumentOption

on HTMLDecode(HTMLString)
	set theString to NSString's stringWithString:HTMLString
	set dataStr to theString's dataUsingEncoding:NSUTF8StringEncoding
	set options to NSDictionary's dictionaryWithObject:NSUTF8StringEncoding forKey:(NSCharacterEncodingDocumentOption)
	set attStr to NSAttributedString's alloc()'s initWithHTML:dataStr options:options documentAttributes:(missing value)
	set outputStr to attStr's |string|()
	return outputStr as text
end HTMLDecode

tell application "Microsoft Outlook"
	set theMessage to item 1 of (get selection) # get the first email message
	#set theBody to (plain text content of theMessage) # get plain text message content
	set theBody to (the content of theMessage) # get plain text message content
end tell


# we are now going to break the content into paragraphs.

# set the delimiters to the carrige return used in outlook. This is \r instead of \n
set AppleScript's text item delimiters to character id 10

#get the seperated items.
set newPara to text items of HTMLDecode(theBody)
# set the delimiters to their normal setting
set AppleScript's text item delimiters to return


#we are now going to iterate through each paragraph item in the new newPara list

repeat with z from 1 to number of items in newPara
	#get an item in newPara 
	set this_para to item z of newPara
	
	#we check to see if the newPara item contains the starter
	#log "para" & this_para
	if this_para contains starter then
		# if it does then we set the delimiters to the store item i.e "name: "
		set AppleScript's text item delimiters to starter
		
		if this_para contains starter2 then
			set AppleScript's text item delimiters to starter2
		end if
		
		#now we set the second item of the item in the store item to the value of text item 2. i.e where we had {"name: ", ""} we will now have {"name: ", "Tootsie Roll"}
		set theaction to text item 2 of this_para
		set logtext to {"Action:"} & theaction
		
		log logtext
		
		
		tell application "Microsoft Outlook"
			set theContent to theaction
			
			set newMessage to make new outgoing message with properties {subject:theaction, content:theContent}
			make new recipient with properties {email address:{address:"jacsteyn@fb.com", name:"Jac Steyn"}} at end of to recipients of newMessage
			
			#open newMessage
			send newMessage
		end tell
		
		
		# set the delimiters to their normal setting
		set AppleScript's text item delimiters to return
		
	end if
end repeat

