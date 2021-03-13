tell application "Microsoft Outlook"
	
	set nextMeet to my getNextMeeting()
	
	if nextMeet ≠ "Untitled Meeting" then
		#display dialog "" & subject of nextMeet
		set html to content of nextMeet
		
		set tid to my text item delimiters
		set my text item delimiters to {"<a href=\"", "\" "}
		set textBits to text items of html
		set my text item delimiters to tid
		
		#groupcall or bluejeans or zoom
		
		set clickUrl to item 1 of textBits
		
		repeat with i from 1 to number of items in textBits
			set clickUrl to item i of textBits
			
			if clickUrl contains "groupcall" or clickUrl contains "bluejeans" or clickUrl contains "zoom" then
				
				tell application "Google Chrome"
					activate
					open location clickUrl
					delay 1
					activate
				end tell
				
				return
			end if
			
		end repeat
		
	end if
	
end tell

on getNextMeeting()
	
	
	-- Set the target Calendar
	set calendars to "Calendar"
	-- Set the start time margin to allow the return of the meeting x minutes before the start
	set startTimeMargin to (5 * minutes)
	-- The default meeting name in case no meeting is found
	set defaultMeeting to "Untitled Meeting"
	
	set now to current date
	copy now to today
	set time of today to 0
	set tomorrow to today + (3 * days)
	
	set delims to AppleScript's text item delimiters
	if calendars contains ", " then
		set AppleScript's text item delimiters to {", "}
	else
		set AppleScript's text item delimiters to {","}
	end if
	
	set calendarNames to every text item of calendars
	set AppleScript's text item delimiters to delims
	
	tell application "Microsoft Outlook"
		set calendarId to {}
		set fallbackMeeting to ""
		repeat with i from 1 to number of items in calendarNames
			set calendarName to item i of calendarNames
			set calendarId to calendarId & (id of every calendar whose name is calendarName)
		end repeat
		
		set calendarEvents to {}
		repeat with i from 1 to number of items in calendarId
			set CalID to item i of calendarId
			
			tell (calendar id CalID)
				set calendarEvents to (every calendar event whose start time is greater than or equal to today and start time is less than tomorrow)
				
				repeat with i from 1 to (count of calendarEvents)
					repeat with j from i + 1 to count of calendarEvents
						if start time of item j of calendarEvents < start time of item i of calendarEvents then
							set temp to item i of calendarEvents
							set item i of calendarEvents to item j of calendarEvents
							set item j of calendarEvents to temp
						end if
					end repeat
				end repeat
				
				repeat with i from 1 to (count of calendarEvents)
					set startTime to start time of item i of calendarEvents
					set endTime to end time of item i of calendarEvents
					set startTimeWithMargin to startTime - startTimeMargin
					set meet to item i of calendarEvents
					set s to subject of meet
					
					
					if (now is greater than or equal to startTimeWithMargin and now is less than endTime) then
						return meet
					end if
					
					-- If no meeting is found select the first one
					if (fallbackMeeting is "" and now is less than startTime) then
						set fallbackMeeting to meet
					end if
				end repeat
				
			end tell
			
		end repeat
		
		if fallbackMeeting is not "" then
			return fallbackMeeting
		else
			return defaultMeeting
		end if
	end tell
	
end getNextMeeting
