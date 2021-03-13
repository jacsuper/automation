set display_text to "Please enter your password:"
repeat
	considering case
		set init_pass to text returned of (display dialog display_text default answer "" with hidden answer)
		set final_pass to text returned of (display dialog "Please verify your password below." buttons {"OK"} default button 1 default answer "" with hidden answer)
		if (final_pass = init_pass) then
			exit repeat
		else
			set display_text to "Mismatching passwords, please try again"
		end if
	end considering
end repeat

tell application "Finder"
	set theItems to selection
	set theItem to (item 1 of theItems) as alias
	set itemPath to quoted form of POSIX path of theItem
	set fileName to name of theItem
	set theFolder to POSIX path of (container of theItem as alias)
	set zipFile to quoted form of (fileName & ".zip")
	do shell script "cd '" & theFolder & "'; zip -x .DS_Store -r0 -P '" & final_pass & "' " & zipFile & " ./'" & fileName & "'"
end tell


