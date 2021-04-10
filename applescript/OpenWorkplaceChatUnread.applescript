set clickUrl to "https://fb.workplace.com/chat"

tell application "Google Chrome"
	activate
	open location clickUrl
	delay 1
	activate
	
	# Wait until page finishes loading.
	repeat until (loading of active tab of front window is false)
		delay 1
	end repeat
	
	
	set js to "

$('[aria-label=\"Unread\"]').click();
$('[aria-label=\"Search Workplace Chat…\"]').click();

"
	set jquery_source to do shell script "curl http://code.jquery.com/jquery-latest.min.js"
	execute front window's active tab javascript jquery_source
	
	execute front window's active tab javascript js
	
	
	activate
	
	#tell application "System Events"
	#	keystroke "k" using {command down}
	#end tell
	
end tell
