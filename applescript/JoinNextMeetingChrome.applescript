set clickUrl to "https://www.internalfb.com/?tab=today"

set monitorPosition to {-1920, 0}

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

var arr = [], l = document.links;
for(var i=0; i<l.length; i++) {
	hr = l[i].href;
	if (hr.includes('groupcall') || hr.includes('bluejeans')){
		document.location = hr;
	}
}

"
	execute front window's active tab javascript js
	
	set bounds of front window to {-900, -1700, -200, -1100}
	
	activate
	
end tell

