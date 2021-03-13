tell application "Google Chrome"
	activate
	repeat with w in (windows)
		set j to 0
		repeat with t in (tabs of w)
			set j to j + 1
			set isWorkplaceTab to title of t contains "Workplace room"
			set isBlueJeansInBrowserTab to title of t contains "Blue Jeans" and URL of t contains "webrtc"
			
			if isWorkplaceTab or isBlueJeansInBrowserTab then
				set (active tab index of w) to j
				set index of w to 1
				tell application "System Events" to tell process "Google Chrome"
					perform action "AXRaise" of window 1 -- `set index` doesn't always raise the window
				end tell
				return
			end if
			
		end repeat
	end repeat
end tell

