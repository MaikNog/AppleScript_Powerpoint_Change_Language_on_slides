# AppleScript_Powerpoint_Change_Language_on_slides
An Applescript for Mac OSX "Powerpoint for Mac" to change an existing powerpoint deck to a different language



# Author: Maik Nogens, 2018
# Inspired from : https://gist.github.com/low-decarie/4440405

# Works on Powerpoint for Mac v16.15
# Works on PPTX extension
# Does NOT work on POTM extension

tell application "Microsoft PowerPoint"
	activate
	
	set theView to view of document window 1
	
	repeat with slideNumber from 1 to count of slides of active presentation
		
		go to slide theView number slideNumber
		
		tell slide slideNumber of active presentation
			activate
			tell shapes
				select
			end tell
		end tell
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				tell menu bar 1
					# Change the entry to the shown name
					# Couldnt get it to work with German PPT and index numbering
					# Index Position could be different in each local OS language version of powerpoint
					tell menu bar item "Extras"
						tell menu "Extras"
							click menu item 5
						end tell
					end tell
				end tell
			end tell
			tell window 1 of process "Microsoft PowerPoint"
				tell scroll area 1
					tell table 1
						
						# I could not get this to work:
						#select static text "Englisch (USA)"  of UI element 1 of row 22 of table 1 of scroll area 1 of window "Sprache" 
						# So I counted the entry I wanted manually in the list:
						#Position 22 = Englisch (USA)
						# UI element = Position = here: 22
						select UI element 22
						
					end tell
				end tell
				click button "OK"
			end tell
		end tell
		
	end repeat
end tell
