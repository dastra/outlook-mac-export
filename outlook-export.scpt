tell application "Microsoft Outlook"
	activate
	delay 0.3
end tell
tell application "System Events"
	tell process "Microsoft Outlook"
		set {year:y, month:m, day:d, time:t} to (current date)
		set m to m * 1
		set m to text -1 thru -2 of ("0" & m) 
		set theExportAsName to "Outlook for Mac Archive " & y & m & d & ".olm"
		set theExportToFolder to "~/Documents/Outlook Archives"

		--input the file name and the destination folder path you want to export as
		tell menu item "Export..." of menu "File" of menu bar item "File" of menu bar 1
			click
			delay 0.5
		end tell

		--show the Export dialog. We need to deal with three dialogs totally.
		--knowledge point 1: how to click menu item in APP menu bar
		tell group 1 of group "What do you want to export?" of front window
			if value of radio button "Items of these types:" is not 1 then
				click radio button "Items of these types:"
			end if
		end tell

		-- choose export items
		tell group 1 of group 1 of group "What do you want to export?" of front window
			if value of checkbox "Emails" is not 1 then
				click checkbox "Emails"
			end if
			if value of checkbox "Calendar" is not 1 then
				click checkbox "Calendar"
			end if
			if value of checkbox "Contacts " is not 1 then
				click checkbox "Contacts "
			end if
		end tell
		tell group 2 of group 1 of group "What do you want to export?" of front window
			if value of checkbox "Notes" is not 1 then
				click checkbox "Notes"
			end if
			if value of checkbox "Tasks" is not 1 then
				click checkbox "Tasks"
			end if
		end tell

		-- click Continue button to go to next dialog window (2/3)
		click button "Continue" of front window
		delay 0.5

		tell text field 1 of sheet 1 of front window
			set value to theExportAsName
		end tell

		-- set the file name we want to use
		set the clipboard to theExportToFolder
		keystroke "G" using {command down, shift down}
		delay 0.5
		keystroke "v" using {command down}
		delay 0.5
		keystroke return
		delay 0.5

		-- set the destination folder path we want to use
		tell sheet 1 of front window
			click button "Save"
			delay 0.5
		end tell
		-- click Save button to go to the last dialog window (3/3)
		tell front window
			delay 30
			click button "Finish"
		end tell
		-- click Finish button to finish this task
	end tell
end tell
