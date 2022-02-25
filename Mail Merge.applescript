on enabledGUIScripting(flag) -- https://gist.github.com/iloveitaly/2ff08138091afd69cf2b
	get system attribute "sysv"
	if result is less than 4240 then -- 4240 is OS X 10.9.0 Mavericks
		
		-- In OS X 10.8 Mountain Lion and older, enable GUI Scripting globally by calling this handler and passing 'true' in the flag parameter before your script executes any GUI Scripting commands, or pass 'false' to disable GUI Scripting. Authentication is required only if the value of the 'UI elements enabled' property will be changed. Returns the final setting of 'UI elements enabled' even if unchanged.
		
		tell application "System Events"
			activate -- brings System Events authentication dialog to front
			set UI elements enabled to flag
			return UI elements enabled
		end tell
	else
		
		-- In OS X 10.9 Mavericks, global access is no longer available and GUI Scripting can only be enabled manually on a per-application basis. Pass true to present an alert with a button to open System Preferences and telling the user to select the current application (the application running the script) in the Accessibility list in the Security & Privacy preference's Privacy pane. Authentication is required to unlock the preference. Returns the current setting of 'UI elements enabled' asynchronously, without waiting for System Preferences to open, and tells the user to run the script again.
		
		tell application "System Events" to set GUIScriptingEnabled to UI elements enabled -- read-only in OS X 10.9 Mavericks
		if flag is true then
			if not GUIScriptingEnabled then
				activate
				set scriptRunner to name of current application
				display alert "You must change some settings to allow Mail Merge to run." message "You must unlock the Security & Privacy preference pane (by clicking the lock icon), check the box next to \"" & scriptRunner & "\" in the Accessibility list, and then run this script again." buttons {"Cancel", "Open System Preferences"} default button "Open System Preferences"
				if button returned of result is "Open System Preferences" then
					tell application "System Preferences"
						tell pane id "com.apple.preference.security" to reveal anchor "Privacy_Accessibility"
						activate
					end tell
				end if
			end if
		end if
		return GUIScriptingEnabled
	end if
end enabledGUIScripting

on isRunning(appName)
	tell application "System Events" to return (name of processes) contains appName
end isRunning

on checkForCorrectUsage()
	set usageErrorMessage to "Mail merge requires both a template (an open Pages document) and a data source (an open Numbers document).

Write \"%Column Name%\" in Pages to insert data from a Numbers column whose top cell reads \"Column Name\"."
	
	set numbersFormatErrorMessage to "Mail merge requires that your data is in the first table of the first sheet of your open Numbers document.

Create your table with one header row and no header columns, then write \"%Column Name%\" in Pages to insert data from a column whose top cell reads \"Column Name\"."
	
	if not isRunning("Numbers") and not isRunning("Pages") then
		display alert "Welcome to mail merge! Please open a Pages document and a Numbers document, then run this script again." message usageErrorMessage
		return false
	end if
	if not isRunning("Numbers") then
		display alert "Welcome to mail merge! Please open a Numbers document and run this script again." message usageErrorMessage
		return false
	end if
	if not isRunning("Pages") then
		display alert "Welcome to mail merge! Please open a Pages document and run this script again." message usageErrorMessage
		return false
	end if
	
	tell application "Pages"
		if (count of documents) < 1 then
			display alert "Welcome to mail merge! Please open a Pages document and run this script again." message usageErrorMessage
			return false
		end if
	end tell
	
	tell application "Numbers"
		if (count of documents) < 1 then
			display alert "Welcome to mail merge! Please open a Numbers document and run this script again." message usageErrorMessage
			return false
		end if
		if (count of sheets of document 1) < 1 then
			display alert "Welcome to mail merge! Please create a sheet in your Numbers document and run this script again." message numbersFormatErrorMessage
			return false
		end if
		if (count of tables of sheet 1 of document 1) < 1 then
			display alert "Welcome to mail merge! Please create a table in the first sheet of your Numbers document and run this script again." message numbersFormatErrorMessage
			return false
		end if
	end tell
	
	return true
end checkForCorrectUsage

on waitForProcessing()
	delay 0.1
	repeat while (do shell script "/bin/ps -xco %cpu,command | /usr/bin/awk '/Pages$/ {print $1}'") > 1
		delay 0.1
	end repeat
	delay 0.1
end waitForProcessing

on waitForClipboardToChangeFrom(oldValue)
	repeat while (the clipboard) is equal to oldValue
		delay 0.1
	end repeat
end waitForClipboardToChangeFrom

on waitForPagesWindowToChangeFrom(oldName)
	tell application "System Events"
		tell process "Pages"
			repeat while title of first window is oldName
				delay 0.1
			end repeat
		end tell
	end tell
end waitForPagesWindowToChangeFrom

on waitForEnabledWithTimeout(interfaceElement, timeoutSeconds)
	set success to true
	set deadline to (current date) + timeoutSeconds
	tell application "System Events"
		tell process "Pages"
			repeat while (enabled of interfaceElement) is false
				delay 0.2
				if (current date) > deadline then
					set success to false
					exit repeat
				end if
			end repeat
		end tell
	end tell
	
	return success
end waitForEnabledWithTimeout

on waitForDisabled(interfaceElement)
	tell application "System Events"
		tell process "Pages"
			repeat while (interfaceElement is enabled)
				delay 0.1
			end repeat
		end tell
	end tell
end waitForDisabled

on waitForPagesToScrollToLastPage()
	tell application "Pages"
		repeat while current page of document 1 is not last item of pages of document 1
			delay 0.1
		end repeat
	end tell
end waitForPagesToScrollToLastPage

on average(valuesList)
	set sum to 0
	repeat with currentItem in valuesList
		set sum to sum + currentItem
	end repeat
	return (sum / (count of valuesList))
end average

on roundToTwoDecimals(unroundedValue)
	return ((round (unroundedValue * 100)) / 100)
end roundToTwoDecimals

on keystrokeCorrectly(someString)
	set AppleScript's text item delimiters to {return & linefeed, return, linefeed, character id 8233, character id 8232}
	set escapedString to text items of (someString as string)
	set AppleScript's text item delimiters to "\\n"
	set escapedString to escapedString as text
	
	repeat with currentChar in the characters of (escapedString as string)
		delay 0.01
		tell application "System Events" to keystroke currentChar
	end repeat
end keystrokeCorrectly

on run
	if my enabledGUIScripting(true) is false then return
	
	set fieldDelimiter to "%"
	set unusedFieldTimeout to 2
	set indexOfReplaceButton to 4
	set backupIndexOfReplaceButton to 1
	
	if not checkForCorrectUsage() then return
	
	tell application "Numbers"
		set numbersDocument to name of document 1
		set tableName to name of table 1 of sheet 1 of document 1
		set entries to value of cells of rows 2 thru -1 of table 1 of sheet 1 of document 1
		set fields to value of cells of row 1 of table 1 of sheet 1 of document 1
	end tell
	
	tell application "Pages"
		set pagesDocument to name of document 1
		set pagesDocumentHasBody to document body of document 1
		display dialog "Ready to merge data from table \"" & tableName & "\" in \"" & numbersDocument & "\" into \"" & pagesDocument & "\".

Don't interact with your computer during the merge." with icon note
	end tell
	
	set scriptStartTime to current date
	
	tell application "Pages" to activate
	tell application "System Events"
		tell process "Pages"
			-- Check setup
			display notification "Checking setup..." with title "Mail Merge" subtitle "Preparing for merge..."
			set oldWindow to title of first window
			keystroke "f" using command down
			my waitForPagesWindowToChangeFrom(oldWindow)
			try
				button indexOfReplaceButton of window 1
			on error
				click menu button 1 of window 1
				delay 0.5
				key code 125 -- down arrow
				delay 0.5
				key code 125 -- down arrow
				delay 0.5
				keystroke return
			end try
			if title of button indexOfReplaceButton of window 1 is missing value then
				set indexOfReplaceButton to backupIndexOfReplaceButton
			end if
			
			set oldWindow to title of first window
			keystroke "w" using command down
			my waitForPagesWindowToChangeFrom(oldWindow)
			
			-- Copy template
			display notification "Copying template..." with title "Mail Merge" subtitle "Preparing for merge..."
			-- Toggle page thumbnails and/or inspector (in case something in them is selected, which will mess up our upcoming "Select All")
			keystroke "p" using command down & option down
			keystroke "i" using command down & option down
			keystroke "p" using command down & option down
			keystroke "i" using command down & option down
			-- Select template
			keystroke "a" using command down
			my waitForProcessing()
			set the clipboard to "temp"
			keystroke "c" using command down
			my waitForClipboardToChangeFrom("temp")
			
			-- Duplicate document and delete contents
			display notification "Creating merged document..." with title "Mail Merge" subtitle "Preparing for merge..."
			keystroke "s" using command down & shift down
			my waitForProcessing()
			keystroke return
			my waitForProcessing()
			key code 51 -- delete
			
			-- Create merged contents
			set skippedFields to {}
			set entryTimes to {}
			set replacementCount to 0
			set entryCount to 0
			repeat with entryIndex from 1 to count of entries
				set beforeEntryTime to current date
				
				-- Paste template
				keystroke "v" using command down
				
				-- Replace fields
				set oldWindow to title of first window
				keystroke "f" using command down
				my waitForPagesWindowToChangeFrom(oldWindow)
				
				-- Display progress to user
				set notificationSubtitle to "Merging entry " & entryIndex & " of " & (count of entries)
				set notificationMessage to ""
				
				if (count of entryTimes) > 0 then
					set remainingTime to (my average(entryTimes)) * ((count of entries) - (entryIndex - 1)) / 60
					set notificationMessage to "Remaining time: " & my roundToTwoDecimals(remainingTime) & " minutes"
				end if
				display notification notificationMessage with title "Mail Merge" subtitle notificationSubtitle
				
				repeat with fieldIndex from 1 to count of fields
					
					-- Meat & potatoes
					if skippedFields does not contain fieldIndex then
						set currentValue to item fieldIndex of item entryIndex of entries
						
						my keystrokeCorrectly(fieldDelimiter & item fieldIndex of fields & fieldDelimiter)
						keystroke tab
						key code 51 -- delete
						if currentValue is not missing value then my keystrokeCorrectly(currentValue)
						
						if my waitForEnabledWithTimeout(button indexOfReplaceButton of window 1, unusedFieldTimeout) then
							click button indexOfReplaceButton of window 1
							my waitForDisabled(button indexOfReplaceButton of window 1) -- button is disabled when replacement is done
							set replacementCount to replacementCount + 1
						else
							copy fieldIndex to end of skippedFields
						end if
						keystroke tab
					end if
				end repeat
				set oldWindow to title of first window
				keystroke "w" using command down
				my waitForPagesWindowToChangeFrom(oldWindow)
				
				-- If needed, create new page
				if not pagesDocumentHasBody then tell document 1 of application "Pages" to make new page
				
				-- If needed, move insertion point to end of document
				if pagesDocumentHasBody then
					keystroke "a" using command down
					key code 124 -- right arrow
				end if
				
				-- If needed, scroll to end of document
				if not pagesDocumentHasBody then
					key code 119 -- end
					my waitForPagesToScrollToLastPage()
				end if
				
				set afterEntryTime to current date
				set totalEntryTime to afterEntryTime - beforeEntryTime
				if (entryIndex > 1) or ((count of skippedFields) is 0) then
					copy totalEntryTime to end of entryTimes
				end if
			end repeat
		end tell
	end tell
	
	set scriptEndTime to current date
	set scriptRunningTime to (scriptEndTime - scriptStartTime) / 60
	
	tell application "Pages" to display alert "Merge complete!" message "Made " & replacementCount & " replacements over " & (count of entries) & " entries. Total running time: " & my roundToTwoDecimals(scriptRunningTime) & " minutes" buttons {"OK"} default button "OK"
end run
