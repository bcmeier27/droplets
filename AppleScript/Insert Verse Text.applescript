-- AppleScript droplet by Bernie Meier 07-May-2025
-- Version 2.1 - Copilot suggested change to on open to ensure it runs on main thread
-- This app opens the dropped Word document to insert a few lines of the verse text 
-- from any BibleGateway verse hyperlink found in the document
-- it first changes the version reference from oldVersion to newVersion (see "make changes here")

-- if this app is started without any docs dropped onto it, it will prompt to choose a Word document

use AppleScript version "2.4" -- Yosemite (10.10) or later
use framework "Foundation"
use framework "AppKit"
use scripting additions

property NSAlert : a reference to current application's NSAlert
property NSTextField : a reference to current application's NSTextField
property NSButton : a reference to current application's NSButton
property NSOnState : a reference to current application's NSOnState

on run
	-- Set up AppKit references
	--set alert to NSAlert's alloc()'s init()
	
	-- open ({"/Users/bernie/Desktop/test document.docx"})
	set chosenFiles to choose file with prompt "Please select a file to process:" of type {"com.microsoft.word.doc", "org.openxmlformats.wordprocessingml.document", "*.doc", "*.docx", "*.docm"} with multiple selections allowed
	
	open chosenFiles
	
end run

on open droppedFiles
	(current application's NSThread's mainThread())'s performSelectorOnMainThread_withObject_waitUntilDone_("processFilesOnMainThread:", droppedFiles, true)
end open

on processFilesOnMainThread_(droppedFiles)
	-- Create alert window
	set alert to NSAlert's alloc()'s init()
	alert's setMessageText:"Bible Verse Hyperlink Insertion (MS Word)"
	alert's setInformativeText:"Please enter values for Source and Target Bible versions, and select how to process the links."
	
	-- Create input fields
	set l1 to NSTextField's alloc()'s initWithFrame:{{0, 90}, {300, 24}}
	l1's setStringValue:"Source:"
	l1's setEditable:false
	l1's setSelectable:false
	l1's setBezeled:false
	l1's setDrawsBackground:false
	
	set inputFieldA to NSTextField's alloc()'s initWithFrame:{{60, 90}, {350, 24}}
	inputFieldA's setStringValue:"NASB"
	
	set l2 to NSTextField's alloc()'s initWithFrame:{{0, 60}, {300, 24}}
	l2's setStringValue:"Target:"
	l2's setEditable:false
	l2's setSelectable:false
	l2's setBezeled:false
	l2's setDrawsBackground:false
	
	set inputFieldB to NSTextField's alloc()'s initWithFrame:{{60, 60}, {350, 24}}
	inputFieldB's setStringValue:"SCH2000"
	
	-- Create checkbox
	set checkbox1 to NSButton's alloc()'s initWithFrame:{{0, 30}, {300, 24}}
	checkbox1's setButtonType:(current application's NSSwitchButton)
	checkbox1's setTitle:"Insert verse text (quoted, after hyperlink)"
	checkbox1's setState:1
	
	-- Create checkbox
	set checkbox2 to NSButton's alloc()'s initWithFrame:{{0, 0}, {300, 24}}
	checkbox2's setButtonType:(current application's NSSwitchButton)
	checkbox2's setTitle:"Insert Screen Tip (when hovering over link)"
	checkbox2's setState:0
	
	-- Add controls to a view
	set theView to current application's NSView's alloc()'s initWithFrame:{{0, 0}, {500, 120}}
	theView's addSubview:l1
	theView's addSubview:inputFieldA
	theView's addSubview:l2
	theView's addSubview:inputFieldB
	theView's addSubview:checkbox1
	theView's addSubview:checkbox2
	
	alert's setAccessoryView:theView
	alert's addButtonWithTitle:"OK"
	alert's addButtonWithTitle:"Cancel"
	
	-- Show the dialog
	set response to alert's runModal()
	if response = (current application's NSAlertSecondButtonReturn) then
		display alert "User canceled."
		return
	end if
	
	-- Get values
	set stringA to (inputFieldA's stringValue()) as text
	set stringB to (inputFieldB's stringValue()) as text
	set insertVerseText to (1 = (checkbox1's state()))
	set insertScreenTip to (1 = (checkbox2's state()))
	
	-- Optional: Show result
	--display dialog "StringA: " & stringA & return & "StringB: " & stringB & return & "Insert verse text: " & insertVerseText & return & "Insert screen tip: " & insertScreenTip buttons {"OK"} default button "OK"
	
	tell application "Microsoft Word"
		
		-- only make changes here --
		--set oldVersion to "NASB"
		--set newVersion to "SCH2000"
		set oldVersion to stringA
		set newVersion to stringB
		-- end of changes ----------
		
		repeat with aFile in droppedFiles
			open aFile
			set theDoc to active document
			set docName to name of theDoc
			set fieldList to hyperlink objects of theDoc -- this returns the actual hyperlink items
			set linkCount to 0
			if (count fieldList) is not 0 then
				set fieldCount to length of fieldList
				repeat with a from 1 to fieldCount
					try
						set thisField to item a of fieldList
						if (type of thisField) is hyperlink then
							set theURL to hyperlink address of thisField
							if theURL is not missing value then -- comparing against empty or "" didn't work
								if theURL starts with "http" then
									set theText to text to display of thisField
									set verseText to ""
									
									-- replace "NASB" with "SCH2000" in the URL
									set theURL to my replace_chars(theURL, oldVersion, newVersion)
									set the hyperlink address of thisField to theURL
									
									-- fetch the verse text from the URL hyperlink
									try
										set curlCommand to "/usr/bin/curl -s " & quoted form of theURL
										set html to do shell script curlCommand
										
										-- Extract content of og:description in the HTML response
										set tid to AppleScript's text item delimiters -- keep track of the delims
										set AppleScript's text item delimiters to "<meta property=\"og:description\" content=\""
										set tmp to text items of html -- should return a tuple of text before the delimiter and text after
										if (count of tmp) > 1 then
											set remainder to item 2 of tmp
											set AppleScript's text item delimiters to "\""
											set verseText to item 1 of text items of remainder
										end if
										set AppleScript's text item delimiters to tid -- set them back to what they were
										-- 
									on error errMsg
										display alert "Error while fetching HTML content: " & errMsg
									end try
									set linkCount to linkCount + 1
									
									-- now insert fetched text into the document after the hyperlink's display text
									if (insertVerseText) then
										set rangeAfter to text object of thisField
										set startPos to (end of content of rangeAfter)
										set endPos to startPos + 5 + (count of verseText)
										set insertionPoint to collapse range rangeAfter direction collapse end
										insert text (" - \"" & verseText & "\"") at insertionPoint
										set newRange to create range theDoc start startPos end endPos
										set italic of font object of newRange to true
									end if
									
									if (insertScreenTip) then
										set the screen tip of thisField to verseText
									end if
								end if
							end if
						else
							display alert "Field " & a & " is not a hyperlink" as warning
						end if
					on error errorText number errNum
						display alert errorText
					end try
				end repeat
			else
				display alert "Document has no links." as warning
			end if
			-- close theDoc saving no
			display alert "Document processing complete!
(" & docName & ")

URL hyperlinks fetched: " & linkCount & "
out of total links found: " & fieldCount & "

*** DON`T FORGET TO SAVE THIS DOCUMENT! ***"
		end repeat
	end tell
end processFilesOnMainThread_

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars
