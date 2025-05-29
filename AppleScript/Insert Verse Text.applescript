-- AppleScript droplet by Bernie Meier 07-May-2025
-- This app opens the dropped Word document to insert a few lines of the verse text 
-- from any BibleGateway verse hyperlink found in the document
-- it first changes the version reference from oldVersion to newVersion (see "make changes here")

-- if this app is started without any docs dropped onto it, it will prompt to choose a Word document

-- open ({"/Users/bernie/Desktop/test document.docx"})
set chosenFiles to choose file with prompt "Please select a file to process:" of type {"com.microsoft.word.doc", "org.openxmlformats.wordprocessingml.document", "*.doc", "*.docx", "*.docm"} with multiple selections allowed
open chosenFiles

on open droppedFiles
	tell application "Microsoft Word"
		
		-- only make changes here --
		set oldVersion to "NASB"
		set newVersion to "SCH2000"
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
									set rangeAfter to text object of thisField
									set startPos to (end of content of rangeAfter)
									set endPos to startPos + 5 + (count of verseText)
									set insertionPoint to collapse range rangeAfter direction collapse end
									insert text (" - \"" & verseText & "\"") at insertionPoint
									set newRange to create range theDoc start startPos end endPos
									set italic of font object of newRange to true
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
end open

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars
