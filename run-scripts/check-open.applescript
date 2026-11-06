on replaceText(theText, searchText, replacementText)
	set AppleScript's text item delimiters to searchText
	set textItems to text items of theText
	set AppleScript's text item delimiters to replacementText
	set replacedText to textItems as text
	set AppleScript's text item delimiters to ""
	return replacedText
end replaceText

on jsonEscape(theText)
	set backslash to ASCII character 92
	set escapedText to my replaceText(theText as text, backslash, backslash & backslash)
	set escapedText to my replaceText(escapedText, quote, backslash & quote)
	set escapedText to my replaceText(escapedText, return, backslash & "r")
	set escapedText to my replaceText(escapedText, linefeed, backslash & "n")
	set escapedText to my replaceText(escapedText, tab, backslash & "t")
	return escapedText
end jsonEscape

on jsonBoolean(theValue)
	if theValue then return "true"
	return "false"
end jsonBoolean

on run argv
	set targetPath to item 1 of argv
	set targetName to name of (info for (POSIX file targetPath)) as text
	set instanceRecords to {}

	if application "Microsoft Excel" is running then
		tell application "Microsoft Excel"
			set excelVisible to visible
			repeat with workbookReference in workbooks
				try
					set workbookPath to POSIX path of ((full name of workbookReference) as alias)
					set workbookName to name of workbookReference as text
					if workbookName is targetName then
						set readOnlyValue to missing value
						try
							set readOnlyValue to read only of workbookReference
						end try

						set accessJSON to quote & "unknown" & quote
						if readOnlyValue is not missing value then
							if readOnlyValue then
								set accessJSON to quote & "read_only" & quote
							else
								set accessJSON to quote & "read_write" & quote
							end if
						end if

						set isSamePath to false
						considering case
							set isSamePath to workbookPath is targetPath
						end considering

						set recordJSON to "{" & quote & "path" & quote & ":" & quote & my jsonEscape(workbookPath) & quote
						set recordJSON to recordJSON & "," & quote & "same_path" & quote & ":" & my jsonBoolean(isSamePath)
						set recordJSON to recordJSON & "," & quote & "source" & quote & ":" & quote & "workbook" & quote
						set recordJSON to recordJSON & "," & quote & "visible" & quote & ":" & my jsonBoolean(excelVisible)
						set recordJSON to recordJSON & "," & quote & "access" & quote & ":" & accessJSON
						set recordJSON to recordJSON & "}"
						set end of instanceRecords to recordJSON
					end if
				end try
			end repeat
		end tell
	end if

	set AppleScript's text item delimiters to ","
	set recordsJSON to instanceRecords as text
	set AppleScript's text item delimiters to ""
	return "{" & quote & "instances" & quote & ":[" & recordsJSON & "]," & quote & "uninspected_excel_process_ids" & quote & ":[]}"
end run
