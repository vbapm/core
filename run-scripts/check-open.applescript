on run argv
	set targetPath to item 1 of argv
	if application "Microsoft Excel" is not running then return "false"

	tell application "Microsoft Excel"
		if not visible then return "false"
		repeat with workbookReference in workbooks
			try
				set workbookPath to POSIX path of ((full name of workbookReference) as alias)
				if workbookPath is targetPath then return "true"
			end try
		end repeat
	end tell

	return "false"
end run
