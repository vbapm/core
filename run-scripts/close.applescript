-- Arguments
--
-- appname (e.g. "excel")
-- file: posix full path to the workbook
-- save: "1" to save before closing, "0" to discard changes (default: "0")

on run argv
	if (count of argv) < 2 then
		error "ERROR #1: Invalid Input (appname and file are required)"
	end if

	set appname to (item 1 of argv)
	set addin to POSIX file (item 2 of argv)
	set should_save to false
	if (count of argv) >= 3 then
		set should_save to (item 3 of argv) is "1"
	end if

	if appname is "excel" then
		set workbook_name to name of (info for addin)

		-- Excel is not running — file is already closed
		if application "Microsoft Excel" is not running then
			return
		end if

		tell application "Microsoft Excel"
			-- Workbook is not open — nothing to close
			if not (exists workbook workbook_name) then
				return
			end if

			if should_save then
				close workbook workbook_name saving yes
			else
				close workbook workbook_name saving no
			end if
		end tell
	else
		error "ERROR #3: Unsupported App"
	end if
end run
