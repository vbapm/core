-- Arguments
--
-- appname (e.g. "excel")
-- addin: posix full path to addin (e.g. "...")
-- command: macro to execute (e.g. "Build.ImportGraph"), or "close" to close the workbook
-- For "close": arg 4 is save flag: "1" to save, "0" to discard (default)
-- For macros: arg 4 is keep_open: "1" to leave open after macro, "0" to close
-- ...args: arguments to pass to macro (up to 10)

on run argv
	set output to ""

	if (count of argv) >= 3 then
		set appname to (item 1 of argv)
		set addin to POSIX file (item 2 of argv)
		set command to (item 3 of argv)

		if command is "close" then
			set should_save to false
			if (count of argv) >= 4 then
				set should_save to (item 4 of argv) is "1"
			end if

			if appname is "excel" then
				set workbook_name to name of (info for addin)

				if application "Microsoft Excel" is not running then
					set output to "{\"success\":true,\"messages\":[\"File is not open\"]}"
				else
					tell application "Microsoft Excel"
						if not (exists workbook workbook_name) then
							set output to "{\"success\":true,\"messages\":[\"File is not open\"]}"
						else if should_save then
							close workbook workbook_name saving yes
							set output to "{\"success\":true}"
						else
							close workbook workbook_name saving no
							set output to "{\"success\":true}"
						end if
					end tell
				end if
			else
				set output to "ERROR #3: Unsupported App"
			end if

		else if (count of argv) >= 4 and (count of argv) <= 14 then
			set keep_open to (item 4 of argv) is "1"

			set args to {}
			repeat with index from 5 to count of argv
				set end of args to (item index of argv)
			end repeat

			if appname is "excel" then
				set workbook_name to name of (info for addin)

				set excel_was_open to application "Microsoft Excel" is running

				tell application "Microsoft Excel"
					set workbook_was_open to (exists workbook workbook_name)
					if not workbook_was_open then
						open workbook workbook file name addin without notify
					end if

					set output to my run_excel_macro(command, args)

					-- A workbook that was already open is never closed by us.
					-- Otherwise close it unless keep_open was requested.
					if not workbook_was_open and not keep_open then
						close workbook workbook_name saving yes
					end if
				end tell

				-- Quit Excel only if we launched it AND we are not keeping the file open.
				if not excel_was_open and not keep_open then
					tell application "Microsoft Excel" to quit
				end if
			end if

		end if

	else
		if (count of argv) < 3 then
			set output to "ERROR #1: Invalid Input (appname, file, and command are required)"
		else
			set output to "ERROR #2: Invalid Input (only 10 arguments are supported)"
		end if
	end if

	return output
end run

on run_excel_macro(command, args)
	set result to ""

	tell application "Microsoft Excel"
		if (count of args) = 0 then
			set result to result & (run VB macro command)
		else if (count of args) = 1 then
			set result to result & (run VB macro command arg1 (item 1 of args))
		else if (count of args) = 2 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args))
		else if (count of args) = 3 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args))
		else if (count of args) = 4 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args))
		else if (count of args) = 5 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args))
		else if (count of args) = 6 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args) arg6 (item 6 of args))
		else if (count of args) = 7 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args) arg6 (item 6 of args) arg7 (item 7 of args))
		else if (count of args) = 8 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args) arg6 (item 6 of args) arg7 (item 7 of args) arg8 (item 8 of args))
		else if (count of args) = 9 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args) arg6 (item 6 of args) arg7 (item 7 of args) arg8 (item 8 of args) arg9 (item 9 of args))
		else if (count of args) = 10 then
			set result to result & (run VB macro command arg1 (item 1 of args) arg2 (item 2 of args) arg3 (item 3 of args) arg4 (item 4 of args) arg5 (item 5 of args) arg6 (item 6 of args) arg7 (item 7 of args) arg8 (item 8 of args) arg9 (item 9 of args) arg10 (item 10 of args))
		end if
	end tell

	return result
end run_macro
