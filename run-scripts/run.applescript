-- Arguments
--
-- appname (e.g. "excel")
-- addin: posix full path to addin (e.g. "...")
-- command: macro to execute in addin (e.g. "Build.ImportGraph", "Build.CloseFile")
-- For macros: arg 4 is keep_open: "1" to leave open after macro, "0" to close
-- ...args: arguments to pass to macro (up to 10)

-- -------
-- Feature parity notes (vs. run-scripts/run.ps1)
-- -------
--
-- This script is the macOS counterpart of run.ps1. Both are invoked the same
-- way from src/utils/run.ts: (appname, file, macro, ...args). The sections
-- below are laid out in the same order as run.ps1 so the two can be read
-- side by side. run.ps1 has grown several Windows-only concerns that this
-- file intentionally does NOT mirror:
--
--   * Multi-instance coordination registry (scripts/ps/Excel-InstanceRegistry.ps1),
--     the EXCEL.EXE instance-count safeguard, and VBA_DEBUG_INSTANCES logging.
--     NOT PORTABLE, by construction: Windows can launch any number of
--     independent `Excel.Application` COM processes (one per hidden/background
--     run), so concurrent agents need a registry to avoid stepping on each
--     other. macOS only ever has ONE "Microsoft Excel" application instance
--     system-wide — `tell application "Microsoft Excel"` always talks to that
--     single instance (launching it if needed). There is nothing to register,
--     count, or attach to; the "attach to an already-open workbook" case in
--     run.ps1 (Find-OpenWorkbook / Attach) is simply always true here.
--   * VBA_BACKGROUND_BUILD (hidden Excel instance for e2e runs). NOT PORTABLE
--     as a separate *instance*: Excel for Mac has no equivalent to a second,
--     invisible COM process. `visible` could be toggled on the single shared
--     app, but that would affect any Excel window the user has open too, so
--     it is deliberately left alone.
--   * Argument unescaping (PS `Unescape`, the `^q` -> `"` workaround). NOT
--     NEEDED here: src/utils/run.ts only shell-escapes args on Windows
--     (`env.isWindows ? escape(arg) : arg`); macOS args are passed to
--     `execFile("osascript", ...)` unescaped since there is no intermediate
--     shell to confuse.
--   * "Reveal window on failure" (PS Set/Clear-ActiveExcelInstance). Not
--     needed: on Mac there's no hidden instance to reveal — the single Excel
--     app a failed run touched is already the one the user can switch to.
--
-- Everything else (opening the workbook/add-in, running the macro with up to
-- 10 positional args, closing/quitting only what we opened, and formatting a
-- JSON error on failure) has a direct equivalent below.

-- -------
-- Helpers
-- -------

-- JSON-string escaping for error messages (the AppleScript analog of the
-- quoting `Fail` does in run.ps1 for its `{"success":false,"errors":[...]}`
-- payload). There is no separate Print/PrintLn/PrintErr here: `return output`
-- from `run` is what the caller reads as stdout.
on replace_chars(theText, searchStr, replaceStr)
	set AppleScript's text item delimiters to searchStr
	set textItems to text items of theText
	set AppleScript's text item delimiters to replaceStr
	set result to textItems as text
	set AppleScript's text item delimiters to ""
	return result
end replace_chars

-- -------
-- Run Macro
-- -------

-- Equivalent of run.ps1's `RunMacro`: Excel's scripting dictionary has no
-- variadic "splat" form, so each arg count needs its own explicit call.
on run_excel_macro(command, args)
	set result to ""

	try
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
	on error errMsg
		set escapedMsg to my replace_chars(errMsg, "\"", "\\\"")
		set result to "{\"success\":false,\"errors\":[\"" & escapedMsg & "\"]}"
	end try

	return result
end run_excel_macro

-- -------
-- Excel instance coordination registry
-- -------
--
-- N/A on macOS — see "Feature parity notes" above. There is exactly one
-- "Microsoft Excel" application to coordinate with, so nothing to register.

-- -------
-- Excel (open/run/close, dispatch, and argument validation)
-- -------
--
-- AppleScript has no classes and no separate `Run`/main entry point, so the
-- lifecycle that run.ps1 spreads across the `Excel` class, the `Run`
-- function, and the bottom-of-file Main block is all handled inline in this
-- one `on run` handler:
--
--   * Argument count check                 <-> Main's `-not $AppName -or ...` / MacroArgs.Count check
--   * `if appname is "excel"`               <-> Run's `switch ($AppName) { "excel" { ... } }`
--   * open workbook/add-in if not already   <-> Excel.OpenExcel + Excel.OpenWorkbook
--   * `run my run_excel_macro(...)`         <-> Excel.Run (which calls RunMacro)
--   * close workbook / quit Excel           <-> Excel.Dispose
on run argv
	set output to ""

	if (count of argv) >= 4 and (count of argv) <= 14 then
		set appname to (item 1 of argv)
		set addin to POSIX file (item 2 of argv)
		set command to (item 3 of argv)
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
	else
		if (count of argv) < 4 then
			set output to "ERROR #1: Invalid Input (appname, file, macro, and keep_open are required)"
		else
			set output to "ERROR #2: Invalid Input (only 10 arguments are supported)"
		end if
	end if

	return output
end run
