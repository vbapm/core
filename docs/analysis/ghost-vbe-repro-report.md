# VBE Ghost VBAProject Reproduction Report

## Summary

The ghost VBAProject is an intermittent Excel/VBE lifecycle problem. It is not currently attributable to VBAPM alone: the same workbook-open and workbook-close sequence has reproduced the ghost through ordinary Excel UI actions, a standalone PowerShell COM helper, and the installed VBAPM `close` command.

No production fix is included yet. Closing the VBE via the X button or `Alt+Q` returns to Excel by hiding the VBE window, but it doesn't destroy the VBE window or closes the workbook. A separate post-close VBE visibility refresh did clear the ghost in the direct COM test and remains the strongest current workaround candidate.

## Manual Excel reproduction

The same behavior has been observed through the ordinary Excel user interface, without using vbapm or an external COM helper:

1. Open Excel.
2. Open a new `Book1` document.
3. Open an existing Excel document from disk.
4. Close the existing document from disk.
5. Open or return to the VBE.

The closed document's VBA project can remain as a ghost in `VBE.VBProjects`. In other words, the ghost can remain visible in the VBE even after the workbook has been closed and the VBE is opened afterward. This manual observation further supports treating the issue as an Excel/VBE lifecycle problem rather than a vbapm-specific close implementation bug.

The same sequence has also reproduced the ghost when Excel is started in Safe Mode. In Safe Mode, open two existing workbooks because the initial `Book1` disappears when the first existing workbook is opened. This indicates that the behavior does not depend on Excel's normal startup add-ins or extension loadout.


**Environment observed during the manual test:** Microsoft Excel for Microsoft 365 MSO (Version 2609 Build 16.0.20430.20032) 64-bit.

## Automated reproduction with the PowerShell harness

Run the full harness from the repository root with Windows PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
	.\scripts\ps\repro-ghost-vbe-ci.ps1
```

**Warning:** The full harness closes all Excel instances at startup, saves unsaved workbooks through `close-excel.ps1 -SaveAll`, and force-stops any remaining Excel processes. Run it only when closing the current Excel session is safe.

The script deletes and recreates `ghost-demo/` in the repository root by default. Use `-WorkDir NAME` to choose another directory. It then:

1. Stops existing Excel instances.
2. Creates and builds a minimal `xlsm` project at `ghost-demo/build/ghost-demo.xlsm`, with codename `GhostDemoProject`.
3. Starts a fresh visible Excel instance.
4. Opens the VBE before opening the generated workbook.
5. Opens the generated workbook through COM.
6. Closes the workbook through the standalone COM helper.
7. Enumerates `VBE.VBProjects` and compares the projects with open workbooks.

To test the vbapm VBA close path instead, add `-UseCliClose`. This invokes the installed `%APPDATA%\vbapm\bin\vba.cmd close --force`, which calls `Build.CloseFile` inside the add-in:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
	.\scripts\ps\repro-ghost-vbe-ci.ps1 `
	-UseCliClose -WorkDir ghost-demo-vba
```

The harness prints `GHOST_COUNT_AFTER_WORKBOOK_CLOSE` and the final `GHOST_COUNT`. A nonzero count means that the detector found a VBE project with no matching open workbook. The optional `-HideVBEBeforeClose`, `-CloseVBEBeforeClose`, `-CloseAndReopenVBEAfterClose`, `-RefreshVBEAfterClose`, and `-HoldAfterDetection` switches add targeted experiments after the base reproduction.

## Detector-only check

To inspect the current VBE state without creating, opening, or closing a workbook, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
	.\scripts\ps\detect-ghost-vbe.ps1
```

This command only enumerates the current `VBE.VBProjects` collection. It does not reproduce the bug by itself.

## Findings

### Both close paths can reproduce the ghost

Repeated runs produced `GHOST_COUNT=1` in both modes:

| Close path | Result |
| --- | --- |
| PowerShell `Workbooks` lookup and `Workbook.Close` | Intermittently leaves one ghost project |
| Installed VBAPM `close --force` / add-in `Build.CloseFile` | Intermittently leaves one ghost project |

Some runs produced `GHOST_COUNT=0` in both modes. The intermittent behavior means a single clean run does not disprove the bug, and a single failing run does not establish that VBAPM is the cause.

### The ghost appears after a healthy open

The project is visible and responsive immediately after `Workbooks.Open`. After the workbook is closed, the project can remain in `VBE.VBProjects` even though the workbook is absent from `Application.Workbooks`.

The project name may appear as `VBAProject` when the installed add-in is stale relative to the worktree codename change. That is a fixture/build freshness issue, not evidence that the ghost is a different project.

### Ordinary COM release is not sufficient

The PowerShell helper was changed to use indexed workbook lookup and `FinalReleaseComObject` for the workbook and `Workbooks` collection. The ghost still occurred. This rules out the initially suspected simple PowerShell RCW lifetime leak as the complete explanation.

VBAPM's production close path calls the add-in's `Build.CloseFile`, which resolves the workbook and dispatches `Close` through `Application.Workbooks(...)`. The matching repro results do not show a VBAPM-specific object leak.

### Hiding the VBE is not sufficient

Setting `VBE.MainWindow.Visible = $false` immediately before closing the workbook did not prevent the ghost.

### Closing the VBE is not yet a validated solution

Calling `VBE.MainWindow.Close()` failed with `E_FAIL`. A Win32 test found the VBE HWND, called `SetForegroundWindow`, verified it with `GetForegroundWindow`, and sent both `WM_CLOSE` and `WM_SYSCOMMAND/SC_CLOSE`; neither message made the VBE window disappear. A later test used `ShowWindow`, `BringWindowToTop`, `AttachThreadInput`, `SetActiveWindow`, `SetFocus`, temporary topmost promotion, and verified `GetForegroundWindow()` before using 64-bit-correct `SendInput` to send `Alt+Q`. `SendInput` succeeded, but the VBE still did not close. This rules out simple misdirected keyboard focus as the explanation.

Manual validation also ruled out the normal VBE close-and-reopen workflow as a fix: closing the VBE with its window `X` and with the VBE menu command `Alt+Q` did not remove the ghost. Reopening the VBE showed the ghost project still present.

### Alt+Q and Alt+F4 hide the same live VBE window

The standalone `Send-AltQ-To-VBE.ps1` helper was extended with `-Key F4`, so both shortcuts use the same HWND discovery, foreground verification, and native `SendInput` path. The `Get-VBEWindowState.ps1` helper captured matching native and COM state before and after each shortcut.

For `Alt+F4`, the VBE was explicitly visible before the test:

| Property | Before | After |
| --- | --- | --- |
| HWND | `0x00630FD4` | `0x00630FD4` |
| Owner PID/thread | `53708` / `71644` | `53708` / `71644` |
| `IsWindow` | `True` | `True` |
| Visible | `True` | `False` |
| Window class | `wndclass_desked_gsk` | unchanged |
| Bounds | `1936x1048` | `1936x1048` |
| COM `VBE.MainWindow.Visible` | `True` | `False` |
| Excel workbook count | `1` | `1` |

The native style changed from `0x16CF0000` to `0x6CF0000`, which is the visible-style bit being removed. The HWND, window object, VBE COM object, and workbook all remain alive. A matching `Alt+Q` comparison produced the same transition. Therefore both shortcuts hide the VBE and return to Excel; they do not destroy the VBE window.

### Repeated Alt+F4 evidence does not show a protective effect

A corrected A/B batch ran the same direct-COM repro ten times without hiding the VBE and ten times with the standalone, foreground-verified `Alt+F4` sender before `Workbook.Close`. The treatment path used one quiet detector scan after the close; an earlier batch was discarded because the inline treatment path failed before closing the workbook and the detector was invoked twice.

| Condition | Valid runs | Ghost runs | Clean runs |
| --- | ---: | ---: | ---: |
| VBE left visible | 10 | 5 | 5 |
| VBE hidden with verified `Alt+F4` | 9 | 4 | 5 |

One Alt+F4 treatment run failed before producing a count and was excluded. Among valid runs, the observed ghost rates were 50% for the baseline and 44.4% for the Alt+F4 treatment. This small sample does not establish a statistically reliable difference, and it provides no evidence that hiding the VBE prevents the ghost. The current evidence therefore supports treating Alt+F4 as a UI visibility operation, not as a workaround for the project-tree bug.

### Refreshing the VBE after the workbook close clears the ghost

After a direct COM close produced `GHOST_COUNT_AFTER_WORKBOOK_CLOSE=1`, setting `VBE.MainWindow.Visible` to `$false`, waiting briefly, and setting it back to `$true` produced `GHOST_COUNT_AFTER_VBE_REFRESH=0` on the final detector scan. This directly answers the ordering question: refreshing the VBE after the workbook has already been closed can clear the ghost without quitting Excel.

The first post-refresh detector pass can still observe a transient stale entry, so the diagnostic now uses one quiet machine-readable scan for its reported result. The refresh is not proof that the VBE process or window was destroyed and recreated; it is a VBIDE visibility refresh that appears to force the project tree to be rebuilt.

## Likely ownership

The evidence points to Excel/VBE internal project-tree state rather than a live PowerShell or VBAPM workbook wrapper:

- The issue occurs with direct COM and VBAPM.
- It is intermittent under the same setup.
- Explicit release of the ordinary workbook-level COM objects does not eliminate it.
- Merely hiding the VBE does not eliminate it.

The remaining object graph may include VBE project objects, active code-window state, or Excel's internal VBE project cache. The public VBIDE object model does not expose a reliable remove/discard operation for a closed document project.

## Potential solutions, in order of confidence

1. **Avoid the triggering order.** Keep the VBE closed while VBAPM opens and closes workbooks. This is the least invasive option and should be tested as a control.
2. **Refresh the VBE after closing the workbook.** Set `VBE.MainWindow.Visible = $false`, wait briefly, then set it to `$true`. This cleared the ghost in the targeted direct-COM run and is less invasive than UI automation, but it needs repeated validation through both close paths. Normal VBE return through the window `X`, `Alt+Q`, or `Alt+F4` only hides the same VBE window and did not clear the ghost.
3. **Close the VBE through a verified UI action.** A Windows API implementation can first attach to the current foreground thread and promote the VBE with `SetForegroundWindow`, then attach to the VBE thread and call `SetFocus`, verify that the focused child belongs to the VBE frame, and verify `GetForegroundWindow` before using `SendInput` for `Alt+Q`. The final coordinated test passed both focus checks and delivered `Alt+Q`, but a fresh controlled rerun left `VBE.MainWindow.Visible=True` and the VBE window open. This is not a viable workaround in the current environment.
4. **Close the Excel instance after the workbook.** If the automation owns the Excel process, quitting Excel should discard the in-memory VBE project tree. This is unsuitable when the caller must preserve a user's Excel session, but it is a useful control experiment.
5. **Use a separate hidden Excel instance for automation.** This prevents a user's interactive VBE session from sharing the same project tree. It is likely the most practical production isolation strategy when the workflow permits it.
6. **Improve close cleanup.** Release every VBE/project/workbook COM wrapper in reverse acquisition order and avoid implicit `foreach` enumerators. This is good hygiene, but the current repro shows it is not sufficient by itself.

## Recommended next diagnostic

For the UI-close experiment, repeat without `SendKeys` ambiguity:

- capture the VBE HWND;
- call `SetForegroundWindow($vbeHwnd)`;
- verify `GetForegroundWindow()` equals that HWND;
- send `WM_SYSCOMMAND/SC_CLOSE` directly to that HWND, or use `SendInput` for Alt+F4;
- verify `VBE.MainWindow.Visible` becomes `$false` rather than requiring the HWND to be destroyed;
- close `ghost-demo.xlsm`;
- run the detector.

Run this several times against both direct COM close and VBAPM close. Only a repeatable zero-ghost result in both modes should be considered a candidate fix. Separately, repeat the post-close visibility refresh in both modes; the direct COM run already produced zero ghosts after that refresh, while the VBAPM run was clean before the refresh.

## CLI warning resolution

The dual-install warning was accurate on this machine: `vbapm@0.8.0` was installed globally under `C:\Program Files\nodejs`, ahead of the development install in `PATH`. That global package was removed with `npm uninstall -g vbapm`.

After removal:

- `vba` resolves only to `%APPDATA%\vbapm\bin`.
- The installed CLI reports `vbapm v0.10.0`.
- `vba --help` no longer prints the standalone/npm conflict warning.

The warning detection code was left intact because it still provides useful protection when both installations genuinely exist.
