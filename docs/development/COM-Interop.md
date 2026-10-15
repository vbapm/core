# Excel COM lookup on Windows

## PowerShell runtime

Excel/VBE automation is intended to run under Windows PowerShell 5.1
(`powershell.exe`). PowerShell 7 (`pwsh`) may invoke compatible entry points,
but those scripts re-launch the COM work under Windows PowerShell. CI steps
that use `pwsh` for non-COM tests do not change this requirement.

## How to get the only Excel instance?

When there is exactly one running Excel instance, this can attach to it:

```powershell
[System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
```

`Marshal.GetActiveObject` returns the first object registered in the COM Running Object Table
(ROT). It does not enumerate Excel processes, accept a process ID, or create a
new Excel instance.

Avoid using the following:

```powershell
[Microsoft.VisualBasic.Interaction]::GetObject($null, "Excel.Application")
```

Based on experience, this will create a new instance of Excel, probably because there is some interoperability issues with `GetObject($null, "Excel.Application")` in powershell not being the same as `GetObject(, "Excel.Application")` in VB.NET. See `demo-vb-getobject.ps1`.

## Multiple Excel instances

When multiple Excel processes exist, `GetActiveObject` always returns the oldest Excel instance (assuming that they were not registered too close together and a race condition caused a different instance to be returned which has never occured to me so far).

The `Hwnd` property identifies the returned Excel window. Map that handle to a
process ID when diagnosing which instance was returned. Different Hwnd values can represent different Excel windows in the same process. Use the window handle to resolve the owning PID before treating it as an instance identifier.

## Creating a new instance

If the caller needs a specific new instance, create it and keep the returned
COM object. Perform all workbook operations through that object. Do not create
Excel with `Start-Process` and then call `GetActiveObject` to find it, because
that can attach to another process.

```powershell
$excelType = [System.Type]::GetTypeFromProgID("Excel.Application")
$excel = [System.Activator]::CreateInstance($excelType)
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
```

The caller owns the new instance. Release child COM references, release the
application reference when finished, and call `$excel.Quit()` when the caller
intends to close Excel.

`Create-NewExcel.ps1` follows this pattern. It also records the exact PID and
workbooks in `%TEMP%\Excel-Instances\instances.json`.

## Registry and workbook listings

`List-ExcelInstances.ps1` synchronizes the registry before reading it. A
process can be visible in the registry while its workbooks are unknown if the
process is not reachable through the ROT. Newly created instances avoid this
gap because `Create-NewExcel.ps1` registers their workbook list using the
exact COM object before releasing it.

## Information available without COM

When an Excel process is running but not COM-reachable, a PID and window handle
still provide useful process-level information:

- Process ID (`PID`)
- Process name, normally `EXCEL`
- Main window handle (`MainWindowHandle`)
- Main window title (`MainWindowTitle`)
- A visibility estimate based on whether `MainWindowHandle` is nonzero
- Process start time
- Process executable path and command line, when Windows permits access
- Parent process ID, when available through process inspection
- Registry metadata already recorded for that PID, such as owner, reason,
  creation time, and previously recorded workbook names
- Titled top-level Excel windows and their owning handles
- Titled `EXCEL7` child windows, which are a useful proxy for workbook names!

## Reproduce the behavior

Run:

```powershell
.\scripts\ps\demo-vb-getobject.ps1
```

The demo prints the selected process ID, window handle, visibility, and open
workbooks. Run it twice and compare the process IDs. The demo may leave hidden
Excel processes behind on systems with this COM registration behavior.

Clean them up without closing visible Excel:

```powershell
.\scripts\ps\Close-AllInvisibleExcelInstances.ps1
```

## Guidance

- Do not assume that either VB.NET `GetObject` or VBA `GetObject` returns the
  visible Excel instance when multiple Excel processes exist.
- Do not use a lookup API to choose among multiple Excel processes. Use the
  exact COM object returned during creation, or verify a known `Hwnd` and PID
  before mutating an existing instance.
- Use `Marshal.GetActiveObject` from Windows PowerShell only to retrieve an
  existing COM object, and verify its `Hwnd`, visibility, process ID, and
  workbooks before mutating anything.
- Treat `ReleaseComObject` as reference cleanup only. Call `Quit()` explicitly
  for an Excel process created by the caller, or use the cleanup scripts for a
  leaked process.
- When diagnosing an add-in lock, inspect both `Application.Workbooks` and
  `VBE.VBProjects`. A VBE project that is missing from `Workbooks` may be a
  ghost, but first confirm that the same COM object represents the Excel
  instance being inspected.

Related scripts:

- [demo-marshal-getactiveobject.ps1](../../scripts/ps/demo-marshal-getactiveobject.ps1)
- [demo-vb-getobject.ps1](../../scripts/ps/demo-vb-getobject.ps1)
- [Create-NewExcel.ps1](../../scripts/ps/Create-NewExcel.ps1)
- [List-ExcelInstances.ps1](../../scripts/ps/List-ExcelInstances.ps1)
- [Close-AllInvisibleExcelInstances.ps1](../../scripts/ps/Close-AllInvisibleExcelInstances.ps1)
- [Close-VisibleExcelInstances.ps1](../../scripts/ps/Close-VisibleExcelInstances.ps1)
- [Close-AllExcelInstances.ps1](../../scripts/ps/Close-AllExcelInstances.ps1)
