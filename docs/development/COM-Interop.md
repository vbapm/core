# Excel COM lookup on Windows

## How to get a hold of the only Excel instance?

In case there is only one Excel instance, it's very easy, just use 
```powershell
[System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
```
`Marshal.GetActiveObject` returned the visible Excel instance containing the
user's workbooks (see demo-marshal-getactiveobject.ps1)

Avoid using the following:
```powershell
[Microsoft.VisualBasic.Interaction]::GetObject($null, "Excel.Application")
```
It would create a new invisible instance of Excel (see demo-vb-getobject.ps1).


## What happens when there's more than one Excel instance?

A process ID. When more than one Excel process exists,
the returned object depends on how Excel registered itself in the COM Running
Object Table and on which COM lookup path the caller uses.




## Why this matters

Both methods ask COM for an Excel object. Neither method enumerates all Excel
processes or accepts 

In this environment, the VB.NET `GetObject` call can activate a new hidden
Excel process when it cannot bind to the visible instance. Releasing the
returned COM reference does not call `Excel.Quit()`, so the hidden process can
remain alive until the cleanup script terminates it.

The `Hwnd` value is a useful identity check. If two lookup calls return
different `Hwnd` values, they refer to different Excel processes. Map the
window handle to a process ID when diagnosing the result.

## Warning

Do not use this call as a general way to attach to Excel:

```powershell
[Microsoft.VisualBasic.Interaction]::GetObject($null, "Excel.Application")
```

On this machine, repeated calls returned new hidden Excel processes instead
of the visible instance. If the goal is to attach to an existing COM object,
use `Marshal.GetActiveObject` and verify the returned `Hwnd`, visibility,
process ID, and workbooks.

`Marshal.GetActiveObject` does not create Excel. If the goal is to create a
new instance intentionally, use an explicit creation call such as:

```powershell
$excel = New-Object -ComObject Excel.Application
```

The caller owns that instance and must call `$excel.Quit()` when finished.

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
- Do not use either lookup method as a way to choose among multiple Excel
	processes. Use a known process ID or window handle when the target matters.
- Use `Marshal.GetActiveObject` from Windows PowerShell to retrieve an existing
	COM object, and verify the returned `Hwnd`, visibility, process ID, and
	workbooks before mutating anything.
- Treat `ReleaseComObject` as reference cleanup only. Call `Quit()` explicitly
	for an Excel process created by the caller, or use the invisible-instance
	cleanup script for a leaked process.
- When diagnosing an add-in lock, inspect both `Application.Workbooks` and
	`VBE.VBProjects`. A VBE project that is missing from `Workbooks` may be a
	ghost, but first confirm that the same COM object represents the Excel
	instance being inspected.

Related scripts:

- [demo-vb-getobject.ps1](../../scripts/ps/demo-vb-getobject.ps1)
- [Close-AllInvisibleExcelInstances.ps1](../../scripts/ps/Close-AllInvisibleExcelInstances.ps1)
- [Close-Addin.ps1](../../scripts/ps/Close-Addin.ps1)
