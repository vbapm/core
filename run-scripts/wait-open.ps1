[CmdletBinding()]
param(
	[Parameter(Mandatory)]
	[string]$Path,

	[int]$TimeoutSeconds = 30,

	[int]$PollMilliseconds = 250
)

$ErrorActionPreference = 'Stop'
$fullPath = [System.IO.Path]::GetFullPath($Path)

if (-not ('WaitOpenRotHelper' -as [type])) {
	Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class WaitOpenRotHelper {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ctx);

    public static List<object> GetRunningObjects() {
        var result = new List<object>();
        IRunningObjectTable rot = null;
        IBindCtx ctx = null;
        IEnumMoniker monikers = null;
        if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return result;
        if (CreateBindCtx(0, out ctx) != 0 || ctx == null) return result;

        try {
            rot.EnumRunning(out monikers);
            if (monikers == null) return result;
            monikers.Reset();
            var fetched = new IMoniker[1];
            while (monikers.Next(1, fetched, IntPtr.Zero) == 0) {
                object value = null;
                try {
                    if (rot.GetObject(fetched[0], out value) == 0 && value != null) {
                        result.Add(value);
                        value = null;
                    }
                } catch {}
                finally {
                    if (value != null) { try { Marshal.ReleaseComObject(value); } catch {} }
                    if (fetched[0] != null) { try { Marshal.ReleaseComObject(fetched[0]); } catch {} }
                }
            }
            return result;
        }
        finally {
            if (monikers != null) Marshal.ReleaseComObject(monikers);
            if (ctx != null) Marshal.ReleaseComObject(ctx);
            if (rot != null) Marshal.ReleaseComObject(rot);
        }
    }
}
"@
}

function Test-WorkbookOpen {
	param([string]$WorkbookPath)

	foreach ($candidate in @([WaitOpenRotHelper]::GetRunningObjects())) {
		$app = $null
		$ownsApp = $false
		try {
			try {
				if ([string]$candidate.Name -eq 'Microsoft Excel') {
					$app = $candidate
					$ownsApp = $true
				}
			} catch {}
			if ($null -eq $app) { $app = $candidate.Application }
			if ($null -eq $app -or [string]$app.Name -ne 'Microsoft Excel') { continue }

			foreach ($workbook in @($app.Workbooks)) {
				try {
					if ([string]::Equals([System.IO.Path]::GetFullPath([string]$workbook.FullName), $WorkbookPath, [System.StringComparison]::OrdinalIgnoreCase)) {
						return $true
					}
				} catch {}
			}
		} catch {}
		finally {
			if ($app -and -not $ownsApp) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {} }
			if (-not $ownsApp) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {} }
		}
	}

	return $false
}

$deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(0, $TimeoutSeconds))
while ([DateTime]::UtcNow -le $deadline) {
	if (Test-WorkbookOpen $fullPath) {
		Write-Output $fullPath
		exit 0
	}
	if ($PollMilliseconds -gt 0) { Start-Sleep -Milliseconds $PollMilliseconds }
}

Write-Error "Excel is not available to open '$fullPath'. Close any open Excel dialog and try again."
exit 1
