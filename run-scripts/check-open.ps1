[CmdletBinding()]
param(
	[Parameter(Mandatory)]
	[string]$Path
)

$ErrorActionPreference = 'Stop'
$fullPath = [System.IO.Path]::GetFullPath($Path)

if (-not ('CheckOpenRotHelper' -as [type])) {
	Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class CheckOpenRotHelper {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ctx);

    public static List<object> GetRunningObjects() {
        var result = new List<object>();
        IRunningObjectTable rot = null;
        IBindCtx ctx = null;
        IEnumMoniker monikers = null;
        try {
            if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return result;
            if (CreateBindCtx(0, out ctx) != 0 || ctx == null) return result;

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

function Test-TargetPath {
	param([string]$CandidatePath)
	try {
		$candidateFullPath = [System.IO.Path]::GetFullPath($CandidatePath)
		return [string]::Equals(
			$candidateFullPath,
			$fullPath,
			[System.StringComparison]::OrdinalIgnoreCase
		)
	} catch {
		return $false
	}
}

function Test-VisibleExcelTarget {
	param([Parameter(Mandatory)][object]$Application)

	try {
		if (-not [bool]$Application.Visible) { return $false }
	} catch {
		return $false
	}

	# Item lookup can find a workbook loaded as an add-in even when collection
	# enumeration does not expose it.
	$workbook = $null
	try {
		$workbook = $Application.Workbooks.Item([System.IO.Path]::GetFileName($fullPath))
		if (Test-TargetPath ([string]$workbook.FullName)) { return $true }
	} catch {
		# Fall back to enumerating the workbook collection.
	} finally {
		if ($workbook) {
			try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null } catch {}
		}
	}

	try {
		foreach ($workbook in $Application.Workbooks) {
			try {
				if (Test-TargetPath ([string]$workbook.FullName)) { return $true }
			} catch {
				# Skip workbooks without a comparable file path.
			} finally {
				try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null } catch {}
			}
		}
	} catch {
		# The instance may have closed while being inspected.
	}

	try {
		foreach ($addin in $Application.AddIns) {
			try {
				if ([bool]$addin.Installed -and (Test-TargetPath ([string]$addin.FullName))) {
					return $true
				}
			} catch {
				# Skip add-ins without a comparable file path.
			} finally {
				try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin) | Out-Null } catch {}
			}
		}
	} catch {
		# The instance may have closed while being inspected.
	}

	return $false
}

$seenApplications = @{}
foreach ($candidate in @([CheckOpenRotHelper]::GetRunningObjects())) {
	$app = $null
	$appIsCandidate = $false
	try {
		try {
			if ([string]$candidate.Name -eq 'Microsoft Excel') {
				$app = $candidate
				$appIsCandidate = $true
			}
		} catch {}
		if ($null -eq $app) { $app = $candidate.Application }
		if ($null -eq $app -or [string]$app.Name -ne 'Microsoft Excel') { continue }

		try { $key = [string]$app.Hwnd } catch { $key = [string]$app.GetHashCode() }
		if ($seenApplications.ContainsKey($key)) { continue }
		$seenApplications[$key] = $true

		if (Test-VisibleExcelTarget $app) {
			Write-Output 'true'
			exit 0
		}
	} catch {
		# Skip inaccessible ROT entries and continue checking other instances.
	} finally {
		if ($app -and -not $appIsCandidate) {
			try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
		}
		if ($candidate) {
			try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {}
		}
	}
}

Write-Output 'false'
