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

function Add-ReportMatch {
	param(
		[Parameter(Mandatory)][string]$CandidatePath,
		[Parameter(Mandatory)][string]$Source,
		[Parameter(Mandatory)][int]$ProcessId,
		[Parameter(Mandatory)][long]$WindowHandle,
		[Parameter(Mandatory)][bool]$Visible,
		[AllowNull()][Nullable[bool]]$ReadOnly
	)

	try {
		$candidateFullPath = [System.IO.Path]::GetFullPath($CandidatePath)
	} catch {
		return
	}
	if ([System.IO.Path]::GetFileName($candidateFullPath) -ine $targetName) { return }

	$normalizedPath = $candidateFullPath -replace '/', '\'
	$instanceKey = if ($ProcessId -gt 0) { "pid:$ProcessId" } else { "hwnd:$WindowHandle" }
	$key = "$instanceKey|$($normalizedPath.ToLowerInvariant())"
	if ($reportMatches.ContainsKey($key)) {
		$existing = $reportMatches[$key]
		if ($existing.source -notlike "*$Source*") {
			$existing.source = "$($existing.source), $Source"
		}
		if ($null -ne $ReadOnly -and $null -eq $existing.read_only) {
			$existing.read_only = [bool]$ReadOnly
			$existing.editable = -not [bool]$ReadOnly
			$existing.write_lock_held = -not [bool]$ReadOnly
		}
		return
	}

	$record = [ordered]@{
		path         = $candidateFullPath
		same_path    = [string]::Equals($candidateFullPath, $fullPath, [System.StringComparison]::OrdinalIgnoreCase)
		source       = $Source
		instance_name = 'Microsoft Excel'
		visible      = $Visible
	}
	if ($ProcessId -gt 0) { $record.process_id = $ProcessId }
	if ($WindowHandle -gt 0) { $record.window_handle = $WindowHandle }
	if ($null -ne $ReadOnly) {
		$record.read_only = [bool]$ReadOnly
		$record.editable = -not [bool]$ReadOnly
		$record.write_lock_held = -not [bool]$ReadOnly
	}

	$match = [pscustomobject]$record
	$reportMatches[$key] = $match
}

$reportMatches = @{}
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
		if ($null -eq $app) {
			try { $app = $candidate.Application } catch { $app = $null }
		}
		if ($null -eq $app -or [string]$app.Name -ne 'Microsoft Excel') { continue }

		$windowHandle = 0L
		try { $windowHandle = [long]$app.Hwnd } catch {}
		$appKey = if ($windowHandle -gt 0) { "hwnd:$windowHandle" } else { "app:$($app.GetHashCode())" }
		if ($seenApplications.ContainsKey($appKey)) { continue }
		$seenApplications[$appKey] = $true

		$processId = [CheckOpenRotHelper]::GetProcessId($windowHandle)
		$visible = [bool]$app.Visible

		$workbook = $null
		try {
			$workbook = $app.Workbooks.Item($targetName)
			$readOnly = $null
			try { $readOnly = [bool]$workbook.ReadOnly } catch {}
			Add-ReportMatch -CandidatePath ([string]$workbook.FullName) -Source 'workbook' -ProcessId $processId -WindowHandle $windowHandle -Visible $visible -ReadOnly $readOnly
		} catch {
			# The workbook may not be open in this instance.
		} finally {
			if ($workbook) {
				try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null } catch {}
			}
		}

		try {
			foreach ($workbook in $app.Workbooks) {
				try {
					$readOnly = $null
					try { $readOnly = [bool]$workbook.ReadOnly } catch {}
					Add-ReportMatch -CandidatePath ([string]$workbook.FullName) -Source 'workbook' -ProcessId $processId -WindowHandle $windowHandle -Visible $visible -ReadOnly $readOnly
				} catch {
					# Skip workbook entries that cannot be inspected.
				} finally {
					try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null } catch {}
				}
			}
		} catch {
			# Continue with other ROT entries if the workbook collection is unavailable.
		}

		try {
			foreach ($addin in $app.AddIns) {
				try {
					if ([bool]$addin.Installed) {
						$addinPath = [string]$addin.FullName
						$addinReadOnly = $null
						$addinWorkbook = $null
						try {
							$addinWorkbook = $app.Workbooks.Item([System.IO.Path]::GetFileName($addinPath))
							if (-not [string]::Equals([System.IO.Path]::GetFullPath([string]$addinWorkbook.FullName), [System.IO.Path]::GetFullPath($addinPath), [System.StringComparison]::OrdinalIgnoreCase)) {
								[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($addinWorkbook) | Out-Null
								$addinWorkbook = $null
							}
							if ($addinWorkbook) { $addinReadOnly = [bool]$addinWorkbook.ReadOnly }
						} catch {
							# Add-in may not be available as a workbook object.
						} finally {
							if ($addinWorkbook) {
								try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($addinWorkbook) | Out-Null } catch {}
							}
						}
						Add-ReportMatch -CandidatePath $addinPath -Source 'addin' -ProcessId $processId -WindowHandle $windowHandle -Visible $visible -ReadOnly $addinReadOnly
					}
				} catch {
					# Skip add-in entries that cannot be inspected.
				} finally {
					try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin) | Out-Null } catch {}
				}
			}
		} catch {
			# Some Excel instances do not expose the add-ins collection.
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

$report = [pscustomobject]@{
	instances = @($reportMatches.Values | Sort-Object process_id, path)
}
ConvertTo-Json -InputObject $report -Depth 6 -Compress
