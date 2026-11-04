[CmdletBinding()]
param(
	[Parameter(Mandatory)][string]$AddinPath
)

$ErrorActionPreference = 'Stop'

try {
	$Excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
	exit 0
}

try {
	$NormalizedAddinPath = [System.IO.Path]::GetFullPath($AddinPath)
	$AddinName = [System.IO.Path]::GetFileName($AddinPath)

	try {
		$Workbook = $Excel.Workbooks.Item($AddinName)
		if ($Workbook -and [System.IO.Path]::GetFullPath($Workbook.FullName) -eq $NormalizedAddinPath) {
			Write-Output "Closing open $AddinName..."
			$Workbook.Close($false)
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
		}
	} catch {
	}

	foreach ($Workbook in @($Excel.Workbooks)) {
		try {
			if ($Workbook.FullName -and ([System.IO.Path]::GetFullPath($Workbook.FullName) -eq $NormalizedAddinPath)) {
				Write-Output "Closing open $AddinName..."
				$Workbook.Close($false)
			}
		} catch {
		}
	}

	foreach ($Addin in @($Excel.AddIns)) {
		try {
			if ($Addin.FullName -and ([System.IO.Path]::GetFullPath($Addin.FullName) -eq $NormalizedAddinPath)) {
				$Addin.Installed = $false
			}
		} catch {
		}
	}
} finally {
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
}