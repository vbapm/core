[CmdletBinding()]
param(
	[Parameter(Mandatory)]
	[string]$Path
)

$ErrorActionPreference = 'Stop'
$fullPath = [System.IO.Path]::GetFullPath($Path)

try {
	$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
	Write-Output 'false'
	exit 0
}

if (-not $app.Visible) {
	Write-Output 'false'
	exit 0
}

foreach ($workbook in $app.Workbooks) {
	try {
		$workbookPath = [System.IO.Path]::GetFullPath([string]$workbook.FullName)
		if ([string]::Equals($workbookPath, $fullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
			Write-Output 'true'
			exit 0
		}
	} catch {
		# Skip workbooks without a comparable file path.
	}
}

Write-Output 'false'
