param(
    [Parameter(Position=0)]
    [string]$AppName,

    [Parameter(Position=1)]
    [string]$File,

    [switch]$Save
)

# -------
# Helpers
# -------

function GetFileBase {
    param([string]$Path)
    return [System.IO.Path]::GetFileName($Path)
}

# -------
# Main
# -------

$ErrorActionPreference = 'Stop'

if (-not $AppName -or -not $File) {
    Write-Error "ERROR #1: Invalid Input (appname and file are required)"
    exit 1
}

switch ($AppName) {
    "excel" {
        $fileBase = GetFileBase $File

        # Try to get a running Excel instance
        $excelApp = $null
        try {
            $excelApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        } catch {
            # Excel is not running — file is already closed
            exit 0
        }

        # Try to find the workbook by filename
        $workbook = $null
        try {
            $workbook = $excelApp.Workbooks($fileBase)
        } catch {
            # Workbook is not open — nothing to close
            exit 0
        }

        $workbook.Close($Save.IsPresent)
        exit 0
    }
    default {
        Write-Error "ERROR #3: Unsupported App `"$AppName`""
        exit 1
    }
}
