# Diagnose-VBIDETypeLib.ps1
# Checks whether the VBIDE type library (VBE6EXT.OLB / VBE7.DLL) is
# registered on the system and reports its GUID, version, and file paths.
#
# Usage: .\scripts\ps\Diagnose-VBIDETypeLib.ps1

$ErrorActionPreference = "Continue"
$guid = "{0002E157-0000-0000-C000-000000000046}"

Write-Output "=== VBIDE Type Library Diagnostic ==="
Write-Output "GUID: $guid"

# ---------- Registry: TypeLib ----------
Write-Output "`n--- Registry: HKCR\TypeLib\$guid ---"
$tlibPath = "Registry::HKEY_CLASSES_ROOT\TypeLib\$guid"
if (Test-Path $tlibPath) {
    Write-Output "FOUND"
    Get-ChildItem $tlibPath -Recurse | ForEach-Object {
        $val = Get-ItemProperty $_.PSPath
        Write-Output "  $($_.PSChildName): $($val.'(default)')"
    }
} else {
    Write-Output "NOT FOUND"
}

# ---------- Registry: Interface ----------
Write-Output "`n--- Registry: HKCR\Interface\$guid ---"
$ifacePath = "Registry::HKEY_CLASSES_ROOT\Interface\$guid"
if (Test-Path $ifacePath) {
    Write-Output "FOUND"
    Get-ChildItem $ifacePath | ForEach-Object {
        $val = Get-ItemProperty $_.PSPath
        Write-Output "  $($_.PSChildName): $($val.'(default)')"
    }
} else {
    Write-Output "NOT FOUND"
}

# ---------- Known file locations ----------
Write-Output "`n--- File search ---"
$searchPaths = @(
    # VBE6 (32-bit, older Office)
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB",
    "${env:ProgramFiles}\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB",
    # VBE7 (32-bit, modern Office)
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL",
    "${env:ProgramFiles}\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL",
    # VBE7 via Office Click-to-Run VFS
    "${env:ProgramFiles}\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1\VBE7.DLL",
    "${env:ProgramFiles}\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL",
    # VBE6 via Office Click-to-Run VFS
    "${env:ProgramFiles}\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
)

$foundAny = $false
foreach ($p in $searchPaths) {
    if (Test-Path $p) {
        $item = Get-Item $p
        Write-Output "FOUND: $p"
        Write-Output "  Version: $($item.VersionInfo.FileVersion)"
        Write-Output "  ProductVersion: $($item.VersionInfo.ProductVersion)"
        $foundAny = $true
    }
}

if (-not $foundAny) {
    Write-Output "No known VBIDE file found at standard locations."

    # Broader search under Common Files
    Write-Output "`n--- Broad search under Common Files ---"
    $commonBase = @(
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VBA",
        "${env:ProgramFiles}\Common Files\Microsoft Shared\VBA"
    )
    foreach ($base in $commonBase) {
        if (Test-Path $base) {
            Get-ChildItem $base -Recurse -Include "VBE*.DLL","VBE*.OLB" -ErrorAction SilentlyContinue |
                Select-Object -First 10 |
                ForEach-Object {
                    Write-Output "FOUND: $($_.FullName)"
                    Write-Output "  Version: $($_.VersionInfo.FileVersion)"
                    $foundAny = $true
                }
        }
    }
}

# ---------- Summary ----------
Write-Output "`n=== Summary ==="
if ($foundAny) {
    Write-Output "VBIDE type library is AVAILABLE."
} else {
    Write-Output "VBIDE type library was NOT FOUND on this system."
}
