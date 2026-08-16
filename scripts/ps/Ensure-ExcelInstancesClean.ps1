# Ensure-ExcelInstancesClean.ps1
#
# Guard run immediately before the e2e integration suite. It checks the Excel
# instance coordination registry and fails (exit 1) if any EXCEL.EXE process is
# running that is NOT tracked in the registry — i.e. a rogue instance that
# belongs to a user session or to another agent that has gone rogue or failed to
# clean up. Failing fast prevents the e2e suite (which drives Excel via COM)
# from hijacking or interfering with those sessions.
#
# The e2e suite itself creates *background* instances (VBA_BACKGROUND_BUILD=1)
# and registers each one, so in the common case there is nothing running when
# this guard fires.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Ensure-ExcelInstancesClean.ps1

$ErrorActionPreference = 'Stop'

$statusScript = Join-Path $PSScriptRoot 'Get-ExcelInstancesStatus.ps1'

# Run the status script as a child process so its exit code propagates via
# $LASTEXITCODE (dot-sourcing would instead terminate this script too).
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $statusScript -FailOnRogue

# Propagate the child's exit code.
exit $LASTEXITCODE
