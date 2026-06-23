<#
.SYNOPSIS
    Multilingual encoding test loop for CI.
    
.DESCRIPTION
    Iterates over Windows ANSI codepages, sets each one in the registry,
    and runs the multilingual Jest suite (mlang.config.mjs) for that
    codepage. Excel must be installed and closed between iterations so
    that new Excel processes pick up the changed ACP.
    
.NOTES
    Requires administrator privileges to modify HKLM registry.
    Restores the original ACP after all tests complete.
#>
param(
    [string[]]$Codepages = @("1252", "1251", "1250", "932", "936", "950")
)

$ErrorActionPreference = "Stop"

# Save original ACP
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
$originalAcp = (Get-ItemProperty $regPath).ACP
Write-Output "Original ACP: $originalAcp"

$failed = @()
$passed = @()

foreach ($cp in $Codepages) {
    Write-Output "`n=============================================="
    Write-Output "Testing codepage: $cp"
    Write-Output "=============================================="

    # Ensure Excel is fully closed before changing codepage
    Write-Output "Closing Excel..."
    taskkill /f /im excel.exe 2>$null
    Start-Sleep -Seconds 3

    # Set the ANSI codepage in registry
    Write-Output "Setting ACP to $cp..."
    Set-ItemProperty -Path $regPath -Name "ACP" -Value $cp -Type String

    # Verify
    $current = (Get-ItemProperty $regPath).ACP
    Write-Output "ACP now: $current"

    # Run the multilingual test suite
    $env:CI = "1"
    npx jest --config mlang.config.mjs --runInBand --no-coverage

    if ($LASTEXITCODE -eq 0) {
        Write-Output "✓ Codepage $cp PASSED"
        $passed += $cp
    } else {
        Write-Output "✗ Codepage $cp FAILED"
        $failed += $cp
    }
}

# Restore original ACP
Write-Output "`nRestoring original ACP: $originalAcp"
Set-ItemProperty -Path $regPath -Name "ACP" -Value $originalAcp -Type String

# Summary
Write-Output "`n=============================================="
Write-Output "RESULTS"
Write-Output "=============================================="
Write-Output "Passed: $($passed -join ', ')"
if ($failed.Count -gt 0) {
    Write-Output "Failed: $($failed -join ', ')"
    throw "Multilingual tests failed for codepages: $($failed -join ', ')"
}
Write-Output "All codepages passed!"
