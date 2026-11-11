[CmdletBinding()]
param(
    [int[]]$Codepages = @(1252, 1251, 1250, 932, 936, 874, 949, 950, 1253, 1254, 1255, 1256, 1257, 1258)
)

$ErrorActionPreference = 'Stop'

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'This verification script requires Windows.'
}

$codepoints = @(0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x2014, 0x2022)
$results = @{}
$rows = @()

foreach ($codepage in $Codepages) {
    $encoding = [System.Text.Encoding]::GetEncoding(
        $codepage,
        [System.Text.EncoderExceptionFallback]::new(),
        [System.Text.DecoderExceptionFallback]::new()
    )
    $row = [ordered]@{ Codepage = $codepage }
    $results[$codepage] = @{}

    foreach ($codepoint in $codepoints) {
        $label = 'U+{0:X4}' -f $codepoint
        $character = [char]::ConvertFromUtf32($codepoint)
        try {
            $bytes = $encoding.GetBytes($character)
            $roundTrip = $encoding.GetString($bytes)
            if ($roundTrip -cne $character) {
                throw "CP$codepage did not round-trip $label."
            }

            $value = [BitConverter]::ToString($bytes)
        } catch [System.Text.EncoderFallbackException] {
            $value = 'not representable'
        }

        $row[$label] = $value
        $results[$codepage][$label] = $value
    }

    $rows += [pscustomobject]$row
}

Write-Output 'Strict .NET Windows codepage round-trip matrix (bytes shown; unavailable characters are explicit):'
$rows | Format-Table -AutoSize | Out-String -Width 240 | Write-Output

$expectedCommon = @('U+2018', 'U+2019', 'U+201C', 'U+201D', 'U+2026')
$observedCommon = @(
    $expectedCommon | Where-Object {
        $label = $_
        @($Codepages | Where-Object { $results[$_][$label] -eq 'not representable' }).Count -eq 0
    }
)
if (($observedCommon -join ',') -ne ($expectedCommon -join ',')) {
    throw "Unexpected common-character results: $($observedCommon -join ', ')"
}

$emDashUnsupported = @(
    $Codepages | Where-Object { $results[$_]['U+2014'] -eq 'not representable' }
)
Write-Output "Common non-ASCII characters verified: $($observedCommon -join ', ')"
Write-Output "Em dash (U+2014) unavailable in: $($emDashUnsupported -join ', ')"
if (($emDashUnsupported -join ',') -ne '932,949') {
    throw "Unexpected em-dash availability: unsupported codepages were $($emDashUnsupported -join ', ')"
}

$bulletUnsupported = @(
    $Codepages | Where-Object { $results[$_]['U+2022'] -eq 'not representable' }
)
Write-Output "Bullet (U+2022) unavailable in: $($bulletUnsupported -join ', ')"
if (($bulletUnsupported -join ',') -ne '932,936,949,950') {
    throw "Unexpected bullet availability: unsupported codepages were $($bulletUnsupported -join ', ')"
}
