# Characters Shared by the Multilingual Test Codepages

The scan found **5 non-ASCII characters** that round-trip through CP1252 and all 13 other codepages in this analysis.

Codepages checked, in order: CP1252, CP1251, CP1250, CP932, CP936, CP874, CP949, CP950, CP1253, CP1254, CP1255, CP1256, CP1257, CP1258.
The set includes the eight codepages with multilingual fixtures plus Windows-1253 through Windows-1258. The Excel build/export loop still runs only for codepages with fixtures.
ASCII characters are not listed; the report focuses on non-ASCII characters.

## Method

The [Node.js scanner](../../../../scripts/check-common-windows-codepage-characters.mjs) enumerates all 128 high-byte values in CP1252. It retains each decoded character only if encoding and decoding it returns the exact same character, then checks that set against every codepage above using the same round-trip test. This checks all **123** non-ASCII CP1252 characters that round-trip.

The [PowerShell verifier](../../../../scripts/ps/Test-CommonWindowsCodepageCharacters.ps1) independently checks the listed characters with strict .NET Windows codepage encoders and decoders, and prints their byte values.

## Common Non-ASCII Characters

| Character | Code point | CP1252 | CP1251 | CP1250 | CP932 | CP936 | CP874 | CP949 | CP950 | CP1253 | CP1254 | CP1255 | CP1256 | CP1257 | CP1258 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| ‘ | U+2018 | 91 | 91 | 91 | 81 65 | A1 AE | 91 | A1 AE | A1 A5 | 91 | 91 | 91 | 91 | 91 | 91 |
| ’ | U+2019 | 92 | 92 | 92 | 81 66 | A1 AF | 92 | A1 AF | A1 A6 | 92 | 92 | 92 | 92 | 92 | 92 |
| “ | U+201C | 93 | 93 | 93 | 81 67 | A1 B0 | 93 | A1 B0 | A1 A7 | 93 | 93 | 93 | 93 | 93 | 93 |
| ” | U+201D | 94 | 94 | 94 | 81 68 | A1 B1 | 94 | A1 B1 | A1 A8 | 94 | 94 | 94 | 94 | 94 | 94 |
| … | U+2026 | 85 | 85 | 85 | 81 63 | A1 AD | 85 | A1 A6 | A1 4B | 85 | 85 | 85 | 85 | 85 | 85 |

## Em Dash

U+2014 (em dash) is not common to all 14 codepages. It cannot be represented in CP932 or CP949.

| Character | Code point | CP1252 | CP1251 | CP1250 | CP932 | CP936 | CP874 | CP949 | CP950 | CP1253 | CP1254 | CP1255 | CP1256 | CP1257 | CP1258 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| — | U+2014 | 97 | 97 | 97 | not representable | A1 AA | 97 | not representable | A1 58 | 97 | 97 | 97 | 97 | 97 | 97 |

## Bullet

U+2022 (bullet) is not common to all 14 codepages. It cannot be represented in CP932, CP936, CP949, or CP950.

| Character | Code point | CP1252 | CP1251 | CP1250 | CP932 | CP936 | CP874 | CP949 | CP950 | CP1253 | CP1254 | CP1255 | CP1256 | CP1257 | CP1258 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| • | U+2022 | 95 | 95 | 95 | not representable | not representable | 95 | not representable | not representable | 95 | 95 | 95 | 95 | 95 | 95 |

> This report establishes codepage round-trip support only. It does not show that VBA accepts these punctuation characters in identifiers.

## Reproduce

From the repository root, run:

```powershell
node scripts/check-common-windows-codepage-characters.mjs --markdown docs/analysis/codepages/commonality/windows-codepage-character-report.md
.\scripts\ps\Test-CommonWindowsCodepageCharacters.ps1
```

The [multilingual GitHub Actions workflow](../../../../.github/workflows/multilingual-encoding-test.yml) continues to run the Excel build/export fixtures; these wider codepage scans are manual checks.
