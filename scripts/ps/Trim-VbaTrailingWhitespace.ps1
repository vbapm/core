[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Path = $PSScriptRoot,
    [switch]$Recurse = $true
)

$extensions = @('.bas', '.cls', '.frm', '.vba', '.doccls')

Get-ChildItem -LiteralPath $Path -File -Recurse:$Recurse |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
    ForEach-Object {
        $file = $_

        if (-not $PSCmdlet.ShouldProcess($file.FullName, 'Trim trailing whitespace')) {
            continue
        }

        # Read as raw bytes to preserve the original encoding (CP1252, etc.)
        [byte[]]$bytes = [System.IO.File]::ReadAllBytes($file.FullName)

        # Work on a copy — scan for trailing spaces/tabs (0x20/0x09) before CRLF (0x0D 0x0A)
        # and mark them for removal. We handle CRLF only (VBA files on Windows).
        $trimmed = new-object 'System.Collections.Generic.List[byte]'
        $i = 0
        while ($i -lt $bytes.Length) {
            # Scan backwards from current position to find start of trailing whitespace
            # We process character by character looking for CRLF to cut lines
            $lineStart = $i
            # Find the next CRLF
            while ($i -lt $bytes.Length -and -not ($bytes[$i] -eq 0x0D -and $i + 1 -lt $bytes.Length -and $bytes[$i + 1] -eq 0x0A)) {
                $i++
            }
            if ($i -ge $bytes.Length) {
                # No more CRLF — rest of file is the last line (no trailing newline)
                # Still trim trailing whitespace from it
                $lineEnd = $bytes.Length
                $trimEnd = $lineEnd
                while ($trimEnd -gt $lineStart -and ($bytes[$trimEnd - 1] -eq 0x20 -or $bytes[$trimEnd - 1] -eq 0x09)) {
                    $trimEnd--
                }
                for ($j = $lineStart; $j -lt $trimEnd; $j++) { $trimmed.Add($bytes[$j]) }
                break
            }
            # We found a CRLF at position $i
            $lineEnd = $i  # Position of CR
            # Trim trailing whitespace from this line (bytes from $lineStart to $lineEnd - 1)
            $trimEnd = $lineEnd
            while ($trimEnd -gt $lineStart -and ($bytes[$trimEnd - 1] -eq 0x20 -or $bytes[$trimEnd - 1] -eq 0x09)) {
                $trimEnd--
            }
            # Copy trimmed line content
            for ($j = $lineStart; $j -lt $trimEnd; $j++) { $trimmed.Add($bytes[$j]) }
            # Copy the CRLF
            $trimmed.Add(0x0D); $trimmed.Add(0x0A)
            # Move past CRLF
            $i += 2
        }

        [System.IO.File]::WriteAllBytes($file.FullName, $trimmed.ToArray())
    }
