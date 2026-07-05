# Encoding Sniffer — Design Document

## Overview

The encoding sniffer (`src/build/encoding-sniffer.ts`) tries to determines the encoding of
a buffer containing VBA source code. VBA's `Component.Export` writes files in the
system ANSI codepage, and `Component.Import` reads files in the same codepage, however some people might want to edit their VBA code in UTF-8 or other encoding that support the majority of unicode characters.

The sniffer's job is to identify the encoding so we can correctly decode the bytes
into a JavaScript string.

---

## Detection logic

The sniffer reports an encoding when it has high confidence. When it cannot
distinguish between encodings reliably, it returns `unknown` and the caller
must provide the encoding explicitly (e.g. from the system codepage).

### Detected encodings

| Signal | Encoding | Confidence |
|--------|----------|------------|
| `FF FE` at start | UTF-16 LE with BOM | Certain |
| `FE FF` at start | UTF-16 BE with BOM | Certain |
| `EF BB BF` at start | UTF-8 with BOM | Certain |
| Valid UTF-8 with multi-byte (non-ASCII) sequences | UTF-8 without BOM | High |
| None of the above | **unknown** | Need external hint |

### Rationale for UTF-8 without BOM

A buffer that decodes as valid UTF-8 *and* contains at least one multi-byte
sequence is almost certainly UTF-8. It is vanishingly unlikely that a
single-byte ANSI file (CP1252, CP1251, etc.) would happen to form valid
UTF-8 multi-byte sequences for its non-ASCII content.

A pure-ASCII or empty buffer returns `unknown` — it could be any encoding
(all single-byte codepages and UTF-8 share the ASCII range).

### Behavior when encoding is unknown

When the sniffer returns `unknown`:

1. **`decodeBuffer`:** Falls through to the system codepage path — uses
   `getSystemCodepage()` to determine the actual encoding. This is correct
   for files exported by VBA on the current machine.

2. **`Component` constructor:** When `Codepage.Unknown` is passed, the
   sniffer returns `unknown`, and `decodeBuffer` resolves it to the system
   codepage. The `encoding` property on the Component reflects `unknown` —
   meaning "we don't know the original encoding, we guessed the system
   codepage."

3. **Callers who know the encoding:** `loadFromExport` calls
   `getSystemCodepage()` directly and passes it to the Component constructor,
   skipping sniffing entirely. This is the preferred path.

---

## Future improvements

When we need to handle files from *other* machines (e.g., a French developer
sharing a project with a Japanese developer), we need to detect encodings
without a BOM or UTF-8 multi-byte signal. The following heuristics could be
added:

### Statistical byte frequency analysis

Different encodings have different byte frequency profiles:
- CP1252 (Western European): bytes 0x80-0x9F are rare (special chars),
  0xC0-0xFF are common (accented chars)
- CP1251 (Cyrillic): bytes 0xC0-0xFF are very common (Cyrillic letters)
- CP932 (Japanese): bytes 0x81-0x9F and 0xE0-0xFC are lead bytes for
  double-byte sequences; single bytes in those ranges are unusual
- CP936/GBK (Chinese): similar lead-byte patterns to CP932 but different ranges

We could score each candidate encoding based on byte frequency and pick the
most likely one.

### Language detection on decoded output

Try decoding the buffer with each candidate encoding, then run language
detection on the result. If decoding as CP1251 produces recognizable Russian
words, while decoding as CP1252 produces garbage, we can confidently choose
CP1251.

**Challenge:** Language detection libraries are heavy dependencies. A
lightweight approach could check for the presence of language-specific Unicode
ranges (Cyrillic U+0400–U+04FF, CJK U+4E00–U+9FFF, etc.).

### Multi-byte sequence pattern matching

CP932, CP936, CP949, and CP950 are multi-byte encodings with specific lead-byte
and trail-byte ranges. We can detect them by:
1. Scanning for valid lead-byte / trail-byte pairs
2. Checking that no isolated trail bytes appear (which would indicate a
   single-byte encoding)
3. Verifying that the decoded result contains CJK characters

### System codepage as a Bayesian prior

If we're running on a Japanese Windows machine, CP932 is the most likely
encoding for any non-UTF buffer. The system codepage should be used as a
strong prior in any statistical model.

### Confidence levels

| Signal | Encoding | Confidence |
|--------|----------|------------|
| BOM detected | (any) | Certain |
| Multi-byte CP932 pattern + CJK output | CP932 | High |
| Multi-byte CP936 pattern + CJK output | CP936 | High |
| Valid UTF-8 with multi-byte sequences | UTF-8 | Medium |
| Matches system codepage byte profile | System CP | Low-Medium |
| Pure ASCII (any encoding) | System CP | Low |

---

## Machine-specific profiles (future)

Each machine could maintain a profile of which encodings it has encountered,
encountered, allowing the sniffer to learn and improve over time. This is
useful for teams where a mix of locales is common (e.g., a Japanese developer
regularly working with French-accented VBA code).
