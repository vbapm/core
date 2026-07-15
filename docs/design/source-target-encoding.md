# Source/Target Encoding in vbaproject.toml — Design Plan

## Problem

Currently vbapm has no way for the user to declare the encoding of their VBA
source files. We auto-detect (via the encoding sniffer) and use the system
codepage as a fallback, but this fails when:

1. The source encoding differs from the system codepage (e.g., a CP1252 file
   opened on a Japanese CP932 machine).
2. The user wants to collaborate across machines with different locales.
3. The user needs to explicitly control the encoding for build reproducibility.

## Proposed TOML syntax

### Source encoding

```toml
[src]
Hello = { path = "src/Hello.bas", encoding = "cp1252" }
```

If all sources share the same encoding, a shorthand may be used under `[src-properties]`:

```toml
[project]
name = "multilingual-cp1252"
target = { type = "xlsm", path = "targets/xlsm" }

[src-properties]
encoding = "cp1252"

[src]
Hello = "src/Hello.bas"
```

Per-source `encoding` overrides the table-level default.

### Target encoding (optional, defaults to system codepage)

```toml
[project]
name = "my-project"
target = { type = "xlsm", path = "targets/xlsm", encoding = "cp932" }
```

When `target.encoding` differs from a source's encoding, the build pipeline
automatically transcodes the source to the target encoding before importing
into VBA.

## Validation: when encoding is required

On build, after loading all components:

1. Scan each component's code for non-ASCII characters (U+0080+).
2. If non-ASCII is found AND no encoding is declared in TOML → **build fails**
   with a helpful error message.

### Error message

```
Non-ASCII characters detected in "src/Hello.bas".
Please specify the encoding of the source code so the build process
can preserve them correctly.

Suggested change:

  [src-properties]
  encoding = "cp1252"

(Detection by jschardet, confidence: 95%)
```

### Auto-detection via jschardet

When no encoding is declared, we run [jschardet](https://www.npmjs.com/package/jschardet)
on the raw bytes of the source file to guess the encoding. We constrain
We constrain detection to the encodings vbapm supports:

```
CP874, CP932, CP936, CP949, CP950, CP1250–CP1258
```

UTF-8 and UTF-16 are NOT included — the encoding sniffer already detects
those (BOM + multi-byte heuristic) before we reach this fallback. jschardet
only runs when the sniffer returned `"unknown"`, meaning the buffer is
neither UTF-8 nor UTF-16.

This avoids jschardet suggesting irrelevant encodings (ISO-8859-*, EUC-*, etc.)
while keeping only the Windows codepages that the sniffer can't distinguish.

The suggestion is included in the error message so the user can copy-paste it.
If jschardet confidence is below 50%, we omit the suggestion and just tell
the user to specify an encoding.

**Implementation note:** jschardet's `detectAll` returns all candidate
encodings with confidence scores. We filter to our supported set and pick
the highest-confidence match:

```ts
const SUPPORTED = ["CP1252", "CP1251", "CP1250", "CP932", "CP936", "CP949", "CP950", "CP874", ...];
const results = jschardet.detectAll(buffer)
    .filter(r => SUPPORTED.includes(r.encoding))
    .sort((a, b) => b.confidence - a.confidence);
const best = results[0];
```

## Build pipeline changes

### Current flow

```
loadFromProject → stageBuildGraph (encode as system ACP) → VBA Import
```

### Proposed flow

```
loadFromProject
  ├── For each component: if encoding is NOT declared, check for non-ASCII.
  │     If non-ASCII found → fail with error + jschardet suggestion.
  │
  └── stageBuildGraph
        ├── For each component:
        │     sourceEncoding = declared encoding (or "unknown" if ASCII-only)
        │     targetEncoding = target.encoding ?? getSystemCodepage()
        │
        │     if sourceEncoding ≠ targetEncoding:
        │       transcode component.code from sourceEncoding → targetEncoding
        │
        └── Write encoded bytes to staging
```

### Transcoding

When source and target encodings differ, we convert the in-memory JS string
(already decoded from the source encoding) to the target encoding bytes via
iconv-lite:

```ts
const targetBytes = iconv.encode(component.code, targetLabel);
```

The component code is always a JS string (Unicode). Transcoding means
re-encoding it to the correct byte sequence for the target codepage.

## Source type changes

```ts
export interface Source {
    name: string;
    path: string;
    binary?: string;
    encoding?: string;   // e.g. "cp1252", "windows-1252"
}
```

## Implementation steps

1. **Add `jschardet` dependency** (`pnpm add jschardet`)
2. **Extend `Source` interface** with optional `encoding` field
3. **Extend `Target` interface** with optional `encoding` field
4. **Parse `encoding` from TOML** in `parseSource` and `parseTarget`
5. **Add `encoding` key under `[src-properties]`** as a default for all sources
6. **Add non-ASCII check + encoding validation** in `loadFromProject` or a
   new pre-build validation step
7. **Implement transcoding** in `stageBuildGraph` when source ≠ target encoding
8. **Tests:**
   - ASCII-only source without encoding → builds fine
   - Non-ASCII source without encoding → fails with jschardet suggestion
   - Non-ASCII source with encoding → builds fine
   - Source CP1252 → target CP932 → transcoded correctly
   - Per-source encoding overrides table-level default

## Edge cases

- **What if jschardet is wrong?** The suggestion is just that — a suggestion.
  The user must explicitly declare the encoding. We never auto-apply jschardet's
  guess.
- **What if the user declares the wrong encoding?** The build proceeds, but
  VBA may display garbled text. This is no worse than the current state.
- **What about existing projects without `encoding` in TOML?** ASCII-only
  files build fine without encoding. Non-ASCII files will start failing with
  a helpful message — this is a controlled breaking change that guides users
  to the fix.
- Let's say that someone create a new .bas file on their machine and decide to add it to the vbaproject. However, by default VS Code would create that file using utf-8. Can we have a test that creates this type of situation?
