# Source Files — Listing, Grouping & Sorting

> **Status:** Design (replaces the narrow "Sorted Source Files" plan).
> The previous implementation of `byComponentTypeThenName` is already in the code;
> this document expands on it with a configurable `[src-properties]` table.

---

## 1. Motivation

VBA projects are more like a cohesive collection of scripts than a formally
compiled assembly. The philosophy of listing every file individually in
`vbaproject.toml` — inherited from C#-style `.csproj` manifests — is overly
verbose for VBA. This design introduces **two complementary listing modes**
controlled by a new `[src-properties]` table:

| Mode | Key | Description |
|---|---|---|
| **Grouped** | 3 reserved keys | Use wildcard patterns under 3 reserved keys (`Modules`, `Forms`, `Classes`). Compact and low-maintenance. **The default for new projects.** |
| **Individual** | one key per file | List every source file by name. The order is governed by the detected convention, optionally enforced via `[src-properties]`. |

In *both* modes the manifest remains easy to scan and the tooling knows how to
add, remove, and reorder entries predictably.

---

## 2. How the Tool Knows What Convention Is in Use

The tool does **not** need a config table to tell it what convention is already
in the file. It infers the convention automatically from the `[src]` entries
themselves, using `detectSrcStructure()` (§6).

| What the tool sees in `[src]` | Detected convention |
|---|---|
| Exactly 3 keys named `Modules`, `Forms`, `Classes` with glob, array, or literal path values | **Grouped** |
| Individual entries, all `.bas` contiguous, all `.frm` contiguous, all `.cls` contiguous | **Sorted by types** |
| Individual entries, globally alphabetical | **Sorted alphabetically** |
| Individual entries, contiguous by type AND alphabetical within each type | **Sorted by type then alphabetically** |
| None of the above | **Unstructured** |

Once detected, the tool respects this convention on subsequent operations
(add, export) — it doesn't reorder or reformat unless explicitly told to.

---

## 3. The `[src-properties]` Table (Optional Enforcement)

The `[src-properties]` section is **optional**. Its purpose is not to describe
what convention is already in place (the tool detects that automatically), but
to tell the tool to **actively enforce** a specific ordering — reordering
entries on add/export, validating structure, and so on.

If `[src-properties]` is absent, the tool is hands-off: it detects the
convention and works with it, but never reorders.

If `[src-properties]` is present, the tool enforces what the keys say.

### 3.1 Key reference

| Key | Type | Description |
|---|---|---|
| `grouping` | `boolean` | Enforce grouped mode: `[src]` must have exactly the 3 reserved keys. If a user accidentally adds an individual entry, the tool flags it. |
| `sort.by-types` | `boolean` | (Individual mode) Enforce type grouping: Objects → Modules → Forms → Classes. On add/export, insert new entries into the correct type section. |
| `sort.alphabetical` | `boolean` | (Individual mode) Enforce alphabetical order. When combined with `sort.by-types`, alphabetical is *within* each type group; alone, it's global. |

All keys default to `false` when the `[src-properties]` table is present but
a key is omitted — i.e., you opt in to each enforcement individually.

### 3.2 Example: enforcement in individual mode

```toml
[src-properties]
sort.by-types = true
sort.alphabetical = true
```

With this in place, `vba add MyModule` will insert the new entry in the
correct alphabetical position within the modules section, rather than
appending at the end.

### 3.3 Example: no enforcement (the common case)

```toml
# No [src-properties] section at all.
# The tool detects the convention and works with it as-is.
# No automatic reordering happens.

[src]
Modules = "src/**/*.bas"
Forms = "src/**/*.frm"
Classes = "src/**/*.cls"
```

---

## 4. Grouped Convention

### 4.1 Manifest shape

The grouped convention uses exactly three keys in `[src]` — no
`[src-properties]` needed. The tool detects this convention automatically.

```toml
[src]
Modules = "src/**/*.bas"
Forms = "src/**/*.frm"
Classes = ["src/**/*.cls", "src/Documents/*.cls"]
```

Note that Classes includes regular Class Modules and Office specific objects.

Each value is either a **single glob string** or an **array of glob strings**
(literal file paths are also valid — e.g. `Modules = "src/MyModule.bas"` for
a single module). The globs follow
[`minimatch`](https://github.com/isaacs/minimatch) syntax (already used in the
project for `.gitignore`-style patterns).

### 4.2 Resolution at build time

At build time the tool:

1. Reads the glob(s) for each key.
2. Resolves them against the project directory (relative to `vbaproject.toml`).
3. Collects all matching `.bas` / `.frm` / `.cls` files.
4. Assembles them in the conventional order: Objects → Modules → Forms → Classes.

Files within each group are used in the order returned by glob resolution
(typically filesystem order). **No sorting is applied unless `sort.*` options
are explicitly set.** Even in grouped mode, the `sort.by-types` and
`sort.alphabetical` keys from `[src-properties]` are respected for ordering
within each type group — but the tool never imposes a sort that the user did
not request.

### 4.3 Adding a new file

When the user runs `vbapm add`, the tool:

- Determines the component type from the extension.
- Does **not** modify `[src]` — the glob already covers the new file.
- Simply creates the file on disk under the appropriate subdirectory.

### 4.4 Why the grouped convention fits VBA

| Aspect | C# `.csproj` | VBA `vbaproject.toml` |
|---|---|---|
| Compilation unit | Strict assembly with explicit include/exclude | Loose collection of scripts |
| File count | Often dozens/hundreds — explicit listing is important for build ordering | Usually < 30 files — order rarely matters |
| Refactoring | Renaming/ moving files requires updating the project file | Glob patterns absorb renames automatically |
| Philosophy | Explicit is better than implicit | Convention over configuration (à la JavaScript/Node.js) |

---

### 4.5 No reserved keywords for document objects

The `Classes` key covers **all** `.cls` files — user classes, `ThisWorkbook`,
`Sheet1`, `Sheet2`, etc. There is intentionally no separate reserved key for
Excel document objects (e.g. `ThisWorkbook`). Rationale:

- **Language neutrality.** A German Excel add-in uses `DieseArbeitsmappe`
  instead of `ThisWorkbook`. A French one uses `Feuil1` instead of `Sheet1`.
  A reserved keyword would force non-English users to rename their components
  or use an English alias in the manifest — unnecessary friction.
- **No type distinction on disk.** VBA exports both `vbext_ct_Document` and
  `vbext_ct_ClassModule` as plain `.cls` files. The TypeScript tooling cannot
  tell them apart by extension or content alone, so a dedicated key would
  require unreliable heuristics.
- **Separate folders work better.** Users who want visual separation between
  user classes and Excel objects can use subdirectories and multiple glob
  entries:

  ```toml
  Classes = [
      "src/Class Modules/*.cls",
      "src/Excel Objects/*.cls"
  ]
  ```

  This gives the user control over grouping without the tool imposing naming
  conventions.

The same principle applies to other Office hosts (Word's `ThisDocument`,
Access forms, etc.) — all `.cls` variants live under `Classes`.

### 4.6 `subfolders` config

The `subfolders` key in `[src-properties]` controls where new files are
placed on disk (used by `vba add` and `vba export`). It supports four keys:

```toml
[src-properties]
subfolders = { Modules = "Modules", Forms = "Forms", Classes = "Class Modules", Objects = "Excel Objects" }
```

| Key | Component type | Default (no config) |
|---|---|---|
| `Modules` | `"module"` | `src/` |
| `Forms` | `"form"` | `src/` |
| `Classes` | `"class"` | `src/` |
| `Objects` | `"document"` | `src/` (falls back to `Classes` dir if set, else `src/`) |

The `Objects` key lets users separate Excel document objects
(`ThisWorkbook`, `Sheet1`, etc.) from regular user classes without reserved
keywords — the separation is by subdirectory, not by naming convention.

> **Prerequisite:** The `"document"` `ComponentType` is currently defined but
> never assigned (all `.cls` files are `"class"`). Before `Objects` can take
> effect, the component loader must detect document modules — either from
> VBA metadata during export or from `.cls` file content patterns.

---

## 5. Individual Listing Convention

Every source file has its own key in `[src]`. No `[src-properties]` is needed —
the tool detects the sorting pattern automatically. Document objects
(`ThisWorkbook`, `Sheet1`, etc.) are listed alongside user classes with
no special treatment.

```toml
[src]
Validation = "src/Validation.bas"
JsonConverter = "src/JsonConverter.bas"

UserForm1 = "src/UserForm1.frm"

MyClass = "src/MyClass.cls"
Sheet1 = "src/Sheet1.cls"
ThisWorkbook = "src/ThisWorkbook.cls"
```

### 5.1 Detected sorting patterns

The tool detects which sorting convention (if any) is in use via
`detectSrcStructure()` (§6):

| Detected pattern | `sortedByTypes` | `sortedAlphabetically` | Description |
|---|---|---|---|
| By type, then name | `true` | `true` | Document objects contiguous, then all `.bas` contiguous, all `.frm` contiguous, all `.cls` contiguous, and alphabetical within each group. |
| By type only | `true` | `false` | Files grouped by extension but in **insertion order** within each group — the order files were added to the project. When creating a new config with `--individual`, the initial listing is alphabetical within each type, but subsequent additions append to the end of the type section unless `sort.alphabetical` enforcement is active. |
| Alphabetical only | `false` | `true` | Globally alphabetical, ignoring type boundaries. |
| Unstructured | `false` | `false` | No detectable pattern — files are in ad-hoc order. |

### 5.2 Adding a new file (no enforcement)

Without a `[src-properties]` enforcement table:

- **Pattern detected** (e.g. sorted-by-type-then-alphabetical): The tool
  respects the convention but does **not** reorder. The new entry is appended
  at the end. The user can manually reposition it.
- **Unstructured**: Append at the end.

With `[src-properties]` enforcement (§3):

- The tool actively inserts at the correct sorted position and rewrites
  `[src]`.

### 5.3 Blank lines between type groups

Blank lines between type groups are a **presentation detail** — the tool
tolerates and preserves them but does not enforce them. There is no config
key for this.

- **New file, `--individual`:** A blank line is inserted between each type
  group as a presentation default. This is a one-time formatting choice at
  creation time.
- **Existing file:** `patchToml()` preserves whatever blank lines are already
  present. If a user removes them, the tool won't add them back.
- **Inserting a new file (with `sort.by-types` enforcement):** The new entry
  goes after the last file of the same type. Existing blank lines around the
  insertion point are preserved by `patchToml()` — the tool doesn't add or
  remove them.

---

## 6. Structure Detection

A new helper module (`src/manifest/src-sort.ts`) analyses an existing `[src]`
section and returns a `SrcStructure` descriptor. This is used when **reading**
a manifest to understand how the user has organised their file, so that write
operations (add, export) can respect the existing convention.

### 6.1 `SrcStructure` type

```ts
/** Describes how the [src] section is currently organised. */
export interface SrcStructure {
    /** true when [src] uses the 3 reserved grouped keys (Modules, Forms, Classes). */
    grouped: boolean;

    /** true when all .bas files are contiguous, all .frm contiguous, all .cls contiguous.
     *  The order of the type groups themselves is not prescribed
     *  (e.g. Forms before Modules is still "sorted by types"). */
    sortedByTypes: boolean;

    /** true when all entries are in alphabetical order, globally. */
    sortedAlphabetically: boolean;

    /** Convenience: true when both sortedByTypes AND within-each-type alphabetical. */
    sortedByTypeThenAlphabetically: boolean;

    /** true when none of the above patterns are detected.
     *  NEW files are appended to the end when unstructured. */
    unstructured: boolean;

    /** When grouped, the raw glob strings keyed by type. */
    groupedPatterns?: Record<"Modules" | "Forms" | "Classes", string | string[]>;
}
```

### 6.2 Detection logic

```
function detectSrcStructure(src: Source[]): SrcStructure
```

1. **Grouped check:** If `src` has exactly 3 entries named `Modules`, `Forms`,
   `Classes` (case-insensitive), and each value looks like a path/glob (not a
   concrete `.bas`/`.frm`/`.cls` file path), then `grouped = true`. All other
   flags are `false` / irrelevant.

2. **SortedByTypes check:** Walk the array. Track the current component type
   (document, module, form, class). Each time the type changes, record a new
   "segment". If at most 4 segments exist and each segment contains only one
   type, `sortedByTypes = true`. (Note: until `"document"` type detection is
   implemented, only 3 segments are possible — `.bas`, `.frm`, `.cls`.)

3. **SortedAlphabetically check:** Walk the array. If every entry's name is
   `>=` the previous entry's name (case-insensitive), `sortedAlphabetically = true`.

4. **SortedByTypeThenAlphabetically:** `sortedByTypes && sortedAlphabetically
   && withinEachSegmentAlphabetical`.

5. **Unstructured**: None of the above patterns match.

### 6.3 Usage

Called by:
- `parseManifest()` in `src/manifest/index.ts` — stores the result on the
  `Manifest` object as `manifest.srcStructure`.
- `applyChangeset()` in `src/build/apply-changeset.ts` — reads
  `manifest.srcStructure` to decide where to insert new entries.

---

## 7. Behaviour by Operation

### 7.1 `vba init --from` / `vba new`

| Step | Action |
|---|---|
| 1 | Export all components from the source workbook/add-in. |
| 2 | Write `[src]` with the 3 grouped keys and default globs: `Modules = "src/**/*.bas"`, `Forms = "src/**/*.frm"`, `Classes = "src/**/*.cls"`. No `[src-properties]` section is written — the grouped convention speaks for itself. |
| 3 | Save the extracted source files to disk under `src/`. |

If the user explicitly passes `--individual` (or sets a flag in an answer
file), the tool uses individual listing instead: sort components by type then
name, write individual entries, and insert blank lines between type groups.

### 7.2 `vba export`

| Step | Action |
|---|---|
| 1 | Load the existing manifest and detect its convention via `detectSrcStructure()`. |
| 2 | Compare with the freshly exported build graph. |
| 3 | **Grouped convention:** No changes to `[src]` needed (the globs still cover all files). Only add/remove individual source files on disk. |
| 4 | **Individual convention, no enforcement:** Append new entries at the end; remove deleted entries. The user can reorder manually. |
| 5 | **Individual convention, with enforcement:** Insert new entries at the correct position per `[src-properties]`; remove deleted entries. |
| 6 | Write back with `patchToml()` — this preserves blank lines and comments already in the file. |

### 7.3 `vba add`

| Convention | `[src-properties]` present? | Action |
|---|---|---|
| Grouped | N/A | Create the file on disk; the glob already covers it. No manifest change. |
| Individual | No | Detect convention, append new entry at end. No reordering. |
| Individual | Yes (enforcement) | Insert at the correct sorted position per `sort.*` settings, rewrite `[src]`. |

### 7.4 `vba build`

No changes. The build graph is always assembled by loading all source files
resolved from the manifest. Internal ordering of the build graph is
**deterministic but not a user-facing sort** — it exists solely so that
two builds of the same source produce byte-identical output. It does **not**
reorder the user's manifest entries or override `[src-properties]` settings.

The principle: the tool never imposes an ordering the user didn't ask for.
If `[src-properties]` is absent, no reordering is enforced — files are used
as they are found or listed.

---

## 8. TOML Formatting: Blank Lines

### 8.1 How blank lines work today

The project uses `@decimalturn/toml-patch`:
- **`patch(existing, value)`** — merges changes into an existing TOML string,
  preserving comments, blank lines, and key order.
- **`stringify(value)`** — generates a TOML string from scratch for new files.

Blank lines between type groups are a *presentation* concern, not a semantic
one. The TOML spec does not define key ordering or spacing.

### 8.2 Strategy

| Scenario | Mechanism |
|---|---|
| **New file, grouped** (`vba init`, `vba new` — the default) | Only 3 keys in `[src]`; no blank-line post-processing needed. |
| **New file, individual** (`vba init --individual`) | After `stringify()`, post-process the `[src]` section to insert blank lines between type groups (a one-time presentation default). |
| **Existing file** (`vba export`, `vba add`) | `patch()` preserves whatever blank lines are already present. If the user removes them, they stay removed. |

### 8.3 Post-processing for new files

A helper `insertTypeGroupBlankLines(toml: string): string`:
1. Split the TOML string into lines.
2. Identify the `[src]` section boundaries.
3. Scan entries within `[src]`; whenever the file extension changes (`.bas` →
   `.frm` or `.frm` → `.cls`), insert an empty line before that entry.
4. Rejoin.

---

## 9. Implementation Plan

### Phase 1 — Foundation

| # | File | Change |
|---|---|---|
| 1 | `src/manifest/src-sort.ts` | **New file.** Define `SrcStructure`, `SrcProperties`, `detectSrcStructure()`, `insertTypeGroupBlankLines()`. |
| 2 | `src/manifest/index.ts` | Parse `[src-properties]` in `parseManifest()`. Add `srcProperties` and `srcStructure` fields to `Manifest`. Export them in `formatManifest()`. |
| 3 | `src/manifest/source.ts` | Update `formatSrc()` to accept `SrcProperties` and `SrcStructure`. When grouped, emit the 3 reserved keys. When individual, sort according to properties. |
| 4 | `src/build/component.ts` | Keep `byComponentName` and `byComponentTypeThenName`. Add `byComponentType` (type-only, no alphabetical tiebreak) for the `sort.by-types=true, sort.alphabetical=false` case. |

### Phase 2 — Consumers

| # | File | Change |
|---|---|---|
| 5 | `src/build/apply-changeset.ts` | In `updateManifest()`, read `manifest.srcStructure`. For structured mode, insert at sorted position. For unstructured, append. For grouped, skip manifest update. |
| 6 | `src/build/load-from-export.ts` | After loading, apply sort based on `manifest.srcProperties` (not hardcoded `byComponentTypeThenName`). |
| 7 | `src/build/load-from-project.ts` | Same as above — use configurable sort. |
| 8 | `src/build/compare-build-graphs.ts` | Keep internal `byComponentTypeThenName` for deterministic changeset ordering (this is an internal concern, not user-facing). |

### Phase 3 — Wiring & init

| # | File | Change |
|---|---|---|
| 9 | `src/actions/init-project.ts` | When generating a fresh manifest, write the 3 grouped keys in `[src]` with default globs. Do **not** write a `[src-properties]` section (the convention is self-evident). |
| 10 | `src/bin/vbapm-init.ts` | Accept `--individual` flag to use individual listing at init time instead of the grouped default. |
| 11 | `src/bin/vbapm-new.ts` | Same `--individual` flag. |

### Phase 4 — Tests

| # | Area | What to test |
|---|---|---|
| 12 | `src/manifest/__tests__/` | Parsing `[src-properties]` (all keys, defaults, invalid values). |
| 13 | `src/manifest/__tests__/` | `detectSrcStructure()` — grouped, sorted-by-types, sorted-alphabetical, combo, unstructured. |
| 14 | `src/manifest/__tests__/` | `formatSrc()` output for each mode. |
| 15 | `src/build/__tests__/` | `applyChangeset` insertion position (structured vs unstructured). |
| 16 | `src/build/__tests__/` | Build graph internal sort unaffected by user-facing settings. |
| 17 | E2E | `vba init --from` generates grouped `[src]` by default (3 keys, no `[src-properties]`); `vba init --from --individual` generates individual listing with blank-line separators. |

### Phase 5 — Migration

| # | Action |
|---|---|
| 18 | Existing projects without `[src-properties]` are left untouched; `detectSrcStructure()` infers their convention on read. |
| 19 | Update the `vbapm add` command to read `srcStructure`: skip manifest update when grouped; append (don't reorder) when individual without enforcement. |
| 20 | Update `vba export` to use `patchToml()` and respect detected convention (no enforcement unless `[src-properties]` is present). |

---

## 10. Migration from Current Implementation

### What stays the same

- `byComponentName()` and `byComponentTypeThenName()` remain in
  `src/build/component.ts` and continue to be used for **internal** build-graph
  ordering (deterministic component order during build).
- `patchToml()` continues to preserve existing formatting in the user's
  `vbaproject.toml`.
- Existing projects **without** `[src-properties]` are not modified; the tool
  detects their convention via `detectSrcStructure()` and works with whatever
  pattern it finds.

### What changes

- New projects (`vba init`, `vba new`) default to the **grouped convention** —
  3 keys in `[src]` with glob patterns. No `[src-properties]` is written.
- `formatSrc()` becomes aware of both the detected convention and any
  `[src-properties]` enforcement settings.
- `updateManifest()` in `apply-changeset.ts`: appends when no enforcement is
  configured; inserts at sorted position when enforcement is active; skips
  entirely for grouped convention.
- A new `detectSrcStructure()` helper analyses the current `[src]` order.
- Users can opt into individual listing at init time with `--individual`.
- `[src-properties]` is now purely an enforcement/validation mechanism, not
  a descriptor of what the file already looks like.

### Risk: `apply-changeset.ts` currently appends unsorted

The current `updateManifest()` pushes new sources to the end of the array,
which means after the first build the `[src]` section loses any sorted order
it may have had. Phase 2, item 5 fixes this: when `[src-properties]`
enforcement is active, entries are inserted at the correct sorted position;
otherwise, the detected convention is respected but no reordering is forced.

---

## 11. Open Questions

*None at this time.*
