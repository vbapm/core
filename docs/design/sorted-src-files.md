# Source Files — Listing, Grouping & Sorting

> **Status:** Design (replaces the narrow "Sorted Source Files" plan).
> The previous implementation of `byComponentTypeThenName` is already in the code;
> this document expands on it with a configurable `[source]` table.

---

## 1. Motivation

VBA projects are more like a cohesive collection of scripts than a formally
compiled assembly. The philosophy of listing every file individually in
`vbaproject.toml` — inherited from C#-style `.csproj` manifests — is overly
verbose for VBA. This design introduces **two complementary listing modes**
controlled by a new `[source]` table:

| Mode | Key | Description |
|---|---|---|
| **Grouped** | 4 reserved keys | Use wildcard patterns under up to 4 reserved keys (`Objects`, `Modules`, `Forms`, `Classes`). `Objects` and `Classes` both match `.cls` files and are pooled — the user controls how to split them (e.g. by subdirectory). **The default for new projects.** |
| **Individual** | one key per file | List every source file by name. The order is governed by the detected convention, optionally enforced via `[source]`. |

In *both* modes the manifest remains easy to scan and the tooling knows how to
add, remove, and reorder entries predictably.

---

## 2. How the Tool Knows What Convention Is in Use

The tool does **not** need a config table to tell it what convention is already
in the file. It infers the convention automatically from the `[src]` entries
themselves, using `detectSrcStructure()` (§6).

| What the tool sees in `[src]` | Detected convention |
|---|---|
| Exactly 3 or 4 keys named from `{Objects, Modules, Forms, Classes}` with glob, array, or literal path values | **Grouped** |
| Individual entries, all `.bas` contiguous, all `.frm` contiguous, all `.cls` contiguous | **Sorted by types** |
| Individual entries, globally alphabetical | **Sorted alphabetically** |
| Individual entries, contiguous by type AND alphabetical within each type | **Sorted by type then alphabetically** |
| None of the above | **Unstructured** |

Once detected, the tool respects this convention on subsequent operations
(add, export) — it doesn't reorder or reformat unless explicitly told to.

---

## 3. The `[source]` Table (Optional Enforcement)

The `[source]` section is **optional**. Its purpose is not to describe
what convention is already in place (the tool detects that automatically), but
to tell the tool to **actively enforce** a specific ordering — reordering
entries on add/export, validating structure, and so on.

If `[source]` is absent, the tool is hands-off: it detects the
convention and works with it, but never reorders.

If `[source]` is present, the tool enforces what the keys say.

### 3.1 Key reference

| Key | Type | Description |
|---|---|---|
| `grouping` | `boolean` | Enforce grouped mode: `[src]` must have exactly the 3 reserved keys. If a user accidentally adds an individual entry, the tool flags it. |
| `sort.by-types` | `boolean` | (Individual mode) Enforce type grouping: Objects → Modules → Forms → Classes. On add/export, insert new entries into the correct type section. |
| `sort.alphabetical` | `boolean` | (Individual mode) Enforce alphabetical order. When combined with `sort.by-types`, alphabetical is *within* each type group; alone, it's global. |
| `subfolders` | `{ Modules?, Forms?, Classes?, Objects? }` | Maps component types to subdirectories under `src/` (§4.6). |

All keys default to `false` when the `[source]` table is present but
a key is omitted — i.e., you opt in to each enforcement individually.

### 3.2 Example: enforcement in individual mode

```toml
[source]
sort.by-types = true
sort.alphabetical = true
```

With this in place, `vba add MyModule` will insert the new entry in the
correct alphabetical position within the modules section, rather than
appending at the end.

### 3.3 Example: no enforcement (the common case)

```toml
# No [source] section at all.
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

The grouped convention uses up to four reserved keys in `[src]` — no
`[source]` needed. The tool detects this convention automatically.
`Objects` and `Classes` both match `.cls` files and are pooled; the user
controls how to split them (typically by subdirectory).

```toml
[src]
Objects = "src/Excel Objects/*.cls"
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
`sort.alphabetical` keys from `[source]` are respected for ordering
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

The `subfolders` key in `[source]` controls where new files are
placed on disk (used by `vba add` and `vba export`). It supports four keys:

```toml
[source]
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

Every source file has its own key in `[src]`. No `[source]` is needed —
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
| By type only | `true` | `false` | Files grouped by extension but in **insertion order** within each group — the order files were added to the project. When creating a new config with `--list-all`, the initial listing is alphabetical within each type, but subsequent additions append to the end of the type section unless `sort.alphabetical` enforcement is active. |
| Alphabetical only | `false` | `true` | Globally alphabetical, ignoring type boundaries. |
| Unstructured | `false` | `false` | No detectable pattern — files are in ad-hoc order. |

### 5.2 Adding a new file (no enforcement)

Without a `[source]` enforcement table:

- **Pattern detected** (e.g. sorted-by-type-then-alphabetical): The tool
  respects the convention but does **not** reorder. The new entry is appended
  at the end. The user can manually reposition it.
- **Unstructured**: Append at the end.

With `[source]` enforcement (§3):

- The tool actively inserts at the correct sorted position and rewrites
  `[src]`.

### 5.3 Blank lines between type groups

Blank lines between type groups are a **presentation detail** — the tool
tolerates and preserves them but does not enforce them. There is no config
key for this.

- **New file, `--list-all`:** A blank line is inserted between each type
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
export interface SrcStructure {
    sortedByTypes: boolean;
    sortedAlphabetically: boolean;
    sortedByTypeThenAlphabetically: boolean;
    unstructured: boolean;
}
```

### 6.2 Detection logic

1. **SortedByTypes check:** Walk the array tracking file extensions. Reject
   non-contiguous extensions (e.g. `.bas → .frm → .bas` is not type-sorted).
   At most 4 segments.

2. **SortedAlphabetically check:** Every entry's name ≥ previous.

3. **SortedByTypeThenAlphabetically:** Both checks pass and within each
   segment entries are alphabetical.

4. **Unstructured**: None of the above.

### 6.3 Usage

Called by `parseManifest()` and stored as `manifest.srcStructure`.

---

## 7. Behaviour by Operation

### 7.1 `vba init --from` / `vba new`

Writes wildcard entries by default:
`Modules = "src/**/*.bas"`, `Forms = "src/**/*.frm"`, `Classes = "src/**/*.cls"`.
Use `--list-all` for individual listing.

### 7.2 `vba export`

Detects convention; appends new entries without enforcement, inserts at
sorted position with enforcement. `patchToml()` preserves existing formatting.

### 7.3 `vba add`

Wildcard coverage check (new file already matched by existing glob → skip
manifest update). Without enforcement, append. With enforcement, insert at
sorted position.

### 7.4 `vba build`

Internal ordering is deterministic for byte-identical output. Does not
reorder manifest entries or override `[source]`.
  ordering (deterministic component order during build).
- `patchToml()` continues to preserve existing formatting in the user's
  `vbaproject.toml`.
- Existing projects **without** `[source]` are not modified; the tool
  detects their convention via `detectSrcStructure()` and works with whatever
  pattern it finds.

### What changes

- New projects (`vba init`, `vba new`) default to the **grouped convention** —
  3 keys in `[src]` with glob patterns. No `[source]` is written.
- `formatSrc()` becomes aware of both the detected convention and any
  `[source]` enforcement settings.
- `updateManifest()` in `apply-changeset.ts`: appends when no enforcement is
  configured; inserts at sorted position when enforcement is active; skips
  entirely for grouped convention.
- A new `detectSrcStructure()` helper analyses the current `[src]` order.
- Users can opt into individual listing at init time with `--list-all`.
- `[source]` is now purely an enforcement/validation mechanism, not
  a descriptor of what the file already looks like.

### Risk: `apply-changeset.ts` currently appends unsorted

The current `updateManifest()` pushes new sources to the end of the array,
which means after the first build the `[src]` section loses any sorted order
it may have had. Phase 2, item 5 fixes this: when `[source]`
enforcement is active, entries are inserted at the correct sorted position;
otherwise, the detected convention is respected but no reordering is forced.

---

## 11. Open Questions

*None at this time.*
