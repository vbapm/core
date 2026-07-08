# VBA Addin-to-Addin References

## Problem Statement

vbapm currently supports references to COM/DLL type libraries (like Outlook, Scripting, Word) via GUID + version. However, when a VBA project references another VBA addin (`.xlam`) or VBA project (`.xlsm`), the export/import pipeline breaks:

- **Export**: Produces `version = "0.0"`, `guid = ""` because VBA project references don't have COM GUIDs or version numbers
- **Import**: `References.AddFromGuid("", 0, 0)` silently fails — the addin reference is lost
- **Changeset detection**: Since all VBA project references have identical `guid=""` and `version="0.0"`, the comparison logic cannot detect when an addin reference changed

### Real-world example (EmailManager)

The `demo/EmailManager` project references `AddinToolbox`, another VBA addin. Running `vbapm export` on EmailManager produces:

```toml
[references]
AddinToolbox = { version = "0.0", guid = "" }
```

This is invalid for re-import because `AddFromGuid` can't resolve an empty GUID.

---

## How VBA References Work

### Two distinct reference types

| Property | COM/DLL Type Library | VBA Project Reference |
|---|---|---|
| `Ref.Guid` | Valid registry GUID (`{...}`) | Empty string `""` |
| `Ref.Major` | Numeric version (e.g. 9) | Always `0` |
| `Ref.Minor` | Numeric version (e.g. 6) | Always `0` |
| `Ref.FullPath` | Path to DLL/OCX | Path to `.xlam`/`.xlsm` file |
| Import API | `References.AddFromGuid()` | `References.AddFromFile()` |

### Current code path (where it breaks)

**Export** (`addins/src/Build.bas`, `ExportTo`):
```vb
RefInfo("guid") = Ref.Guid      ' "" for VBA project refs
RefInfo("major") = Ref.Major    ' 0 for VBA project refs
RefInfo("minor") = Ref.Minor    ' 0 for VBA project refs
RefInfo("name") = Ref.Name      ' "AddinToolbox"
```

**Import** (`addins/vba-installer/src/Installer.bas`, `AddReference`):
```vb
Project.References.AddFromGuid Guid, MajorVersion, MinorVersion
' Fails silently when Guid="" and versions are 0
```

**Manifest** (`src/manifest/reference.ts`):
```ts
export interface Reference {
    name: string;
    guid: string;
    major: number;
    minor: number;
}
// No field for: path, type (COM vs project), or isAddin flag
```

---

## Feasibility: Can we programmatically add a peer reference?

### `References.AddFromFile` works for VBA projects

Yes. VBA's `References.AddFromFile(path)` can add a reference to another VBA project
(.xlam or .xlsm) **from any workbook**, including a blank one. This is the same mechanism
Excel uses under the hood when you use Tools → References → Browse in the VBA IDE.

### What's available at export time

VBA's `Reference` object exposes `Ref.FullPath` for **all** reference types:

| Reference Type | `Ref.FullPath` | `Ref.Guid` | `Ref.Major/Minor` |
|---|---|---|---|
| COM/DLL | Path to `.dll` or `.ocx` | Valid GUID | Numeric (e.g. 9.6) |
| VBA Project | Path to `.xlam` / `.xlsm` | `""` (empty) | `0`, `0` |

**Current gap**: `Build.ExportTo` reads `Guid`, `Major`, `Minor` but **never reads
`FullPath`**. The path exists at runtime but is discarded during export.

### What's needed for import

`Installer.bas` currently only has `AddReference(Project, Guid, Major, Minor)` which
calls `AddFromGuid`. A new method (or branching logic) using `AddFromFile` is needed
for VBA project references.

### The absolute path problem in built files

Even with perfect TOML portability, VBA stores project-to-project references as
**absolute paths** inside the `.xlsm`/`.xlam` binary. The VBE References dialog
*always* displays the full absolute path — there is no special "name-only" mode,
even for addins installed in `%APPDATA%\Microsoft\AddIns\`.

When `References.AddFromFile("C:\\Users\\Alice\\...\\AddinToolbox.xlam")` runs at
build time, that path is baked into the file forever.

```mermaid
flowchart LR
    TOML["vbaproject.toml<br/>AddinToolbox = { peer = true }"]
    Build["vbapm build"]
    Resolve["resolve peer path<br/>C:\\Users\\Alice\\...\\AddinToolbox.xlam"]
    VBA["References.AddFromFile(...)"]
    Binary["built .xlsm<br/>(absolute path baked in)"]
    TOML --> Build --> Resolve --> VBA --> Binary
```

This means the **vbaproject.toml** is portable, but the **built .xlsm** is not.

**How cross-machine resolution actually works**:

Excel resolves VBA project references in two steps:
1. **By name first**: If a VBProject with the same name is already loaded in the
   current Excel session, use it (regardless of the stored path)
2. **By path fallback**: If not already loaded, try the stored absolute path

So Bob *can* use Alice's built `.xlsm` if the peer addin is already loaded in his
Excel session — either opened manually or installed as an Excel addin. Excel matches
by VBProject name and ignores the stale path.

**Scenarios**:

| Scenario | Outcome |
|---|---|
| Alice builds and runs on her machine (path matches) | ✅ Works |
| Alice builds, sends `.xlsm` to Bob, addin NOT loaded | ❌ `MISSING: AddinToolbox` |
| Alice builds, sends `.xlsm` to Bob, addin IS loaded | ✅ Resolves by VBProject name |
| Bob has addin installed as Excel addin (auto-loaded) | ✅ Resolves by name on startup |
| CI builds, developer opens with addin loaded | ✅ Resolves by name |
| CI builds, developer opens without addin loaded | ❌ MISSING |

**Mitigations**:
- **For development**: Each developer loads the peer addin into Excel (or installs
  it as an Excel addin) — then the reference resolves by name, path irrelevant
- **For distribution**: Document that recipients must either install the peer addin
  or rebuild from source with vbapm
- **vbapm can't fix this**: It's VBA/Excel runtime behavior. The TOML stays portable;
  the built binary is only portable if the peer is loaded in Excel

---

## Design Decisions Needed

### 1. How to identify a VBA-project reference vs a COM reference

VBA project references have `guid = ""` and `major = 0`, `minor = 0`. These sentinel
values reliably distinguish them from COM references (COM type libraries always have
GUIDs).

### 2. What to store in `vbaproject.toml` for VBA project references

Since VBA project references don't have GUIDs or versions, we need different metadata.
Options:

**Option A: Explicit path**
```toml
[references]
AddinToolbox = { path = "../AddinToolbox/build/AddinToolbox.xlam" }
```
- Pro: Can resolve at build time via relative path; no magic
- Con: Tight coupling to filesystem layout; breaks if addin moves; verbose

**Option B: Package-based (vbapm registry)**
```toml
[references]
AddinToolbox = { package = "AddinToolbox", version = "^1.0" }
```
- Pro: Proper dependency management like npm/cargo
- Con: Requires a package registry; large scope increase; not MVP

**Option C: Hybrid — path with fallback**
```toml
[references]
AddinToolbox = { path = "../AddinToolbox", type = "project" }
```
- Pro: Works immediately with existing `[dependencies]`-like path resolution
- Con: Still filesystem-dependent; `type` key is redundant if path is present

**Option D: Named project reference (workspace resolution)**
```toml
[references]
AddinToolbox = { project = "AddinToolbox" }
```
- vbapm resolves the addin name to a project in the workspace or registry
- Requires workspace-level project discovery

**Option E: Peer flag** ⭐ (simplest)
```toml
[references]
AddinToolbox = { peer = true }
```
- Minimal syntax — just declares "this is a peer VBA project reference"
- vbapm resolves the path at build time by:
  1. Searching sibling directories for `vbaproject.toml` with matching `[project].name`
  2. Building the peer project first (if needed)
  3. Using the built `.xlam` path for `AddFromFile`
- No GUID, no version, no hardcoded path — all resolved automatically
- Analogous to npm peer dependencies or Cargo workspace members

**Comparison**:

| | A (path) | B (registry) | C (hybrid) | D (project) | E (peer) |
|---|---|---|---|---|---|
| Minimal TOML syntax | ❌ | ✅ | ❌ | ✅ | ✅ |
| No hardcoded paths | ❌ | ✅ | ❌ | ✅ | ✅ |
| Works without registry | ✅ | ❌ | ✅ | ✅ | ✅ |
| Declarative intent | ❌ | ✅ | ❌ | ✅ | ✅ |
| Implementation complexity | Low | High | Med | Med | Low-Med |

**Recommendation**: Option E (`peer = true`) is the right balance of simplicity and
power for initial implementation. The path can be resolved at build time without
the user needing to know build output paths.

### 3. How to store the reference in the `Reference` interface (TypeScript side)

The current `Reference` interface needs to be extended. A peer reference has no GUID
or version — it's identified by name and resolved at build time.

```ts
export interface Reference {
    name: string;
    guid: string;       // "" for peer references
    major: number;      // 0 for peer references
    minor: number;      // 0 for peer references
    peer?: boolean;     // true = VBA project reference, resolve via workspace
    path?: string;      // resolved at build time, stored transiently
}
```

This is backward-compatible — existing COM references continue to work unchanged.
Peer references add only one optional boolean field.

**TOML representation** for the EmailManager example:

```toml
# COM references (unchanged)
[references]
Outlook = { guid = "{00062FFF-0000-0000-C000-000000000046}", version = "9.6" }
Scripting = { guid = "{420B2830-E718-11CF-893D-00A0C9054228}", version = "1.0" }
Word = { guid = "{00020905-0000-0000-C000-000000000046}", version = "8.7" }

# Peer VBA project reference (new)
AddinToolbox = { peer = true }
```

### 4. How to import VBA project references (VBA side)

Replace the `AddFromGuid`-only approach with type-aware logic. When `Ref("peer")` is true
(or `Ref("guid") = ""`), use `AddFromFile` instead:

```vb
' In Installer.bas or Build.bas
If Ref.Exists("peer") And Ref("peer") = True Then
    Project.References.AddFromFile Ref("path")   ' path resolved by vbapm CLI
Else
    Project.References.AddFromGuid Ref("guid"), CLng(Ref("major")), CLng(Ref("minor"))
End If
```

The `path` is resolved by the TypeScript CLI at build time before passing to VBA.
The VBA side doesn't need to know about workspace layout — it just gets a resolved
absolute path.

### 5. How to resolve the peer addin file path at build time

vbapm resolves `peer = true` references by searching the workspace:

1. Scan parent directories for a vbapm workspace root (where `pnpm-workspace.yaml` or similar lives)
2. Search sibling directories for `vbaproject.toml` files
3. Match `[project].name` to the reference name (`AddinToolbox`)
4. Use the peer project's `build/` output path as the resolved path
5. If the peer hasn't been built yet, build it first (topological order)

Fallback strategies (if workspace search fails):
- Check a `VBAPM_ADDINS_PATH` environment variable
- Look in a configured addins directory
- Warn and skip (broken reference)

### 6. How to detect changes in VBA project references

Current changeset comparison in `compare-build-graphs.ts`:
```ts
reference.guid !== before_reference.guid ||
reference.major !== before_reference.major ||
reference.minor !== before_reference.minor
```

For peer references, all three are always identical (`""`, `0`, `0`). Instead, compare
by the resolved path:
```ts
reference.peer !== before_reference.peer ||
reference.path !== before_reference.path
```

Or, for a more robust check, hash the referenced file's content.

---

## Affected Files

### VBA (addin code)
| File | Change |
|---|---|
| `addins/src/Build.bas` (`ExportTo`) | Export `peer`, `path` fields; detect empty GUID |
| `addins/src/Build.bas` (`ImportGraph`) | Pass `peer` and `path` to installer |
| `addins/vba-installer/src/Installer.bas` | Add `AddProjectReference` using `AddFromFile` |

### TypeScript (CLI/library)
| File | Change |
|---|---|
| `src/manifest/reference.ts` | Add `peer?: boolean`, `path?: string` to `Reference` |
| `src/manifest/index.ts` | `formatReferences`: write `peer = true` for peers; `parseReferences`: handle `peer` key |
| `src/build/load-from-export.ts` | Parse `peer` and `path` from `project.json` |
| `src/build/load-from-project.ts` | Load peer references from manifest |
| `src/build/compare-build-graphs.ts` | Compare peer refs by `path` instead of `guid`/`major`/`minor` |
| `src/build/apply-changeset.ts` | Handle peer reference additions/removals |
| `src/build/stage-build-graph.ts` | Include peer references in staged graph |
| `src/addin.ts` | Serialize `peer` and `path` in JSON passed to VBA |
| `src/build/resolve-peers.ts` *(new)* | Workspace search + path resolution for `peer = true` |

### Tests
| File | Change |
|---|---|
| `src/manifest/__tests__/manifest.test.ts` | Test parsing of `peer = true` references |
| `src/build/__tests__/compare-build-graphs.test.ts` | Test peer reference change detection |
| `src/build/__tests__/stage-build-graph.test.ts` | Test peer reference staging |
| `src/build/__tests__/resolve-peers.test.ts` *(new)* | Test workspace peer resolution |

---

## Implementation Phases

### Phase 0: Exploration (integration tests)
- Implement the integration test scenarios listed above
- Validate that `References.AddFromFile` works for `.xlam` and `.xlsm` targets
- Confirm `Ref.FullPath` behavior at export time
- Confirm that peer references can be detected by `Guid = ""` sentinel
- Answer open questions with real VBA behavior (transitivity, `.xlsm` vs `.xlam`, etc.)

### Phase 1: Detection & Export
- In `Build.ExportTo`, detect when `Ref.Guid = ""` → mark as `peer`
- Export `Ref.FullPath` alongside existing fields in `project.json`
- New `project.json` schema:
  ```json
  {
    "name": "EmailManager",
    "references": [
      { "name": "Outlook", "guid": "{00062FFF...}", "major": 9, "minor": 6 },
      { "name": "AddinToolbox", "guid": "", "major": 0, "minor": 0, "peer": true, "path": "C:\\...\\AddinToolbox.xlam" }
    ]
  }
  ```
- Update `load-from-export.ts` to parse `peer` and `path` fields
- Add `peer?: boolean` and `path?: string` to `Reference` interface
- Update `formatReferences` to write `peer = true` (omit guid/version for peers)
- Update `parseReferences` to handle `peer = true` (set guid="" and version="0.0")

### Phase 2: Import (AddFromFile)
- Add `AddProjectReference` method to `Installer.bas`:
  ```vb
  Public Sub AddProjectReference(Project As VBProject, FilePath As String)
      Project.References.AddFromFile FilePath
  End Sub
  ```
- Update `Build.ImportGraph` to branch: if `Ref("peer")` is true, call `AddProjectReference`
- Implement TypeScript-side path resolution:
  - Search workspace for matching `vbaproject.toml` by `[project].name`
  - Resolve to `<peer_build_dir>/<peer_name>.xlam`
  - Build peer first if needed (topological sort)
- Update `addin.ts` to serialize `peer` and `path` in the JSON passed to VBA

### Phase 3: Changeset Detection
- Update `compare-build-graphs.ts` to compare peer references by `path`
- When `peer === true`, diff on `path` instead of `guid`/`major`/`minor`
- Add tests for peer reference addition, removal, and path changes

### Phase 4: Workspace Resolution (future)
- Implement workspace-level addin discovery for `peer = true` without pre-built artifacts
- Support multi-project builds with dependency ordering
- Add version tracking via file hash or manifest version
- Support `[dependencies]` style resolution for addin packages

---

## Open Questions

1. **Should VBA project references be transitive?** If Project A references Addin B, and Addin B references DLL C, should Project A also get DLL C? Current COM references are transitive in VBA. Peer references may not need transitivity since the referenced addin carries its own COM references.

2. **What about `.xlsm` vs `.xlam`?** Should vbapm distinguish between workbook references and addin references? For `peer = true`, the build output type (`.xlam` vs `.xlsm`) is determined by the peer project's `[project].target`.

3. **Reference name collision**: What if a COM reference and a peer reference have the same name? The TOML table key `[references.Name]` can only appear once. Since peer refs use `peer = true` and COM refs use `guid` + `version`, they can coexist under different keys — but a COM ref and a peer ref to the same addin would need distinct names.

4. **Broken reference handling**: When a referenced addin file doesn't exist, should the build warn, error, or skip? For `peer = true`, if the peer project hasn't been built yet, vbapm should attempt to build it. If the peer project can't be found at all, the build should error with a clear message.

5. **Build order**: Peer references imply build dependencies. vbapm needs a dependency graph: if Project A has `AddinToolbox = { peer = true }`, then `AddinToolbox` must be built before Project A. This requires workspace-level awareness (Phase 4).

6. **Path portability**: `Ref.FullPath` from export gives an absolute path. When writing to TOML, should vbapm convert this to a `peer = true` declaration (portable) or keep the absolute path? Recommendation: convert to `peer = true` during export so the TOML is portable across machines.

7. **What if the user doesn't use a workspace?** A standalone project with `peer = true` needs a way to find the peer. Options: (a) require workspace layout, (b) fall back to a configured path, (c) error with guidance to set up a workspace.

---

## Integration Test Scenarios

The first implementation step should be exploratory integration tests to validate
that `References.AddFromFile` works correctly across all reference combinations.
These tests answer the open questions above with real VBA behavior.

### Test setup

Each test uses a blank `.xlsm` workbook created by `vbapm new`. The vbapm addin
(`vbapm.xlam`) drives the test via `Build.ImportGraph` and `Build.ExportTo`.

### Scenarios

#### 1. Add peer reference to `.xlam` addin
- Create blank `.xlsm` project A
- Build `.xlam` project B (the peer)
- Programmatically add reference from A → B via `AddFromFile`
- **Assert**: Reference appears in `VBProject.References` with `Guid = ""`, `Major = 0`, `Minor = 0`
- **Assert**: VBA code in A can call public functions from B
- **Assert**: `Ref.FullPath` points to B's `.xlam` file

#### 2. Export round-trip: detect and preserve peer ref
- Start with `.xlsm` that has a peer reference to `.xlam`
- Run `ExportTo` → `project.json`
- **Assert**: Exported ref has `guid = ""`, `major = 0`, `minor = 0`
- **Assert**: Exported ref has `peer = true` and `path` field
- Parse into `Reference` interface, format to TOML
- **Assert**: TOML writes `peer = true` (not `guid`/`version`)
- Parse TOML back, resolve peer path, run `ImportGraph`
- **Assert**: Reference is re-added successfully (not silently dropped)

#### 3. Import round-trip: build from TOML with `peer = true`
- Create vbaproject.toml with `AddinToolbox = { peer = true }`
- Have `AddinToolbox.xlam` built and present in workspace
- Run `vbapm build`
- **Assert**: `AddFromFile` is called with resolved path
- **Assert**: Built `.xlsm` has working reference to AddinToolbox
- **Assert**: Reference is NOT `MISSING` when opened in Excel

#### 4. Peer ref to `.xlsm` (not `.xlam`)
- Same as scenario 1 but peer is `.xlsm`
- **Assert**: `AddFromFile` works for `.xlsm` targets too
- **Question**: Does VBA treat `.xlsm` refs differently from `.xlam` refs?

#### 5. Broken peer ref: file missing
- Create project with `peer = true` to nonexistent addin
- Run build
- **Assert**: vbapm errors with clear message (not silent failure)
- **Assert**: error message includes the peer name and search paths tried

#### 6. Peer ref with both projects open in VBA IDE
- Simulate user having both workbooks open in Excel
- Add peer reference while target is open
- **Assert**: Reference resolves correctly
- **Assert**: Closing/reopening host workbook preserves the reference

#### 7. Name-resolution theory: pre-loaded addin resolves stale path
- Build host `.xlsm` with peer reference to AddinToolbox at path A
- Close Excel entirely, then reopen host `.xlsm` **without** AddinToolbox loaded
- **Assert**: Reference shows `MISSING: AddinToolbox` (stored path doesn't exist or is wrong)
- Now open AddinToolbox `.xlam` in the **same** Excel session (not as installed addin —
  just `Application.Workbooks.Open`)
- **Assert**: `MISSING` reference **auto-resolves** by VBProject name
- **Assert**: VBA code in host can call AddinToolbox functions
- **Assert**: `Ref.FullPath` now shows the **actual** path where AddinToolbox was opened
  (Excel updates the stored path to match the loaded file)
- Close AddinToolbox, keep host open
- **Assert**: Reference stays resolved (path was updated by Excel in step 6)
- Save host `.xlsm` — the reference now points to the new path
- **Key insight tested**: Excel resolves by VBProject name first; stored path is a
  fallback. Once resolved, Excel rewrites the stored path to the loaded file's location.

#### 8. Changeset detection: peer path changes
- Build project with peer ref at path A
- Move peer addin to path B, update workspace
- Run export → compare graphs
- **Assert**: Changeset detects the reference as `changed` (path differs)
- **Assert**: `compare-build-graphs` does NOT incorrectly treat it as unchanged
  (which would happen with current `guid`/`major`/`minor` comparison)

#### 9. Mixed COM + peer references in same project
- Project has both `Outlook = { guid = "...", version = "9.6" }` and `AddinToolbox = { peer = true }`
- Run full export → TOML → import cycle
- **Assert**: COM references use `AddFromGuid`, peer refs use `AddFromFile`
- **Assert**: Both types coexist without interference

#### 10. Cross-machine TOML portability (manual verification)
- Build project on Machine 1 (Windows, `C:\Users\Alice\...`)
- Commit vbaproject.toml (contains `peer = true`, no absolute paths)
- Clone and build on Machine 2 (different path)
- **Assert**: Workspace resolution finds the peer at Machine 2's path
- **Assert**: Build succeeds with peer reference at correct path

### Test implementation notes

- These should run as Jest integration tests using the existing e2e test infrastructure
  (`e2e.config.mjs`, `npm run test:e2e:background`)
- Each test creates temp directories, initializes projects, runs vbapm commands
- The VBA addin is used to drive Excel COM automation
- Scenario 10 may need a manual runbook since cross-machine can't be automated in CI
