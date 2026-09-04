# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Deprecated
- `vba export` command is deprecated in favor of `vba extract`. It still functions but prints a deprecation warning and will be removed in a future version.
- The top-level `[src]` section in `vbaproject.toml` is deprecated in favor of `[source.files]`. Legacy manifests remain valid: `[src]` entries are treated as `[source.files]`, but a deprecation warning is printed until the manifest is updated.

### Added
- `vba extract` command as the new canonical name for extracting source from a built target (replaces `vba export`).
- `include-empty-objects` option under `[source]` (default: `true`). When `false`, blank document objects (`ThisWorkbook`, sheet modules) are skipped during export, matching the pre-0.10 behavior. When `true`, all document objects are exported regardless of content.
- `[source]` table in `vbaproject.toml` for optional enforcement of source file ordering. Keys: `sort.by-types`, `sort.alphabetical`, `subfolders`. When absent, the tool detects and respects the existing convention without enforcement via `detectSrcStructure()`.
  - `folder` key (default: `"src"`). Controls the base directory for all source files. Set to a relative path like `"src/WorkbookName"` to nest sources under a subfolder. When absent, `"src"` is used implicitly — the key only needs to be written when overriding the default.
  - The `subfolders` key controls where new source files are placed (supports `Modules`, `Forms`, `Classes`, and `Objects` for document-type components like `ThisWorkbook` and `Sheet` modules). Example: `subfolders = { Modules = "Modules", Forms = "Forms", Classes = "Class Modules", Objects = "Excel Objects" }`.
  - The `sort.by-types` option will ensure that all Modules, Classes, Forms and Objects are kept together in `[src]`.
- Wildcard support in `[src]` entries (e.g. `Modules = "src/**/*.bas"`). Wildcards follow `minimatch` syntax and are applied to any key. This is the default for new projects. Use `--list-all` to opt into individual listing instead.
- `build-dir` field in `vbaproject.toml` to specify where the built `.xlsm`/`.xlam` is written. Defaults to `"build"` when omitted. Set to `"."` to output in the project root.
- `vba init --from workbook.xlsm` automatically sets `build-dir = "."` when the workbook is at the project root.
- `vbaproject.toml` now validates section keys and suggests corrections for snake_case misspellings (e.g. `build_dir` → `build-dir`, `src_encoding` → `src-encoding`).
- Peer references in `[references]` for referencing another VBA project (addin or workbook). A peer reference has no GUID or version and is declared as `AddinToolbox = { peer = true, path = "..." }`. `vba extract` detects peer references (empty GUID) and stores them with a path — relative when the peer lives inside the project folder or a sibling folder, absolute otherwise. `vba build` resolves the path and adds the reference via `References.AddFromFile`. Peer paths are written with forward slashes even on Windows (e.g. `path = "C:/Users/me/Toolbox/Toolbox.xlam"`).
- Global and project-local `vba.toml` settings for background Excel execution, with explicit CLI flags taking precedence over local and global configuration. Background mode is propagated explicitly through build, update, extract/export, and macro execution instead of relying on `VBA_BACKGROUND_BUILD`.

### Changed
- New projects (`vba init`, `vba new`) now default to wildcard entries (e.g. `Modules = "src/**/*.bas"`) instead of listing every source file individually.
- Source files are now written under `[source.files]` in `vbaproject.toml` instead of the top-level `[src]`. New manifests use `[source.files]`, and the deprecated `[src]` section is migrated to `[source.files]` on write.
- Wildcard entries in `[src]` now suppress duplicate individual entries on extract. When a component's path is already covered by an existing wildcard pattern, no redundant individual `[src]` listing is added, keeping `vbaproject.toml` clean.
- `src-encoding` has moved from `[project]`/`[package]` to `[source]` as `encoding`. Existing `vbaproject.toml` files with `src-encoding` under `[project]` or `[package]` will now receive a clear error message suggesting the migration. Per-source `encoding` on `[src]` entries is unchanged.

### Fixed
- Missing target files now report whether the target directory is missing and suggest `target.name` when matching workbook files are found.
- Name conflict resolution on extract: when a `[src]` entry (e.g. `Module1`) points to a different component name (`Validation.bas`) but the workbook also contains a real `Module1`, both entries are preserved with a warning about the rename.
- VBA addin errors now surface to the CLI via structured JSON output instead of being silently swallowed by `Err.Raise`.
- Windows path separators are normalized when comparing wildcard-covered paths during extract.
- Document objects (`ThisWorkbook`, sheet modules) are now placed in the correct subfolder (e.g. `src/Excel Objects/`) instead of being misclassified as class modules.
- `Workbooks.Open` in the VBA addin no longer triggers VBA events, preventing external addins from interrupting CLI operations.
- Referencing another VBA addin no longer produces a broken `version = "0.0", guid = ""` entry on extract. VBA project references are now preserved as peer references.


## [0.9.0-pre] - 2026-07-06

### Added
- Multilingual encoding support for VBA source files across Windows ANSI codepages (CP874, CP932, CP936, CP949, CP950, CP1250–CP1258). ([#103])
  - Source files can declare their encoding via `src-encoding` in `[project]`/`[package]` or per-source `encoding`.
  - Target encoding can be declared via `encoding` on `[project] target`.
  - Build-time validation fails with a `jschardet` suggestion when non-ASCII characters are detected without a declared encoding.
  - Transcoding during build (source → target encoding) and extract (system codepage → source encoding).
  - Extract warns when transcoding to a non-UTF encoding could lose characters.
  - When extracting from an existing workbook with `vba init --from`, encoding is auto-detected and written to `vbaproject.toml`.
- `vba init` and `vba new` now create starter `.gitignore`, `.gitattributes`, and `.editorconfig` files unless `--no-conf` flag is used. ([#64] and [#98]).
- New `open` command to open the current built target file in Excel ([#63]).
- New `close` command to close the current built target file in Excel, with optional `--save` flag ([#63]).
- Excel XML export: renames worksheet XML files to stable, identity-based names (`sht{codeName}.xml`) so that reordering sheets produces clean diffs with only ordering metadata changed ([#57]).

### Fixed
- XML formatting now uses CRLF line endings to match the OOXML standard ([#62]).
- Opening workbooks now matches by full path to avoid picking up same-named files from different directories.
- Minor XML export bug related to empty `dc:creator` element.

### Changed
- Switched package manager from npm to pnpm v11.

### Security
- Replaced shell-based `exec()` with `execFile()` (macOS) and `spawn()` (Windows) in bridge script runner to prevent shell injection from user-controlled arguments. ([#75])

## [0.8.0] - 2026-04-05

### Added
- XML formatting transformation that pretty-prints `.xml` and `.rels` files on export ([#47])
- Add `--xml-only` and `--vba-only` options when exporting ([#54])
- New `update` command that allows updating VBA source code directly into an existing built target file ([#56]).

### Fixed
- Extracting now preserves extra metadata fields from `[project]` and `[package]` sections instead of silently discarding them ([#47]).

## [0.7.0] - 2026-03-31

### Added
- New `add` command support to create and register source files ([#46]).
- Better error message when the file extension is included in the VBA Component name ([#39]).
- Add support for .xlam and .xlsx files ([#40]).

### Fixed
- Improved duplicate source detection across `src` and `dev-src`([#46]).
- Added VBA component name validation ([#46]).

## [0.6.15] - 2026-02-25

### Added
- Added installer submodule and integrated local installer workflow.

### Changed
- Updated Rollup build setup, including ESM rollup config usage.

### Fixed
- Fixed nested module resolution for vendored `archiver`.

## [0.6.14] - 2026-02-25

### Fixed
- Included `archiver` as a vendored dependency in release builds.

## [0.6.12] - 2026-02-23

First release of vbapm after forking vba-blocks v0.5.3.
Note: The online registry remains the vba-blocks.com registry.

### Changed
- Name was changed from "vba-blocks" to "vbapm"
- Project is now dual: CLI-tool and NPM package. NPM Package should be the preferred installation method.
- Update Node to v22
- Update dependencies (all moderate to critical vulnerabilities were resolved)
- VBScript replaced with PowerShell

[Unreleased]: https://github.com/DecimalTurn/vba-blocks/compare/v0.7.0...HEAD
[0.7.0]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.15...v0.7.0
[0.6.15]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.14...v0.6.15
[0.6.14]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.12...v0.6.14
[0.6.12]: https://github.com/DecimalTurn/vba-blocks/releases/tag/v0.6.12
[#40]: https://github.com/vbapm/core/pull/40
[#46]: https://github.com/vbapm/core/pull/46
[#39]: https://github.com/vbapm/core/pull/39
[#47]: https://github.com/vbapm/core/pull/47
[#54]: https://github.com/vbapm/core/pull/54
[#56]: https://github.com/vbapm/core/pull/56
[#63]: https://github.com/vbapm/core/pull/63
[#62]: https://github.com/vbapm/core/pull/62
[#64]: https://github.com/vbapm/core/pull/64
[#98]: https://github.com/vbapm/core/pull/98
