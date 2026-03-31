# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.7.0] - 2026-03-31

### Added
- New `add` command support to create and register source files.
- Better error message when the file extension is included in the VBA Component name.

### Fixed
- Improved duplicate source detection across `src` and `dev-src`.
- Added VBA component name validation.

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

[Unreleased]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.15...HEAD
[0.6.15]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.14...v0.6.15
[0.6.14]: https://github.com/DecimalTurn/vba-blocks/compare/v0.6.12...v0.6.14
[0.6.12]: https://github.com/DecimalTurn/vba-blocks/releases/tag/v0.6.12
