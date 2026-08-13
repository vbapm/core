# The `.block` File Format

## Overview

A `.block` file is a **ZIP archive** with a `.block` extension. Renaming one to `.zip` and opening it with any archive tool will reveal its contents. The format packages a VBA package for distribution and installation via the vbapm/vba-blocks registry.

## Naming Convention

Block files follow this naming pattern:

```
<sanitized-package-name>-v<version>.block
```

Examples:
- `json-v2.3.0.block`
- `dictionary-v1.4.1.block`

## Internal Structure

A `.block` archive contains a flat list of files (no subdirectories):

| File | Required | Description |
|------|----------|-------------|
| `vba-block.toml` | Yes | Package manifest (name, version, authors, source map, dependencies) |
| `*.bas`, `*.cls`, `*.frm` | Yes | VBA source files declared under `[source.files]` in the manifest |
| `LICENSE` / `LICENCE` | No | License file |
| `README.md` | No | Readme |
| `CHANGELOG.md` / `HISTORY.md` | No | Changelog |
| `NOTICE` | No | Notice file |

### Example: `dictionary-v1.4.1.block`

```
Dictionary.cls
vba-block.toml
```

`vba-block.toml`:
```toml
[package]
name = "dictionary"
version = "1.4.1"
authors = ["Tim Hall <tim.hall.engr@gmail.com> (https://github.com/timhall)"]
license = "MIT"

[source.files]
Dictionary = "Dictionary.cls"
```

### Example: `json-v2.3.0.block`

```
JsonConverter.bas
LICENSE
vba-block.toml
```

`vba-block.toml`:
```toml
[package]
name = "json"
version = "2.3.0"
authors = ["Tim Hall <tim.hall.engr@gmail.com> (https://github.com/timhall)"]
license = "MIT"

[source.files]
JsonConverter = "JsonConverter.bas"

[dependencies]
dictionary = "^1"
```

## Creating a `.block` File

Use the `pack` script from the project root:

```
node scripts/pack <path/to/package/dir>
```

This reads `vbaproject.toml` in the given directory, collects the source files and any readme/license/changelog files, and zips them into `<dir>/build/<name>-v<version>.block`.

Use `--force` to overwrite an existing block:

```
node scripts/pack <path/to/package/dir> --force
```

## Inspecting a `.block` File

Use the `unpack` script to extract a block:

```
node scripts/unpack tests/__fixtures__/.vbapm/packages/vba-blocks/json-v2.3.0.block
# Extracts to: tests/__fixtures__/.vbapm/packages/vba-blocks/json-v2.3.0/

node scripts/unpack tests/__fixtures__/.vbapm/packages/vba-blocks/dictionary-v1.4.1.block ./output/dictionary
# Extracts to: ./output/dictionary/
```

Alternatively, rename the file to `.zip` and open it with any archive tool.
