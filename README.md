# vbapm

A package manager and build tool for VBA. It lets you version-control your VBA source code and Excel XML structure as plain text files, then build them back into a working workbook or add-in.

Support for other MS Office file format is planed.

## Installation

### Option 1: Install as a global npm package

**Prerequisites:** [Node.js](https://nodejs.org/) v22 or higher.

```txt
npm install -g vbapm
```

### Option 2: Standalone installer

<details>
<summary>Installation instructions</summary>

**Windows**

In powershell, run the following:

```txt
iwr https://raw.githubusercontent.com/vbapm/installer/refs/heads/main/install.ps1 | iex
```

**Mac**

In terminal, run the following:

```txt
curl -fsSL https://raw.githubusercontent.com/vbapm/installer/refs/heads/main/install.sh | sh
```

For more recent versions of Office for Mac, you will need to trust access to the VBA project object model for vbapm to work correctly:

<details>
  <summary>Trust access to the VBA project object model</summary>
  <ol>
    <li>Open Excel</li>
    <li>Click "Excel" in the menu bar</li>
    <li>Select "Preferences" in the menu</li>
    <li>Click "Security" in the Preferences dialog</li>
    <li>Check "Trust access to the VBA project object model" in the Security dialog</li>
 </ol>
</details>



</details>

<hr>

If you run into any issues during installation, please see the [known issues](https://github.com/vbapm/installer#known-issues) for the installer or [create a new issue](https://github.com/vbapm/installer/issues/new) with details about what's happening.

:rocket: You're ready to go! Open a new command-line session (cmd / terminal) and try `vba --help`

> **macOS users:** You must enable "Trust access to the VBA project object model" in Excel → Preferences → Security for vbapm to work.

## Getting Started

### Initialize a project from an existing workbook

If you already have an `.xlsm` or `.xlam` file and want to start managing it with vbapm, use the `--from` flag:

```txt
vba new my-project --from "C:\path\to\existing.xlsm"
```

If you want to start in the current directory instead of creating a new one:

```txt
vba init
```

This creates a new directory `my-project/` containing:

```
my-project/
├── vbaproject.toml          # Project manifest (name, target type, sources)
├── src/                     # Extracted VBA source files
│   ├── Module1.bas          #   Standard modules → .bas
│   ├── Sheet1.cls           #   Document modules → .cls
│   ├── ThisWorkbook.cls     #   Workbook module → .cls
│   ├── MyClass.cls          #   Class modules → .cls
│   └── UserForm1.frm        #   UserForms → .frm
├── target/                  # Extracted Excel XML structure
│   ├── [Content_Types].xml
│   └── xl/
│       ├── workbook.xml
│       ├── worksheets/
│       ├── styles.xml
│       └── ...
├── build/                   # Built workbook (customizable via build-dir in vbaproject.toml)
│   └── my-project.xlsm
└── .git/, .gitignore, ...   # Git version control (by default)
```

vbapm automates two extraction steps:

1. **VBA source code** — The vbapm add-in runs a macro that extracts every `VBComponent` from the workbook as its native file type, along with a `project.json` listing non-built-in references.

2. **XML structure** — The workbook is unzipped and its internal XML (sheets, styles, ribbon, etc.) is extracted to `target/`.

> Use `vba init --from existing.xlsm` instead of `vba new` if you want to initialize in the *current* directory rather than creating a new one.

### The basic workflow

Once your project is initialized, the day-to-day loop is:

1. **Edit** your source files in `src/` and register new ones with [`vba add`](#add)
2. **Build** the workbook with [`vba build`](#build)
3. **Test** your macros in Excel, or use [`vba run`](#run) from the command line
4. **Extract** changes back to source with [`vba extract`](#extract), then `git diff` and commit

### Putting it together

```txt
# Start with a new project from an existing file
vba new expense-tracker --from C:\workbooks\budget-2025.xlsm
cd expense-tracker

# Make edits to src/ files, then rebuild
vba build --open

# After tweaking in Excel, capture the changes
vba extract

# Commit your work
git add . && git commit -m "Add expense tracker project"
```

### Programmatic Usage

You can also use `vbapm` as a library (e.g. from a VS Code extension):

```js
const { buildProject, loadProject, env } = require("vbapm");

// Override working directory
env.cwd = "/path/to/project";

const project = await loadProject();
await buildProject(project);
```

## Usage

### `new`

Create a new folder with a blank/generated vbapm project inside

Create a folder "project-name" with a blank xlsm project:

```txt
vba new project-name.xlsm
```

(equivalent to above)

```txt
vba new project-name --target xlsm
```

Create a folder "from-existing" with a project from an existing workbook:

```txt
vba new from-existing --from existing.xlsm
```

Create a blank package for sharing as a library between projects:

```txt
vba new json-converter --package
```

By default, `vba new` initializes a git repository and creates `.gitignore`, `.gitattributes` and `.editorconfig` template files. Use `--no-git` to skip git init and `--no-conf` to skip the template files:

```txt
vba new project-name.xlsm --no-git
vba new project-name.xlsm --no-conf
```

### `init`

Create a blank/generated vbapm project in the current folder

Create a blank xlsm project with the current folder's name:

```txt
vba init --target xlsm
```

Create a project from an existing workbook:

```txt
vba init --from existing.xlsm
```

Create a blank package:

```txt
vba init --package
```

By default, `vba init` initializes a git repository and creates `.gitignore`, `.gitattributes` and `.editorconfig` template files. Use `--no-git` to skip git init and `--no-conf` to skip the template files:

```txt
vba init --target xlsm --no-git
vba init --target xlsm --no-conf
```

### `add`

Create a new source file in `src/` and register it in `vbaproject.toml`.

Create a standard module:

```txt
vba add Module1
```

Create a class module:

```txt
vba add JsonParser --type class
```

Register an existing source file by path (no overwrite):

```txt
vba add .\src\Test.bas
```

Create and register a file in a nested path (missing folders are created):

```txt
vba add .\src\features\auth\LoginHelper.bas
```

Add a development-only source to `[dev-src]`:

```txt
vba add TestHelpers --dev
```

### `build`

Build an Excel workbook from the project's source. The built file is written to the directory specified by `build-dir` in `vbaproject.toml` (defaults to `build/`). If a previously built file is found it is moved to a `.backup` directory under the build directory to protect against losing any previously saved work.

To output the built file in the project root instead, set `build-dir = "."` in `vbaproject.toml`.

Build a project:

```txt
vba build
```

Build and open a project for editing:

```txt
vba build --open
```

Build a package using a blank target:

```txt
vba build --target xlsm
```

Build a project, excluding any development src, dependencies, or references:

```txt
vba build --release
```

### `extract`

Once you've completed your edits and are ready to commit your changes, extract your project with `vba extract`.

Extract a project:

```txt
vba extract
```

Extract a previously-built package:

```txt
vba extract --target xlsm
```

Only extract the XML files (skip VBA source extraction):

```txt
vba extract --xml-only
```

Only extract the VBA source (skip XML extraction):

```txt
vba extract --vba-only
```

> **Note:** `vba export` is deprecated and will be removed in a future version. Use `vba extract` instead.

### `update`

`vba update` writes the current VBA source directly into an existing built target file (including one currently open in Excel), without going through a full build cycle.

Update VBA source in the built target:

```txt
vba update
```

Update VBA in a specific target type:

```txt
vba update --target xlsm
```

Update excluding dev-src, dev-dependencies, and dev-references:

```txt
vba update --release
```

Update and leave the target open in Excel after updating:

```txt
vba update --open
```

### `open`

`vba open` opens the current built target file in Excel. This is a convenience shortcut when you want to start editing a target you've already built.

Open the built target:

```txt
vba open
```

Open a specific target type:

```txt
vba open --target xlsm
```

### `close`

`vba close` closes the built target file that is currently open in Excel. By default changes are **discarded** — pass `--save` to keep them.

Close the built target (discard unsaved changes):

```txt
vba close
```

Close and save changes:

```txt
vba close --save
```

Close a specific target type:

```txt
vba close --target xlsm
```

### `run`

`vba run` opens the built workbook in Excel, runs a public VBA function, captures its return value and prints it to stdout.

#### Syntax

```txt
vba run <Module.Function> [<arg1> <arg2> ...] [--file PATH] [--target TYPE]
```

#### VBA code conventions

The `vba run` command calls a **public function** (not a `Sub`). The function:
- Must be `Public`
- Must return a `String` (or a type that VBA can coerce to a string)
- Can accept up to 10 arguments (passed positionally from the CLI)
- Arguments arrive as `Variant` on the VBA side

```vb
' (Module: Messages.bas)
Public Function SayHi(Name As Variant) As String
  SayHi = "Howdy " & Name & "!"
End Function
```

```txt
vba run Messages.SayHi Tim
Howdy Tim!
```

#### Passing arguments

Arguments are passed positionally and arrive as `Variant` in VBA:

```txt
# Single argument
vba run ExcelManipulator.WriteCell "Sheet1,5,3,Hello World"

# Multiple arguments
vba run Math.Add 3 7

# Quoted strings with spaces
vba run Messages.SayHi "Tim Hall"
```

#### Return value

If the VBA function returns a string, it is printed to stdout. This is how you receive data back from Excel.

If the function returns a JSON string matching `{"success": true/false, "messages": [...], "errors": [...]}` it is parsed as a structured result.

#### Targeting options

Run against the built project (default — run from project directory):

```txt
vba run ExcelManipulator.AddSheet "MyNewSheet"
```

Run against a specific file:

```txt
vba run Messages.SayHi Tim --file "C:\path\to\workbook.xlsm"
```

Run against a specific target type (if project has multiple targets):

```txt
vba run MyModule.MyFunction arg1 --target xlsm
```

#### Practical workflow for Excel manipulation

```
vba add ExcelManipulator
# → edit src/ExcelManipulator.bas with your VBA functions
vba build
vba run ExcelManipulator.AddSheet "Report"
vba run ExcelManipulator.WriteCell "Report,1,1,Title"
vba run ExcelManipulator.GetCell "Report,1,1"
# → output: "Title"
```

Example VBA module for cell manipulation:

```vb
Public Function AddSheet(SheetName As Variant) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = CStr(SheetName)
    AddSheet = "Created sheet: " & SheetName
End Function

Public Function WriteCell(Args As Variant) As String
    Dim parts() As String
    parts = Split(CStr(Args), ",")
    ThisWorkbook.Sheets(parts(0)).Cells(CLng(parts(1)), CLng(parts(2))).Value = parts(3)
    WriteCell = "Wrote '" & parts(3) & "' to " & parts(0) & "!" & parts(1) & "," & parts(2)
End Function

Public Function GetCell(Args As Variant) As String
    Dim parts() As String
    parts = Split(CStr(Args), ",")
    GetCell = CStr(ThisWorkbook.Sheets(parts(0)).Cells(CLng(parts(1)), CLng(parts(2))).Value)
End Function
```

> For complex operations, consider encoding data as JSON strings in arguments and parsing with a VBA JSON library.

## Tips

- **Build before running** — `vba run` targets the built workbook in `build/`, not the source files.
- **Extract after manipulation** — Changes made by macros (new sheets, cell values) live in the built workbook. Run `vba extract` to persist them back to source.
- **Close before rebuilding** — If a workbook is open in Excel, close it with `vba close --save` before running `vba build` again.
- **Run from the project root** — Commands expect `vbaproject.toml` in the current directory, unless you use `--file` to target a specific workbook.
- **Quote arguments with spaces** — Shell escaping is handled by vbapm, but wrap arguments containing spaces in quotes.

## Manifest (vbaproject.toml)

The vbapm manifest (vbaproject.toml) serves as the foundation for your project and provides information on your package, source, dependencies, references, and targets, as detailed below.

### [project] or [package]

The `[package]` / `[project]` section includes general information about your package. You should choose `[package]` if your project is only intended to be used as a utility inside another project and `[project]` if your project is a standalone tool.

Here are the main properties:
- `name` (_required_): The name of the project/package
- `version` (_required_ for `[package]`): version number following [semver](https://semver.org/).
- `authors` (_required_ for `[package]`): name(s) of the author(s)
- `target` (_required_ for `[project]`): define what application/extension to use when building your project.
- `codename` custom VBA project name shown in the VBE (defaults to `"VBAProject"`)

**Example 1**
```toml
[project]
name = "awesome-excel-project"
target = "xlsm"
codename = "AwesomeProject"
```

**Example 2**
```toml
[package]
name = "awesome-vba-package"
authors = ["Me <me@email.com>"]
version = "0.1.0"
```

#### version
vbapm follows [Semantic Versioning](https://semver.org/). Make sure you adopt a compatible versioning approach if you intend to publish to the repository.

#### target
`target` is used to define what application/extension to use when building your project. It can be a string for the extension, in which case `target/` includes the source files for creating the target. Otherwise, `type` and `path` can be used to define a custom target path.

Example 1:
```toml
target = "xlsm"
# equivalent to target = { type = "xlsm", path = "target" }
```

Example 2:
```toml
target = { type = "xlam", path = "targets/xlam" }
```

#### codename

`codename` sets the VBA project name displayed in the Visual Basic Editor (VBE). When omitted, the default `"VBAProject"` is used. Only valid for `[project]` — ignored for `[package]`.

```toml
codename = "MyCustomProject"
```


#### [build-dir]
`build-dir` specifies where the built `.xlsm` / `.xlam` file is written. Defaults to `"build"` when omitted. Set to `"."` to output in the project root.

```toml
build-dir = "."    # output to project root instead of build/
```

### [src]

Will contain the list of source code files to be included in the VBA-Enabled Document at build time.

Two listing modes are supported:

**Individual entries** — each file is listed explicitly by name (`name = "path"` or `{ path }`):

```toml
[src]
A = "src/A.bas"
B = "src/B.cls"
C = { path = "src/C.bas" }
```

Use this mode with `vba init --list-all` or `vba init --from <workbook>` for full control over which files are tracked.

**Wildcard entries** — groups of files are discovered by glob pattern (the default for `vba init`):

```toml
[src]
Objects = "src/Excel Objects/**/*.cls"
Forms = "src/Forms/**/*.frm"
Modules = "src/Modules/**/*.bas"
Classes = "src/Class Modules/**/*.cls"
```

Wildcards are expanded at build time — any file on disk matching the glob is picked up automatically, so you don't need to run `vba add` after creating a new module. When you delete a module in the VBE and run `vba extract`, the corresponding file on disk is removed automatically.

> **Note:** Wildcard and individual entries can coexist. Components already covered
> by an existing wildcard pattern do not get duplicate individual entries added which means the wildcard handles discovery on the next build or extract.

### [src-properties]

Optional settings that control how source files are organised and written to disk.

```toml
[src-properties]
# folder = "src"            # Implicit overridable default. 
subfolders = { Modules = "Modules", Forms = "Forms", Classes = "Classes" }
```

| Key | Default | Description |
|---|---|---|
| `folder` | `"src"` | Base directory for all source files (relative to project root). New components extracted from a workbook are placed under this directory. |
| `subfolders` | — | Organises source files into type-based subdirectories: `.bas` → `Modules/`, `.frm` → `Forms/`, `.cls` → `Classes/`. `Objects` maps document modules (`ThisWorkbook`, sheets) to a subfolder. |
| `encoding` | — | Default encoding for source files (e.g. `"cp1252"`, `"utf-8"`). Can be overridden per-source in `[src]`. |
| `sort` | — | Enforce ordering in the `[src]` section. Options: `by-types` (group by component type), `alphabetical` (sort by name). |
| `include-empty-objects` | `true` | When `false`, blank document objects (ThisWorkbook, sheet modules) are skipped during export. |

### [dependencies]

The dependencies section lists the packages that your project depends on. Each dependency can be specified in one of three ways:

`name = "version"` or

- `version`
- `path`
- `git` (and `branch`, `tag`, or `rev`)

```toml
[dependencies]
a = "1" # Equivalent to ^1
b = "=2.0.0" # Precisely 2.0.0
c = { version = "3" }

d = { path = "./packages/d" }

e = { git = "https://..." } # master
f = { git = "https://...", branch = "dev" }
g = { git = "https://", tag = "bugfix" }
h = { git = "https://", rev = "abc1234" }
```

### [references]

- `version` (`"MAJOR.MINOR"`)
- `guid` (`"{...}"`)

```toml
[references]
Scripting = { version = "1.0", guid = "{...}" }
```

### [dev-src,dependencies,references]

`[dev-src]`, `[dev-dependencies]`, and `[dev-references]` are included during development and are excluded when building with the `--release` flag (i.e. `vba build --release`)

## Development

### Prerequisites

1. `git clone` this repo
2. Install [Node.js](https://nodejs.org/) v22 or later
   - Note: For CLI builds, Node v23+ only supports Windows x64 (win-x64). 32-bit Windows (win-x86) is no longer available upstream.
3. Install node-gyp dependencies for [Mac](https://github.com/nodejs/node-gyp#on-macos) or [Windows](https://github.com/nodejs/node-gyp#on-windows)

### Build

1. Run `pnpm install`
2. Run `pnpm run format`
3. Run `pnpm run build:cli`
   <br>It will build the CLI/library in `lib`, plus ensured vendor node runtime is available.
4. Run `pnpm run build:addins`
   <br>It will build the Excel addin that performs workbook/VBA operations from inside Office.

### Test

1. Run `pnpm test`
   <br>It will run unit tests
2. Run `pnpm run test:e2e`
   <br>It will run the end-to-end CLI scenarios in excel.e2e.ts, covering workflows like build, extract, new, and version against fixtures.
   <br>To keep temporary e2e work folders for manual inspection, set `KEEP_E2E_TMP=1` before running (PowerShell: `$env:KEEP_E2E_TMP=1; pnpm run test:e2e`, cmd: `set KEEP_E2E_TMP=1 && pnpm run test:e2e`).
   <br>To echo each e2e command output even on successful runs, use `--verbose` (PowerShell: `pnpm run test:e2e:background -- --verbose`) or set `E2E_VERBOSE=1` (PowerShell: `$env:E2E_VERBOSE=1; pnpm run test:e2e:background`).

### Install local version

To install the local version you can use the `devinstall.ps1` script available in the installer submodule.

```powershell
# Update submodule
git submodule update --init --recursive installer

# Run devinstall
.\installer\devinstall.ps1
```

### Release

1. Run `pnpm version`
2. Run `pnpm run release`

## Acknowledgments

This project is a fork of the original [vba-blocks](https://github.com/vba-blocks/vba-blocks) project by [Tim Hall](https://github.com/timhall).
