# src-subfolders Feature

## What

A `src-subfolders` key in `vbaproject.toml` that maps VBA component types to subdirectories under `src/`, so exported source files are organized by type instead of all landing flat in `src/`.

### TOML syntax

```toml
[project]
name = "my-project"
target = "xlsm"
src-subfolder = { Modules = "Modules", Forms = "Forms", Classes = "Classes" }
```

- `.bas` → `src/Modules/`
- `.frm` → `src/Forms/`
- `.cls` → `src/Classes/`

Without the key, behavior is unchanged (all files in `src/`).
