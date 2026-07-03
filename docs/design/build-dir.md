# Custom Build Directory — Design Decision

## Motivation

By default, `vba build` writes output (`.xlsm` / `.xlam`) to a `build/` subfolder. The `build-dir` field in `vbaproject.toml` lets users override this — for example, to keep the built file in the project root alongside the source workbook.

## Manifest Schema

`build-dir` is an optional key under `[project]` / `[package]`, following the kebab-case convention used by other multi-word keys (`dev-src`, `dev-dependencies`, etc.).

```toml
# vbaproject.toml
[project]
name = "EmailManager"
target = "xlsm"
build-dir = "."        # output to project root instead of build/
```

When omitted, the default is `"build"` — fully backward compatible.

## TypeScript Interface

The property is `buildDir` on the `Manifest` interface, matching the camelCase convention used for all other parsed keys.

```ts
// src/manifest/index.ts
export interface Manifest extends Snapshot {
    // ...
    buildDir?: string;
}
```

## Parsing and Serialization

| Operation | Detail |
|---|---|
| `parseManifest()` | Destructures `"build-dir"` from the TOML section, normalizes via `normalize()` (strips trailing slashes). |
| `formatManifest()` | Round-trips `buildDir` back to `"build-dir"` when writing TOML. Omits it when the value is `"build"` (the default). |

## Path Construction

Both `loadProject()` and `initProject()` in `src/project.ts` use the manifest's `buildDir` to construct the `paths` object:

```ts
const buildDir = manifest.buildDir || "build";

const paths = {
    root: workspace.paths.root,
    dir,
    build: join(dir, buildDir),
    backup: join(dir, buildDir, ".backup"),
    staging: await tmpFolder({ dir: env.staging })
};
```

All downstream consumers reference `project.paths.build`, so they automatically respect the configured directory.

## Action Files

`build-project.ts` and `run-macro.ts` resolve the built file path through the paths object:

```ts
join(project.paths.build, target.filename)
```

## Init from Workbook — Automatic Root Detection

When running `vbapm init --from workbook.xlsm` and the workbook is at the project root (`dirname(from) === dir`), `build-dir` is automatically set to `"."`. This means the built file lands alongside the source workbook without manual configuration.

```ts
// src/actions/init-project.ts
if (from && dirname(from) === dir) {
    project.manifest.buildDir = ".";
    project.paths.build = join(dir, ".");
    project.paths.backup = join(dir, ".", ".backup");
}
```

## Edge Cases

| Case | Behavior |
|---|---|
| `build-dir = "."` | Output in project root; `.backup` becomes `./.backup` |
| `build-dir` omitted | Defaults to `"build"` — fully backward compatible |
| Trailing slash (`"./"`) | Normalized to `"."` via `normalize()` |
| Absolute path | Normalized but not rejected; relative paths are the expected input |
| Package mode (`[package]`) | Supported but rarely useful; not serialized unless explicitly set |

## Files

| File | Role |
|---|---|
| `src/manifest/index.ts` | `buildDir` interface property, parse, format, normalize |
| `src/project.ts` | `loadProject()` / `initProject()` derive paths from `manifest.buildDir` |
| `src/actions/build-project.ts` | Resolves built file via `project.paths.build` |
| `src/actions/run-macro.ts` | Resolves macro target via `project.paths.build` |
| `src/actions/init-project.ts` | Auto-detects root workbook and enforces `buildDir = "."` |
