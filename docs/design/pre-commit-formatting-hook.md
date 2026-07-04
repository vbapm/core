# Pre-commit Formatting Hook with oxlint + oxfmt

## Problem

CI builds frequently fail because contributors forget to run `npm run format`
before pushing. The existing Prettier-based `format:check` script is too slow
to run as a pre-commit hook, and there is no automated enforcement at commit
time.

## Goal

- A **fast** pre-commit hook that blocks commits containing unformatted code.
- Replace **Prettier** with **oxfmt** (Rust-based, 10–100× faster).
- Add **oxlint** for linting (replaces the need for a separate ESLint setup).
- Keep the hook fast enough that it doesn't disrupt the development flow
  (target: < 500 ms for typical commits).

## Why oxlint / oxfmt

| Tool | Written in | Speed (vs JS alternative) |
|---|---|---|
| **oxfmt** | Rust | ~30× faster than Prettier |
| **oxlint** | Rust | ~50–100× faster than ESLint |

Both are designed for pre-commit use. oxfmt is a drop-in replacement for
Prettier with near-identical formatting output. oxlint provides a curated
set of lint rules without requiring configuration.

## Design

### Hook Architecture

```
git commit
  └── lefthook pre-commit
        ├── oxfmt --check (staged .ts/.js files only)
        └── oxlint        (staged .ts files only)
              ↑
              If either fails → commit blocked
```

### Tool: lefthook

[lefthook](https://github.com/evilmartians/lefthook) is a fast Git hooks
manager written in Go. It's the recommended companion for oxlint/oxfmt:

- Single binary, no Node.js dependency chain.
- Native support for staging-area-only checks.
- Auto-install on `pnpm install` via a `postinstall` script.

### Format check (oxfmt)

```
oxfmt --check <staged_files>
```

- Runs only on staged `.ts`/`.js` files to keep it fast.
- `--check` mode: exits non-zero if any file would be reformatted.
- The error message tells the contributor to run `pnpm run format`.

### Lint check (oxlint)

```
oxlint --deny-warnings <staged_files>
```

- Catches common mistakes (unused vars, wrong types, etc.).
- `--deny-warnings` makes warnings fail the hook (consistent with CI).

### Auto-fix escape hatch

A `pnpm run format` script is available to auto-fix all files:

```json
{
  "scripts": {
    "format": "oxfmt --write \"**/*.{ts,js}\"",
    "format:check": "oxfmt --check \"**/*.{ts,js}\"",
    "lint": "oxlint",
    "lint:check": "oxlint --deny-warnings"
  }
}
```

If the hook blocks a commit, the contributor runs `pnpm run format`, stages
the changes, and retries the commit.

## Migration from Prettier

### Config mapping

| Prettier (.prettierrc) | oxfmt (oxfmt.toml / package.json) |
|---|---|
| `useTabs: true` | `indent_style = "tab"` |
| `printWidth: 100` | `print_width = 100` |
| `trailingComma: "none"` | `trailing_commas = "never"` |
| `arrowParens: "avoid"` | `arrow_parens = "avoid"` |

oxfmt also reads `.editorconfig` for `indent_style`, `indent_size`, and
`end_of_line`, so the existing `.editorconfig` settings are respected
automatically.

### Ignored paths

Prettier ignores `lib/` and `worktrees/` via `.prettierignore`. oxfmt
uses `--ignore-path` or an `oxfmt.toml` with an `ignore` list. Same
paths will be carried over.

## Pre-commit Hook Implementation

### lefthook.yml

```yaml
# lefthook.yml
pre-commit:
  parallel: true
  commands:
    format:
      glob: "*.{ts,js}"
      run: oxfmt --check {staged_files}
      stage_fixed: false
    lint:
      glob: "*.ts"
      run: oxlint --deny-warnings {staged_files}
```

### Auto-install hook

In `package.json`:

```json
{
  "scripts": {
    "postinstall": "lefthook install",
    "format": "oxfmt --write \"**/*.{ts,js}\"",
    "format:check": "oxfmt --check \"**/*.{ts,js}\"",
    "lint": "oxlint",
    "lint:check": "oxlint --deny-warnings"
  }
}
```

The `postinstall` script ensures every contributor has the hook installed
after running `pnpm install`.

## CI Integration

The same checks run in CI to catch any bypass attempts:

```yaml
# .github/workflows/ci.yml (addition)
- name: Format check
  run: pnpm run format:check
- name: Lint check
  run: pnpm run lint:check
```

If the pre-commit hook was skipped with `--no-verify`, CI still catches it.

## Rollout Steps

1. **Install dependencies**
   ```bash
   pnpm add -D oxfmt oxlint lefthook @oxlint/oxlint
   ```
   (Note: oxfmt may be published as `oxc-oxfmt` or similar — verify
   the exact package name at implementation time.)

2. **Create oxfmt config** (`oxfmt.toml` or via `package.json`)
   - Mirror the existing `.prettierrc` settings.
   - Add ignore paths (`lib/`, `worktrees/`).

3. **Create lefthook.yml** with the pre-commit commands.

4. **Update package.json scripts**
   - Replace `prettier` with `oxfmt` in `format` / `format:check`.
   - Add `lint` / `lint:check` scripts.
   - Add `postinstall: "lefthook install"`.

5. **Run initial format**
   ```bash
   pnpm run format
   ```
   Commit any formatting differences oxfmt produces vs. Prettier.

6. **Remove Prettier**
   ```bash
   pnpm remove prettier
   ```
   Delete `.prettierrc` and `.prettierignore`.

7. **Update CI workflow** to call `format:check` and `lint:check`.

8. **Test the hook** — make an intentional formatting mistake and verify
   the commit is blocked.

## Edge Cases

- **Partial commits** (`git add -p`): lefthook's `{staged_files}` only
  passes the staged version, so a partially staged file is checked
  correctly.
- **Merge commits**: lefthook can be configured to skip merge commits
  (`skip: merge`).
- **Emergency bypass**: `git commit --no-verify` skips the hook. CI
  still catches it.
- **Windows**: oxlint/oxfmt binaries are cross-platform. lefthook
  works on Windows via its Go binary.
- **First-time contributor**: `pnpm install` auto-installs the hook.
  No manual setup needed.

## Future Considerations

- **oxfmt v1**: As of July 2026, oxfmt is still pre-1.0. Monitor for
  breaking config changes and update the migration guide accordingly.
- **oxlint rules**: Start with the default rule set. Additional rules
  can be enabled in `oxlintrc.json` as the team identifies pain points.
- **CI caching**: Cache the oxlint/oxfmt binaries in CI for faster runs.
