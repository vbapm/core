## Working In A Git Worktree

When developing in a worktree (for example `worktrees/<branch-name>`), each worktree must be prepared independently.

### Required setup inside the worktree

Run all commands from the worktree root:

```powershell
cd worktrees/<branch-name>
pnpm install
pnpm run build:cli
pnpm run build:addins
```

Why this is required:

- Worktrees do not share `node_modules`.
- End-to-end tests execute the local CLI (`bin/vba`), which depends on a built `lib/`.
- Excel integration tests require `addins/build/vbapm.xlam`.

### Running tests from a worktree

Always run tests from inside the worktree directory itself:

```powershell
cd worktrees/<branch-name>
pnpm run dev
pnpm run test:e2e:background
```

Important notes:

- The Jest config ignores `/worktrees/` only when executed from the main repository root.
- When executed from inside a worktree, `<rootDir>` resolves to that worktree, so tests are discovered and run correctly.

For background test commands where location must be preserved, use `Push-Location`:

```powershell
Push-Location worktrees/<branch-name>; pnpm run test:e2e:updateSnapshots 2>&1 | Tee-Object "$env:TEMP\e2e.log"; Pop-Location
```
