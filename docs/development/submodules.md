# Git Submodules

This project uses **git submodules** to pull in external dependencies. These commands are useful when working with submodules, especially inside git worktrees.

## Current Submodules

| Path | URL |
|------|-----|
| `installer` | https://github.com/vbapm/installer.git |
| `src/actions/templates/VBA-on-GitHub` | https://github.com/DecimalTurn/VBA-on-GitHub.git |

## Common Commands

### Check submodule status

Shows which submodules exist and what commit they are pinned to (a `-` prefix means not initialized):

```powershell
git submodule status
```

### Initialize and clone submodules

After cloning or checking out a branch that references submodules, run this to pull them down:

```powershell
git submodule update --init --recursive
```

#### What `--init` and `--recursive` do

| Flag | Purpose |
|------|---------|
| `--init` | If a submodule hasn't been cloned yet (status shows a `-` prefix), this flag initializes it by cloning the repository. Without `--init`, `git submodule update` only updates **existing** submodules — it skips uninitialized ones. |
| `--recursive` | If any submodule **itself** contains submodules (nested submodules), this flag ensures they are also initialized and updated. Without `--recursive`, only the top-level submodules are handled. |

In short: `--init` handles first-time cloning, `--recursive` handles nested submodules. Always use both together unless you have a specific reason not to.

### Sync submodule URLs then update

If the `.gitmodules` file has changed (e.g. after switching branches or pulling), sync the URLs first, then update:

```powershell
git submodule sync
git submodule update --init --recursive
```

### One-liner (sync + update)

```powershell
git submodule sync ; git submodule update --init --recursive
```

### Update submodules to latest remote commit

Pull the latest commit from each submodule's tracked branch (default: `main`):

```powershell
git submodule update --remote --recursive
```

### Full submodule loop (iterates each one)

```powershell
git submodule foreach "git fetch origin && git checkout main && git pull"
```

## Worktree Usage

When working inside a **git worktree** (`worktrees/<branch-name>`), submodules are **not shared** with the main repository. Each worktree needs its own submodule setup.

### New worktree — clone submodules

```powershell
cd worktrees/<branch-name>
git submodule update --init --recursive
```

### Worktree — sync + update

```powershell
cd worktrees/<branch-name>
git submodule sync
git submodule update --init --recursive
```

### Worktree — full one-liner from repo root

```powershell
Push-Location worktrees/<branch-name>; git submodule sync; git submodule update --init --recursive 2>&1; Pop-Location
```

## Troubleshooting

### "Cannot find path" error in PowerShell

If you see an error like:

```
Set-Location: Cannot find path '...\worktrees\<branch>\worktrees\<branch>' because it does not exist.
```

This happens when a `cd` or `Set-Location` inside a `&&`/`;` chain tries to change directory relative to the wrong location. Solution: use `Push-Location` / `Pop-Location` or `cd` into the worktree first, then run the command in a separate line.

### Submodule shows modified status unexpectedly

Submodules can show as modified if the checked-out commit differs from what the parent repo expects. To reset:

```powershell
git submodule update --force --recursive
```

To see what changed:

```powershell
git diff --submodule
```

#### "New commits" change (submodule is ahead of parent)

If `git status` shows `modified: <submodule> (new commits)`, it means the submodule is checked out to a commit **newer** than what the parent repo references. This can happen after running `git submodule update --remote` or committing directly inside the submodule.

**To reset a single submodule** back to what the parent expects:

```powershell
git submodule update --force <submodule-path>
```

For example, to reset only the `installer` submodule:

```powershell
git submodule update --force installer
```

**To reset all submodules** at once:

```powershell
git submodule update --force --recursive
```

**To identify the difference** between what the parent expects and what's checked out:

```powershell
git diff --submodule <submodule-path>
```

For example:

```powershell
git diff --submodule installer
```

This shows the commits that were added (or removed) relative to the parent's expected commit.
