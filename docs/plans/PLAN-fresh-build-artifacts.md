# Plan: fresh build artifacts

## What exists today

- `scripts/ensure-fresh-build.ts` checks only `lib/vbapm.js` and `lib/index.js`.
  It uses modification times and runs `pnpm run build:cli` when the sources are
  newer.
- `scripts/lib/hash.js` and `src/utils/hash.ts` already provide SHA-256 hashing.
  `scripts/lib/checksum.js` hashes one file and is used for package checksums.
- No current helper records a hash of a source tree or stores build provenance
  beside `addins/build/vbapm.xlam` or
  `scripts/bootstrap/build/bootstrap.xlsm`.
- `bootstrap.xlsm` is tracked in Git. `lib/` and `addins/build/` are ignored,
  so a fresh checkout cannot rely on those generated files being present.

## Build dependency

The build order is:

1. `ensure-fresh-lib.ts` makes the CLI used by the other scripts current.
2. `ensure-fresh-addin.ts` uses the checked-in bootstrap workbook to build
   `addins/build/vbapm.xlam`.
3. `ensure-fresh-bootstrap.ts` uses the current xlam to build
   `scripts/bootstrap/build/bootstrap.xlsm`.

The two VBA files are bootstrap hosts for each other. They should not be used
as raw binary inputs to each other's freshness hash. Doing that would make a
normal refresh invalidate the other artifact again on every run. Both files
will instead record the same fingerprint for the add-in project and the CLI
toolchain that builds it. The xlam must be current before the bootstrap build
runs.

## Freshness record

Add a small JSON sidecar next to each output, using a name such as
`vbapm.xlam.fresh.json` and `bootstrap.xlsm.fresh.json`.

The record will contain:

- schema version and `sha256` as the algorithm
- the canonical input fingerprint
- the normalized, sorted input paths used to calculate it

The fingerprint will hash each input's normalized repository-relative path and
file bytes in deterministic order. It will cover the add-in source tree,
VBA dependency source trees, both VBA manifests and lockfiles, the build
wrapper scripts and the already-fresh `lib` outputs. Generated build folders
and the two VBA outputs will be excluded. A changed file, renamed file,
removed file or changed build script will therefore produce a new fingerprint.

The existing Node `crypto` SHA-256 helpers can be reused for the hash
calculation. MD5 is unnecessary here. Reading a small source tree and one
sidecar is cheap, and content hashing avoids false freshness caused by copied
files or preserved modification times.

## Script changes

- Rename `ensure-fresh-build.ts` to `ensure-fresh-lib.ts` and update its log
  labels and documentation to describe only the CLI library check.
- Add `ensure-fresh-addin.ts`. It will require the bootstrap host, compare the
  current fingerprint with the xlam sidecar, run `pnpm run build:addins` when
  stale or missing, then write the sidecar only after a successful build.
- Add `ensure-fresh-bootstrap.ts`. It will require the xlam host, compare the
  same current fingerprint with the bootstrap sidecar, run
  `pnpm run build:bootstrap` when stale or missing, then write its sidecar only
  after a successful build.
- Replace `ensure-fresh-build.ts` with a small orchestrator that invokes the
  three scripts in order and forwards a failing exit status. The existing
  `build:check` package script can keep its name and continue to be the prefix
  for the e2e scripts.
- Update the bootstrap refresh workflow to commit the tracked bootstrap
  sidecar with `bootstrap.xlsm`.

## Checks

Add focused tests for deterministic fingerprints, changed and missing inputs,
missing outputs and stale sidecars. Exercise the orchestration with a fresh
checkout, a source edit, an add-in-only rebuild and a bootstrap-only sidecar
recovery. Finish with formatting, type checking, unit tests and the background
e2e suite.
