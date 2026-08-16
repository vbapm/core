/**
 * Stub for the ESM-only `open` package.
 *
 * The real `open` package is ESM-only (uses `import` statements), which Jest
 * cannot transform in-process. Commands that transitively import `open` (via
 * `open-target` → `openTarget`) would otherwise cause a "Cannot use import
 * statement outside a module" parse error.
 *
 * In-process e2e tests never actually need to launch an external application:
 * - The `open` CLI command is routed to the spawned CLI in `execute()`, so the
 *   real `open` is exercised there (in a separate `node` process, unaffected by
 *   Jest's module resolution).
 * - `build`/`extract`/`update` import `openTarget` for static wiring but never
 *   call it.
 */
export default async function openStub(_target: string, _options?: unknown) {
	return { exitCode: 0 };
}
