/**
 * Per-worker teardown for the persistent PowerShell/Excel session.
 *
 * Registered via `setupFilesAfterEnv`, so this runs once per test *file* (and
 * therefore once per Jest worker process).
 *
 * Why not `globalTeardown`: that hook runs in Jest's main process, which never
 * owns a session. When `VBA_PERSISTENT_SESSION=1` + `E2E_IN_PROCESS=1` are set,
 * the session (and the hidden Excel it owns) belongs to the *worker*, so it has
 * to be shut down from inside the worker or it leaks one EXCEL.EXE per file.
 *
 * Two singletons must be closed because the e2e suite loads `run.ts` twice:
 *   - `vbapm`              → mapped to `lib/index.js` (used by the `run()` helper)
 *   - `src/utils/run.ts`   → used by in-process command dispatch
 * Each module instance holds its own session, so closing only one still leaks.
 *
 * When no session was ever started (the default paths), `closePowerShellSession`
 * is a no-op, so this is safe for every e2e script.
 */
afterAll(async () => {
	const closers = [
		import("vbapm").then(m => m.closePowerShellSession),
		import("../../src/utils/run").then(m => m.closePowerShellSession)
	];

	// Never fail a passing suite over teardown: a session that already exited
	// (or was never started) must not turn into a red test run.
	const results = await Promise.allSettled(
		closers.map(async loaded => {
			const close = await loaded;
			await close();
		})
	);

	for (const result of results) {
		if (result.status === "rejected") {
			console.warn(`[e2e] failed to close PowerShell session: ${result.reason}`);
		}
	}
});
