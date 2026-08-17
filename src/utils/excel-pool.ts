import { Semaphore, withPermit } from "./semaphore";

/**
 * A shared, process-local cap on the number of concurrent Excel-driven runs.
 *
 * The e2e suite normally runs serially (`--runInBand`), but can be run with
 * Jest workers in parallel. Each `vba run` invocation spawns (or attaches to)
 * an Excel.Application instance; without a bound, N parallel workers would
 * spawn N simultaneous Excel processes and exhaust the machine. This module
 * provides a single counting semaphore that all `run()` calls go through, so
 * the number of live Excel instances never exceeds `VBA_EXCEL_POOL_SIZE`.
 *
 * Size is resolved at first use from `VBA_EXCEL_POOL_SIZE` (default 4). Set it
 * to `1` to force fully serialized Excel access, or higher for beefy CI.
 */
const POOL_SIZE = (() => {
	const raw = process.env.VBA_EXCEL_POOL_SIZE;
	const n = raw ? parseInt(raw, 10) : 0;
	return Number.isFinite(n) && n >= 1 ? n : 4;
})();

export const excelPool = new Semaphore(POOL_SIZE);

/**
 * Run `fn` only once a slot in the Excel instance pool is free (waiting
 * otherwise), and release the slot afterward.
 */
export async function withExcelSlot<T>(fn: () => Promise<T> | T): Promise<T> {
	return withPermit(excelPool, fn);
}
