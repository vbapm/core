/**
 * A minimal counting semaphore used to bound concurrency of a resource.
 *
 * The e2e suite uses it to cap the number of Excel COM instances that are
 * alive at once (see `excel-pool.ts`), so parallel test workers share a
 * bounded pool of Excel instances instead of spawning one per test.
 */
export class Semaphore {
	private available: number;
	private waiters: Array<() => void> = [];

	constructor(count: number) {
		this.available = count;
	}

	/**
	 * Acquire a permit, waiting (asynchronously) if none is currently available.
	 * Returns a release function; call it exactly once when done with the slot.
	 */
	async acquire(): Promise<() => void> {
		if (this.available > 0) {
			this.available -= 1;
			return () => this.release();
		}

		await new Promise<void>(resolve => {
			this.waiters.push(resolve);
		});

		// A permit became available for us; reserve it.
		return () => this.release();
	}

	private release(): void {
		const next = this.waiters.shift();
		if (next) {
			// Hand the freed slot directly to the next waiter (no gap).
			next();
		} else {
			this.available += 1;
		}
	}
}

/**
 * Run `fn` while holding a permit from `sem`, releasing it in all cases
 * (including when `fn` throws).
 */
export async function withPermit<T>(sem: Semaphore, fn: () => Promise<T> | T): Promise<T> {
	const release = await sem.acquire();
	try {
		return await fn();
	} finally {
		release();
	}
}
