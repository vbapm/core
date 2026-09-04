export async function measure<T>(label: string, action: () => Promise<T>): Promise<T> {
	const started = performance.now();
	if (isTimingEnabled()) console.log(`[e2e-timing] start ${label}`);
	try {
		return await action();
	} finally {
		if (isTimingEnabled()) {
			const elapsed = ((performance.now() - started) / 1000).toFixed(2);
			console.log(`[e2e-timing] end ${label} (${elapsed}s)`);
		}
	}
}

export function timedTest(name: string, action: () => Promise<void>): void {
	test(name, () => measure(`test ${name}`, action));
}

function isTimingEnabled(): boolean {
	return /^(1|true|yes)$/i.test(process.env.E2E_TIMING || "");
}
