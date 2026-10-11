const { mkdtempSync, readFileSync, rmSync, writeFileSync } = require("node:fs");
const { tmpdir } = require("node:os");
const { join } = require("node:path");
const {
	checkFreshness,
	computeFingerprint,
	createFreshnessRecord,
	getFingerprintInputs,
	writeFreshness
} = require("../freshness.ts");

describe("freshness metadata", () => {
	let directory: string;

	beforeEach(() => {
		directory = mkdtempSync(join(tmpdir(), "vbapm-freshness-"));
	});

	afterEach(() => {
		rmSync(directory, { recursive: true, force: true });
	});

	function fixture(name: string, contents: string): string {
		const path = join(directory, name);
		writeFileSync(path, contents);
		return path;
	}

	test("sorts inputs and excludes generated output folders", () => {
		const record = createFreshnessRecord();
		const sorted = [...record.inputs].sort();

		expect(record.inputs).toEqual(sorted);
		expect(record.inputs.some((path: string) => path.startsWith("addins/build/"))).toBe(false);
		expect(record.inputs.some((path: string) => path.startsWith("scripts/bootstrap/build/"))).toBe(
			false
		);
		expect(getFingerprintInputs().length).toBe(record.inputs.length);
	});

	test("uses only add-in files for the add-in fingerprint", () => {
		const inputs = createFreshnessRecord("addin").inputs;

		expect(inputs.length).toBeGreaterThan(0);
		expect(inputs.every((path: string) => path.startsWith("addins/"))).toBe(true);
		expect(inputs.some((path: string) => path.startsWith("addins/build/"))).toBe(false);
	});

	test("validates add-in sidecars with the add-in fingerprint", () => {
		const output = fixture("addin.xlam", "built");
		const metadata = join(directory, "addin.xlam.fresh.json");
		writeFreshness(metadata, "addin");

		expect(checkFreshness(output, metadata, "addin").fresh).toBe(true);
		expect(checkFreshness(output, metadata).fresh).toBe(false);
	});

	test("is deterministic regardless of input order", () => {
		const first = fixture("first.txt", "first");
		const second = fixture("second.txt", "second");

		expect(computeFingerprint([first, second])).toBe(computeFingerprint([second, first]));
	});

	test("changes when an input changes", () => {
		const input = fixture("input.txt", "before");
		const before = computeFingerprint([input]);

		writeFileSync(input, "after");

		expect(computeFingerprint([input])).not.toBe(before);
	});

	test("changes when an input is missing", () => {
		const existing = fixture("existing.txt", "content");
		const missing = join(directory, "missing.txt");
		const before = computeFingerprint([existing, missing]);

		writeFileSync(missing, "content");

		expect(computeFingerprint([existing, missing])).not.toBe(before);
	});

	test("validates and rejects stale sidecars", () => {
		const output = fixture("output.bin", "built");
		const metadata = join(directory, "output.fresh.json");
		writeFreshness(metadata);

		expect(checkFreshness(output, metadata).fresh).toBe(true);

		const record = JSON.parse(readFileSync(metadata, "utf8"));
		record.fingerprint = "sha256-stale";
		writeFileSync(metadata, JSON.stringify(record));

		const result = checkFreshness(output, metadata);
		expect(result.fresh).toBe(false);
		expect(result.reason).toContain("source fingerprint changed");
	});

	test("rejects missing outputs and sidecars", () => {
		const output = join(directory, "output.bin");
		const metadata = join(directory, "output.fresh.json");

		expect(checkFreshness(output, metadata).fresh).toBe(false);

		writeFileSync(output, "built");
		expect(checkFreshness(output, metadata).fresh).toBe(false);
	});
});

export {};
