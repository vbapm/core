/**
 * Shared Jest snapshot serializer for `readdir()` file-map results.
 *
 * `readdir()` returns a `{ [path]: string }` map of a project's files. This
 * serializer formats those maps as a readable `Object { "path": "contents", ... }`
 * block instead of a raw JSON dump, so snapshot diffs stay reviewable.
 *
 * Import and call `addFileMapSnapshotSerializer()` in any e2e file whose tests
 * assert `readdir()` results with `toMatchSnapshot()`.
 */

export function addFileMapSnapshotSerializer(): void {
	expect.addSnapshotSerializer({
		test: value => isSnapshotFileMap(value),
		print: value => formatSnapshotFileMap(value as { [path: string]: string })
	});
}

function isSnapshotFileMap(value: any): value is { [path: string]: string } {
	if (!value || typeof value !== "object" || Array.isArray(value)) {
		return false;
	}

	const entries = Object.entries(value);
	if (!entries.length) {
		return false;
	}

	return (
		entries.every(([_, contents]) => typeof contents === "string") &&
		entries.some(([path]) => path.includes("/") || path.endsWith(".toml"))
	);
}

function formatSnapshotFileMap(value: { [path: string]: string }): string {
	const lines = ["Object {"];

	for (const [path, contents] of Object.entries(value)) {
		if (contents.includes("\n")) {
			lines.push(`  ${quote(path)}:`);
			lines.push(`  ${quote(contents)},`);
		} else {
			lines.push(`  ${quote(path)}: ${quote(contents)},`);
		}
	}

	lines.push("}");
	return lines.join("\n");
}

function quote(value: string): string {
	return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}
