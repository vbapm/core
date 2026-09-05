/**
 * Build a single HTML viewer from the Excel registry evolution snapshots.
 *
 * Usage:
 *   node scripts/generate-evolution-html.ts
 *   node scripts/generate-evolution-html.ts --input <dir> --output <file>
 */
const { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } = require("node:fs");
const { spawnSync } = require("node:child_process");
const { createServer } = require("node:http");
const { dirname, join, resolve } = require("node:path");

function argumentValue(name: string): string | undefined {
	const index = process.argv.indexOf(name);
	return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasArgument(name: string): boolean {
	return process.argv.includes(name);
}

function defaultEvolutionDirectory(): string {
	const temp = process.env.TEMP || process.env.TMP || ".";
	return join(temp, "Excel-Instances", "evolution");
}

function powershellLiteral(value: string): string {
	return `'${value.replace(/'/g, "''")}'`;
}

function readSnapshot(listScript: string, snapshotPath: string): unknown {
	const result = spawnSync(
		"powershell.exe",
		[
			"-NoProfile",
			"-ExecutionPolicy",
			"Bypass",
			"-Command",
			`& ${powershellLiteral(listScript)} -RegistryPath ${powershellLiteral(snapshotPath)} -OutputJson:$true`
		],
		{ encoding: "utf8", env: { ...process.env, VBA_JSON_BASE64: "1" } }
	);

	if (result.error) {
		throw new Error(`Could not run List-ExcelInstances.ps1: ${result.error.message}`);
	}
	if (result.status !== 0) {
		throw new Error(
			`List-ExcelInstances.ps1 failed for ${snapshotPath}: ${result.stderr || result.stdout}`
		);
	}

	try {
		const encoded = result.stdout.trim();
		const json = Buffer.from(encoded, "base64").toString("utf8");
		return JSON.parse(json);
	} catch (error) {
		throw new Error(`Invalid JSON from ${snapshotPath}: ${String(error)}`);
	}
}

function htmlPage(snapshots: Array<{ name: string; data: unknown }>): string {
	const serialized = JSON.stringify(snapshots).replace(/</g, "\\u003c");

	return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel registry evolution</title>
<style>
:root { color-scheme: light; font-family: "Segoe UI", system-ui, sans-serif; background: #edf1f5; color: #17212b; }
* { box-sizing: border-box; }
body { margin: 0; min-height: 100vh; background: #edf1f5; }
header { padding: 28px clamp(20px, 5vw, 72px) 18px; background: #17212b; color: #f5f8fa; }
h1 { margin: 0 0 8px; font-size: clamp(24px, 4vw, 42px); letter-spacing: 0; }
#snapshot-label { color: #b8c6d1; font-size: 14px; }
main { width: min(1180px, calc(100% - 32px)); margin: 24px auto 96px; }
.summary { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 18px; }
.metric { padding: 10px 14px; border: 1px solid #ccd6de; border-radius: 7px; background: #fff; font-size: 14px; }
.instance { margin: 16px 0; padding: 18px; border: 1px solid #ccd6de; border-radius: 7px; background: #fff; box-shadow: 0 3px 12px #17212b12; }
.instance h2 { margin: 0 0 12px; font-size: 18px; }
.instance-meta { color: #536474; font-size: 13px; margin-bottom: 14px; }
.columns { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px; }
section h3 { margin: 0 0 8px; font-size: 14px; color: #304353; }
ul { margin: 0; padding-left: 22px; }
li { margin: 5px 0; overflow-wrap: anywhere; }
.muted { color: #7b8994; }
footer { position: fixed; right: 0; bottom: 0; left: 0; display: flex; align-items: center; justify-content: center; gap: 18px; padding: 12px 16px; border-top: 1px solid #ccd6de; background: #ffffffee; backdrop-filter: blur(8px); }
button { width: 44px; height: 38px; border: 1px solid #9eabb5; border-radius: 6px; background: #fff; color: #17212b; font-size: 23px; line-height: 1; cursor: pointer; }
button:hover:not(:disabled) { background: #e6f0f5; border-color: #31708f; }
button:disabled { cursor: default; opacity: .38; }
#position { min-width: 110px; text-align: center; font-variant-numeric: tabular-nums; color: #536474; font-size: 13px; }
.error { padding: 18px; border-left: 4px solid #bf3e30; background: #fff3f1; }
</style>
</head>
<body>
<header><h1>Excel registry evolution</h1><div id="snapshot-label"></div></header>
<main id="content"></main>
<footer>
<button id="previous" type="button" aria-label="Previous snapshot" title="Previous snapshot">&#8592;</button>
<div id="position"></div>
<button id="next" type="button" aria-label="Next snapshot" title="Next snapshot">&#8594;</button>
</footer>
<script id="snapshot-data" type="application/json">${serialized}</script>
<script>
const snapshots = JSON.parse(document.getElementById("snapshot-data").textContent);
let current = 0;
const content = document.getElementById("content");
const label = document.getElementById("snapshot-label");
const position = document.getElementById("position");
const previous = document.getElementById("previous");
const next = document.getElementById("next");

function escapeHtml(value) {
	return String(value ?? "").replace(/[&<>"']/g, character => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[character]));
}
function list(items, emptyText) {
	const values = Array.isArray(items) ? items : [];
	return values.length ? '<ul>' + values.map(value => '<li>' + escapeHtml(value) + '</li>').join('') + '</ul>' : '<div class="muted">' + emptyText + '</div>';
}
function workbookRows(instance) {
  const readOnly = instance.workbookReadOnly || {};
  const ghosts = new Set((instance.ghostWorkbooks || []).map(ghost => String(ghost.file || "").toLowerCase()));
  return (Array.isArray(instance.workbooks) ? instance.workbooks : []).map(workbook => {
    const value = String(workbook);
    const key = value.toLowerCase();
    const icon = ghosts.has(key) ? "👻" : (readOnly[key] ? "📖🔭" : "📖✏️");
    return icon + " " + escapeHtml(value);
  });
}
function renderInstance(instance) {
  const workbooks = workbookRows(instance);
  const addins = (Array.isArray(instance.addins) ? instance.addins : [])
    .filter(addin => addin.isOpen)
    .map(addin => "📚" + (addin.readOnly ? "🔭" : "✏️") + " " + escapeHtml(addin.fullPath || addin.name));
  const ghostCount = Number(instance.ghostDetection?.ghostCount || 0);
	return '<article class="instance"><h2>' + escapeHtml(instance.number) + ') PID ' + escapeHtml(instance.pid) + '</h2>' +
		'<div class="instance-meta">' + escapeHtml(instance.visible ? "Visible" : "Hidden") + ' · ' +
    escapeHtml(instance.comReachable ? "COM reachable" : "COM unavailable") + " · " +
		escapeHtml(instance.owner || "unknown") + ' · ghost projects: ' + ghostCount + '</div>' +
		'<div class="columns"><section><h3>Workbooks</h3>' + list(workbooks, "None recorded") +
		'</section><section><h3>Open add-ins</h3>' + list(addins, "None recorded") + '</section></div></article>';
}
function render() {
  const snapshot = snapshots[current];
  const data = snapshot.data || {};
  const active = Array.isArray(data.active) ? data.active : [];
  const inactive = Array.isArray(data.inactive) ? data.inactive : [];
  label.textContent = snapshot.name + " · " + active.length + " active · " + inactive.length + " inactive";
  position.textContent = (current + 1) + " / " + snapshots.length;
	content.innerHTML = active.length ? active.map(renderInstance).join("") : '<div class="error">No active Excel instances in this snapshot.</div>';
  previous.disabled = current === 0;
  next.disabled = current === snapshots.length - 1;
}
previous.addEventListener("click", () => { if (current > 0) { current--; render(); } });
next.addEventListener("click", () => { if (current < snapshots.length - 1) { current++; render(); } });
document.addEventListener("keydown", event => { if (event.key === "ArrowLeft") previous.click(); if (event.key === "ArrowRight") next.click(); });
render();
</script>
</body>
</html>
`;
}

function main(): void {
	const input = resolve(argumentValue("--input") || defaultEvolutionDirectory());
	const output = resolve(argumentValue("--output") || join(input, "instances-evolution.html"));
	if (!existsSync(input)) throw new Error(`Evolution directory not found: ${input}`);

	const listScript = join(__dirname, "ps", "List-ExcelInstances.ps1");
	const files = readdirSync(input)
		.filter((file: string) => /^instances-.*\.json$/i.test(file))
		.sort((left: string, right: string) => left.localeCompare(right));
	if (!files.length) throw new Error(`No evolution snapshots found in ${input}`);

	const snapshots = files.map((name: string) => ({
		name,
		data: readSnapshot(listScript, join(input, name))
	}));

	mkdirSync(dirname(output), { recursive: true });
	writeFileSync(output, htmlPage(snapshots), "utf8");
	console.log(`Generated ${output} from ${snapshots.length} snapshot(s).`);

	if (hasArgument("--serve")) {
		const server = createServer((request: any, response: any) => {
			if (request.url !== "/" && request.url !== "/instances-evolution.html") {
				response.writeHead(404);
				response.end("Not found");
				return;
			}
			response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
			response.end(readFileSync(output));
		});
		server.listen(0, "127.0.0.1", () => {
			const address = server.address();
			const port = typeof address === "object" && address ? address.port : 0;
			console.log(`Serving evolution viewer at http://127.0.0.1:${port}/`);
		});
	}
}

main();

type CommonJsEntryPoint = never;
export type { CommonJsEntryPoint };
