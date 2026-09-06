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
const { pathToFileURL } = require("node:url");

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

/** Read the optional run-info.json (flags + start/finish + duration) for a run dir. */
function readRunInfo(dir: string): Record<string, unknown> | null {
	try {
		return JSON.parse(readFileSync(join(dir, "run-info.json"), "utf8"));
	} catch {
		return null;
	}
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

function htmlPage(
	runs: Array<{
		name: string;
		snapshots: Array<{ name: string; data: unknown; sourceUrl: string }>;
		meta?: Record<string, unknown> | null;
	}>
): string {
	const serialized = JSON.stringify(runs).replace(/</g, "\\u003c");

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
#snapshot-label { display: flex; align-items: center; gap: 8px; color: #b8c6d1; font-size: 14px; }
#snapshot-stats { margin-left: 2px; }
.snapshot-link { display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; border: 1px solid #718494; border-radius: 5px; color: #f5f8fa; }
.snapshot-link:hover { background: #304353; border-color: #b8c6d1; }
.snapshot-link:focus-visible { outline: 2px solid #f5f8fa; outline-offset: 2px; }
.snapshot-link svg { width: 15px; height: 15px; }
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
.run-meta { margin: 14px 0 2px; }
.run-meta-grid { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
.run-chip { padding: 4px 10px; border: 1px solid #3e5868; border-radius: 999px; background: #22303c; color: #e9f1f6; font-size: 12.5px; }
.run-chip code { color: #ffd479; }
.run-chip strong { color: #8fe3a0; }
.run-command { flex-basis: 100%; font-size: 12px; color: #9fb3c2; overflow-wrap: anywhere; font-family: ui-monospace, SFMono-Regular, Consolas, monospace; }
.run-meta .muted { color: #7b8994; font-size: 13px; }
footer { position: fixed; right: 0; bottom: 0; left: 0; display: flex; align-items: center; justify-content: center; gap: 18px; padding: 12px 16px; border-top: 1px solid #ccd6de; background: #ffffffee; backdrop-filter: blur(8px); }
button { width: 44px; height: 38px; border: 1px solid #9eabb5; border-radius: 6px; background: #fff; color: #17212b; font-size: 23px; line-height: 1; cursor: pointer; }
button:hover:not(:disabled) { background: #e6f0f5; border-color: #31708f; }
button:disabled { cursor: default; opacity: .38; }
#position { min-width: 110px; text-align: center; font-variant-numeric: tabular-nums; color: #536474; font-size: 13px; }
.error { padding: 18px; border-left: 4px solid #bf3e30; background: #fff3f1; }
</style>
</head>
<body>
<header><h1>Excel registry evolution</h1><div id="run-meta" class="run-meta"></div><label for="run-select">Evolution run</label> <select id="run-select"></select><div id="snapshot-label"><span id="snapshot-name"></span><a id="open-snapshot" class="snapshot-link" target="_blank" rel="noopener" aria-label="Open snapshot JSON file" title="Open snapshot JSON file"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3h7v7"></path><path d="M10 14 21 3"></path><path d="M21 14v6a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h6"></path></svg></a><span id="snapshot-stats"></span></div></header>
<main id="content"></main>
<footer>
<button id="previous" type="button" aria-label="Previous snapshot" title="Previous snapshot">&#8592;</button>
<div id="position"></div>
<button id="next" type="button" aria-label="Next snapshot" title="Next snapshot">&#8594;</button>
</footer>
<script id="snapshot-data" type="application/json">${serialized}</script>
<script>
const runs = JSON.parse(document.getElementById("snapshot-data").textContent);
let currentRun = runs.length - 1;
let current = 0;
const content = document.getElementById("content");
const snapshotName = document.getElementById("snapshot-name");
const openSnapshot = document.getElementById("open-snapshot");
const snapshotStats = document.getElementById("snapshot-stats");
const position = document.getElementById("position");
const runSelect = document.getElementById("run-select");
const previous = document.getElementById("previous");
const next = document.getElementById("next");

function escapeHtml(value) {
	return String(value ?? "").replace(/[&<>"']/g, character => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[character]));
}
function list(items, emptyText) {
	const values = collection(items);
	return values.length ? '<ul>' + values.map(value => '<li>' + escapeHtml(value) + '</li>').join('') + '</ul>' : '<div class="muted">' + emptyText + '</div>';
}
function collection(value) {
	if (Array.isArray(value)) return value;
	if (value == null || (typeof value === "object" && Object.keys(value).length === 0)) return [];
	return [value];
}
function nonEmptyCollection(value) {
	return collection(value).filter(item => String(item ?? "").trim().length > 0);
}
function workbookRows(instance) {
  const readOnly = instance.workbookReadOnly || {};
	const ghostItems = collection(instance.ghostWorkbooks);
	const ghosts = new Set(ghostItems.map(ghost => String(ghost.file || "").toLowerCase()));
	const workbooks = collection(instance.workbooks);
	return workbooks.map(workbook => {
    const value = String(workbook);
    const key = value.toLowerCase();
    const icon = ghosts.has(key) ? "👻" : (readOnly[key] ? "📖🔭" : "📖✏️");
    return icon + " " + escapeHtml(value);
  });
}
function renderInstance(instance, state = "Active") {
	const analyzed = instance.comReachable === true;
  const workbooks = workbookRows(instance);
	const addinItems = collection(instance.addins);
	const addins = addinItems
    .filter(addin => addin.isOpen)
    .map(addin => "📚" + (addin.readOnly ? "🔭" : "✏️") + " " + escapeHtml(addin.fullPath || addin.name));
	const logs = nonEmptyCollection(instance.logs);
  const ghostCount = Number(instance.ghostDetection?.ghostCount || 0);
	return '<article class="instance"><h2>' + escapeHtml(instance.number) + ') PID ' + escapeHtml(instance.pid) + ' · ' + escapeHtml(state) + '</h2>' +
		'<div class="instance-meta">' + escapeHtml(instance.visible ? "Visible" : "Hidden") + ' · ' +
    escapeHtml(instance.comReachable ? "COM reachable" : "COM unavailable") + " · " +
		escapeHtml(instance.owner || "unknown") + (analyzed ? ' · ghost projects: ' + ghostCount : '') + '</div>' +
		(analyzed ? '<div class="columns"><section><h3>Workbooks</h3>' + list(workbooks, "None recorded") +
		'</section><section><h3>Open add-ins</h3>' + list(addins, "None recorded") + '</section></div>' +
		'<section><h3>Logs</h3>' + list(logs, "None recorded") + '</section>' : '') + '</article>';
}
function isoLocal(value) {
	if (!value) return '';
	return String(value).replace('T', ' ').replace('Z', '').slice(0, 19) + ' UTC';
}
function durationText(meta) {
	const ms = meta && typeof meta.elapsedMs === 'number' ? meta.elapsedMs : 0;
	if (!ms) return '';
	const total = Math.round(ms / 1000);
	const m = Math.floor(total / 60);
	const s = total % 60;
	return m > 0 ? m + 'm ' + s + 's' : s + 's';
}
function renderRunMeta() {
	const el = document.getElementById('run-meta');
	if (!el) return;
	const run = runs[currentRun];
	const meta = (run && run.meta) || {};
	const flags = meta.flags && typeof meta.flags === 'object' ? meta.flags : {};
	const keys = Object.keys(flags).sort();
	const chips = [];
	if (meta.script) chips.push('<span class="run-chip">script: <code>' + escapeHtml(meta.script) + '</code></span>');
	const duration = durationText(meta);
	if (duration) chips.push('<span class="run-chip">duration: <strong>' + escapeHtml(duration) + '</strong></span>');
	const started = isoLocal(meta.startedAt);
	const finished = isoLocal(meta.finishedAt);
	if (started) chips.push('<span class="run-chip">started: <code>' + escapeHtml(started) + '</code></span>');
	if (finished) chips.push('<span class="run-chip">finished: <code>' + escapeHtml(finished) + '</code></span>');
	if (keys.length) chips.push('<span class="run-chip">flags: <code>' + keys.map(key => escapeHtml(key + '=' + String(flags[key]))).join(' &nbsp;·&nbsp; ') + '</code></span>');
	const command = meta.command ? '<div class="run-command">$ ' + escapeHtml(meta.command) + '</div>' : '';
	if (!chips.length && !command) {
		el.innerHTML = '<div class="muted">No run metadata recorded for this evolution run.</div>';
		return;
	}
	el.innerHTML = '<div class="run-meta-grid">' + chips.join('') + '</div>' + command;
}
function render() {
	const run = runs[currentRun];
	const snapshots = run.snapshots;
	const snapshot = snapshots[current];
  const data = snapshot.data || {};
  const active = Array.isArray(data.active) ? data.active : [];
  const inactive = Array.isArray(data.inactive) ? data.inactive : [];
	snapshotName.textContent = snapshot.name;
	openSnapshot.href = snapshot.sourceUrl;
	openSnapshot.setAttribute("aria-label", "Open " + snapshot.name);
	openSnapshot.title = "Open " + snapshot.name;
	snapshotStats.textContent = " · " + active.length + " active · " + inactive.length + " inactive";
	position.textContent = (current + 1) + " / " + snapshots.length;
	const inactiveInstances = Array.isArray(data.inactive) ? data.inactive : [];
	const instances = active.map(instance => renderInstance(instance, "Active"))
		.concat(inactiveInstances.map(instance => renderInstance(instance, "Inactive")));
	content.innerHTML = instances.length ? instances.join("") : '<div class="error">No Excel instances in this snapshot.</div>';
  previous.disabled = current === 0;
  next.disabled = current === snapshots.length - 1;
  renderRunMeta();
}
runSelect.innerHTML = runs.map((run, index) => '<option value="' + index + '">' + escapeHtml(run.name) + ' (' + run.snapshots.length + ' snapshots)</option>').join('');
runSelect.value = String(currentRun);
runSelect.addEventListener("change", event => { currentRun = Number(event.target.value); current = 0; render(); });
previous.addEventListener("click", () => { if (current > 0) { current--; render(); } });
next.addEventListener("click", () => { if (current < runs[currentRun].snapshots.length - 1) { current++; render(); } });
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
	const entries = readdirSync(input, { withFileTypes: true });
	const runs = entries
		.filter((entry: any) => entry.isDirectory() && /^\d{4}-\d{2}-\d{2}--/.test(entry.name))
		.map((entry: any) => ({
			name: entry.name,
			meta: readRunInfo(join(input, entry.name)),
			snapshots: readdirSync(join(input, entry.name))
				.filter(
					(file: string) =>
						/^instances-.*\.json$/i.test(file) && existsSync(join(input, entry.name, file))
				)
				.sort((left: string, right: string) => left.localeCompare(right))
				.map((name: string) => ({
					name,
					data: readSnapshot(listScript, join(input, entry.name, name)),
					sourceUrl: pathToFileURL(join(input, entry.name, name)).href
				}))
		}))
		.filter((run: { snapshots: unknown[] }) => run.snapshots.length)
		.sort((left: { name: string }, right: { name: string }) => left.name.localeCompare(right.name));

	const legacyFiles = entries
		.filter(
			(entry: any) =>
				entry.isFile() &&
				/^instances-.*\.json$/i.test(entry.name) &&
				existsSync(join(input, entry.name))
		)
		.map((entry: any) => entry.name)
		.sort((left: string, right: string) => left.localeCompare(right));
	if (legacyFiles.length) {
		runs.unshift({
			name: "Legacy snapshots",
			meta: readRunInfo(input),
			snapshots: legacyFiles.map((name: string) => ({
				name,
				data: readSnapshot(listScript, join(input, name)),
				sourceUrl: pathToFileURL(join(input, name)).href
			}))
		});
	}
	if (!runs.length) throw new Error(`No evolution snapshots found in ${input}`);

	mkdirSync(dirname(output), { recursive: true });
	writeFileSync(output, htmlPage(runs), "utf8");
	console.log(`Generated ${output} from ${runs.length} evolution run(s).`);

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
