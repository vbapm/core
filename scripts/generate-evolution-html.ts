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
.run-summary { margin: 0 0 24px; border: 1px solid #ccd6de; border-radius: 7px; background: #fff; overflow: hidden; }
.run-summary > summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 10px; padding: 12px 16px; background: #f4f8fb; user-select: none; }
.run-summary > summary::-webkit-details-marker { display: none; }
.run-summary > summary::before { content: '\\25B8 '; color: #3a7ca5; font-size: 13px; }
.run-summary[open] > summary::before { content: '\\25BE '; }
.run-summary > summary h2 { margin: 0; font-size: 18px; color: #17212b; }
.run-summary > summary .summary-hint { margin-left: auto; font-size: 12px; color: #7b8994; }
.run-summary[open] > summary { border-bottom: 1px solid #e3eaf0; }
.run-summary-body { padding: 18px; }
.run-summary .muted { margin: 8px 0; }
.run-summary-stats { display: flex; flex-wrap: wrap; gap: 10px; margin: 4px 0 16px; }
.run-summary-stats .metric strong { color: #17212b; }
.run-summary-chart { width: 100%; height: auto; }
.lane-label { font-size: 11px; fill: #536474; font-family: ui-monospace, SFMono-Regular, Consolas, monospace; }
.legend { display: flex; flex-wrap: wrap; gap: 14px; margin-top: 12px; font-size: 12px; color: #536474; }
.legend-item { display: inline-flex; align-items: center; gap: 6px; }
.swatch { width: 18px; height: 4px; border-radius: 3px; display: inline-block; }
.swatch-dot { width: 9px; height: 9px; border-radius: 50%; display: inline-block; border: 2px solid; }
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
const runSummaryState = {};
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
function displayLog(value) {
	const text = String(value ?? "");
	const match = text.match(/^\\d{4}-\\d{2}-\\d{2}T(\\d{2}:\\d{2}:\\d{2}(?:\\.\\d+)?(?:Z|[+-]\\d{2}:\\d{2}))(\\s+::\\s+.*)?$/);
	return match ? match[1] + (match[2] || "") : text;
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
	const logs = nonEmptyCollection(instance.logs).map(displayLog);
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
function fmtDuration(ms) {
	if (!ms) return '';
	const total = Math.round(ms / 1000);
	const m = Math.floor(total / 60);
	const s = total % 60;
	return m > 0 ? m + 'm ' + s + 's' : s + 's';
}
function runTimeRange(run) {
	let first = 0;
	let last = 0;
	(run.snapshots || []).forEach(snapshot => {
		const t = snapshotTime(snapshot.name);
		if (t) {
			if (!first || t < first) first = t;
			if (t > last) last = t;
		}
	});
	return { first, last };
}
function executionMode(command) {
	if (!command) return '';
	if (/(^|\\s)--runInBand(?:\\s|$)/.test(command)) return 'runInBand';
	const maxWorkers = command.match(/(?:^|\\s)--maxWorkers(?:=|\\s+)([^\\s]+)/);
	return maxWorkers ? 'maxWorkers ' + maxWorkers[1] : '';
}
function renderRunMeta() {
	const el = document.getElementById('run-meta');
	if (!el) return;
	const run = runs[currentRun];
	const meta = (run && run.meta) || {};
	// When run-info.json is missing (runs captured before metadata was recorded),
	// fall back to timing derived from the snapshots so the header still shows
	// the run window and duration.
	const range = runTimeRange(run);
	const startMs = meta.startedAt ? new Date(meta.startedAt).getTime() : range.first;
	const endMs = meta.finishedAt ? new Date(meta.finishedAt).getTime() : range.last;
	let elapsedMs = typeof meta.elapsedMs === 'number' ? meta.elapsedMs : 0;
	if (!elapsedMs && startMs && endMs && endMs > startMs) elapsedMs = endMs - startMs;

	const flags = meta.flags && typeof meta.flags === 'object' ? meta.flags : {};
	const keys = Object.keys(flags).sort();
	const chips = [];
	if (meta.script) chips.push('<span class="run-chip">script: <code>' + escapeHtml(meta.script) + '</code></span>');
	if (elapsedMs) chips.push('<span class="run-chip">duration: <strong>' + escapeHtml(fmtDuration(elapsedMs)) + '</strong></span>');
	if (startMs) chips.push('<span class="run-chip">started: <code>' + escapeHtml(isoLocal(new Date(startMs).toISOString())) + '</code></span>');
	if (endMs) chips.push('<span class="run-chip">finished: <code>' + escapeHtml(isoLocal(new Date(endMs).toISOString())) + '</code></span>');
	if (keys.length) chips.push('<span class="run-chip">flags: <code>' + keys.map(key => escapeHtml(key + '=' + String(flags[key]))).join(' &nbsp;·&nbsp; ') + '</code></span>');
	const mode = executionMode(meta.command);
	if (mode) chips.push('<span class="run-chip">execution: <strong>' + escapeHtml(mode) + '</strong></span>');
	const command = meta.command ? '<div class="run-command">$ ' + escapeHtml(meta.command) + '</div>' : '';
	if (chips.length) {
		const derived = !meta.startedAt && !meta.finishedAt && !meta.script && !command;
		const note = derived
			? '<div class="run-command">Timing derived from evolution snapshots — run configuration was not recorded for this run.</div>'
			: '';
		el.innerHTML = '<div class="run-meta-grid">' + chips.join('') + '</div>' + note + command;
	} else {
		el.innerHTML = '<div class="muted">No run metadata recorded for this evolution run.</div>';
	}
}
function snapshotTime(name) {
	const m = String(name).match(/^instances-(\\d{4})-(\\d{2})-(\\d{2})--(\\d{2})h(\\d{2})m(\\d{2})s(\\d{3})(?:-\\d+)?\\.json$/i);
	if (!m) return 0;
	return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6], +m[7]);
}
function fmtUtc(ms) {
	const d = new Date(ms);
	const p = value => String(value).padStart(2, '0');
	return p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds()) + ' UTC';
}
function fmtTick(ms) {
	const d = new Date(ms);
	const p = value => String(value).padStart(2, '0');
	return p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds());
}
function instanceColor(inst) {
	const warm = inst.intent === 'keep-warm' || String(inst.owner || '').indexOf('session#') === 0;
	if (warm) return '#e09a3c';
	return inst.visible ? '#3a7ca5' : '#5f8fb5';
}
function buildRunSummary(run) {
	const snaps = run.snapshots
		.map(snapshot => ({ time: snapshotTime(snapshot.name), data: snapshot.data || {} }))
		.filter(snapshot => snapshot.time)
		.sort((a, b) => a.time - b.time);
	const byPid = {};
	let peak = 0;
	snaps.forEach(snapshot => {
		const active = Array.isArray(snapshot.data.active) ? snapshot.data.active : [];
		const inactive = Array.isArray(snapshot.data.inactive) ? snapshot.data.inactive : [];
		let alive = 0;
		active.forEach(instance => {
			if (!instance || instance.pid == null) return;
			alive++;
			const key = String(instance.pid);
			if (!byPid[key]) {
				byPid[key] = {
					pid: instance.pid,
					owner: instance.owner,
					intent: instance.intent,
					visible: instance.visible,
					birth: snapshot.time,
					death: null,
					lastActive: snapshot.time
				};
			} else {
				byPid[key].lastActive = snapshot.time;
			}
		});
		inactive.forEach(instance => {
			if (!instance || instance.pid == null) return;
			const key = String(instance.pid);
			if (byPid[key] && byPid[key].death == null) byPid[key].death = snapshot.time;
		});
		peak = Math.max(peak, alive);
	});
	const instances = Object.keys(byPid)
		.map(key => byPid[key])
		.sort((a, b) => (a.birth - b.birth) || (a.pid - b.pid));
	const lastSnap = snaps.length ? snaps[snaps.length - 1].time : 0;
	const fallbackStart = instances.length ? instances[0].birth : 0;
	const meta = (run && run.meta) || {};
	const start = meta.startedAt ? new Date(meta.startedAt).getTime() : fallbackStart;
	const end = meta.finishedAt ? new Date(meta.finishedAt).getTime() : (lastSnap || fallbackStart);
	instances.forEach(instance => {
		if (instance.death == null) {
			instance.death = instance.lastActive >= lastSnap ? end : instance.lastActive;
			if (instance.lastActive >= lastSnap) instance.stillAlive = true;
		}
		if (instance.death < instance.birth) instance.death = instance.birth;
		if (instance.death > end) instance.death = end;
	});
	return { instances, peak, start: start || fallbackStart, end: end || start || fallbackStart };
}
function svgLifecycleChart(instances, start, end) {
	if (!instances.length) return '<div class="muted">No Excel instances were observed in this run.</div>';
	const span = (end - start) || 1;
	const W = 1180;
	const XL = 64;
	const XR = 24;
	const laneStep = 30;
	const top = 16;
	const bottom = 36;
	const H = top + instances.length * laneStep + bottom;
	const X = time => XL + ((W - XL - XR) * (time - start)) / span;
	const parts = [];
	parts.push('<svg class="run-summary-chart" viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg">');
	instances.forEach((instance, index) => {
		const y = top + index * laneStep + laneStep / 2;
		parts.push('<line x1="' + XL + '" y1="' + y + '" x2="' + (W - XR) + '" y2="' + y + '" stroke="#eef2f5" stroke-width="1"/>');
	});
	// Vertical gridline every 10 seconds (labels drawn where they do not collide
	// with the start/end time labels or each other).
	const tickMs = 10000;
	let previousTickX = -1e9;
	for (let tick = Math.ceil(start / tickMs) * tickMs; tick <= end; tick += tickMs) {
		const x = X(tick);
		if (x < XL || x > W - XR) continue;
		parts.push('<line x1="' + x + '" y1="' + top + '" x2="' + x + '" y2="' + (H - bottom) + '" stroke="#d9e3ea" stroke-width="1"/>');
		if (x - XL >= 108 && W - XR - x >= 86 && x - previousTickX >= 52) {
			parts.push('<text x="' + x + '" y="' + (H - 12) + '" text-anchor="middle" class="lane-label">' + fmtTick(tick) + '</text>');
			previousTickX = x;
		}
	}
	parts.push('<text x="' + XL + '" y="' + (H - 12) + '" class="lane-label">' + fmtUtc(start) + '</text>');
	parts.push('<text x="' + (W - XR) + '" y="' + (H - 12) + '" text-anchor="end" class="lane-label">' + fmtUtc(end) + '</text>');
	instances.forEach((instance, index) => {
		const color = instanceColor(instance);
		const y = top + index * laneStep + laneStep / 2;
		const x1 = Math.max(XL, X(instance.birth));
		let x2 = Math.min(W - XR, X(instance.death));
		if (x2 - x1 < 2) x2 = Math.min(W - XR, x1 + 2);
		parts.push('<line x1="' + x1 + '" y1="' + y + '" x2="' + x2 + '" y2="' + y + '" stroke="' + color + '" stroke-width="3" stroke-linecap="round"/>');
		parts.push('<circle cx="' + x1 + '" cy="' + y + '" r="4" fill="#fff" stroke="' + color + '" stroke-width="2"/>');
		parts.push('<circle cx="' + x2 + '" cy="' + y + '" r="4" fill="' + color + '" stroke="#fff" stroke-width="1.5"/>');
		parts.push('<text x="' + (XL - 6) + '" y="' + (y + 3) + '" text-anchor="end" class="lane-label">PID ' + instance.pid + '</text>');
	});
	parts.push('</svg>');
	return parts.join('');
}
function summaryHtml(run) {
	const summary = buildRunSummary(run);
	const open = runSummaryState[currentRun] !== false;
	return '<details id="run-summary-details" class="run-summary"' + (open ? ' open' : '') + '>' +
		'<summary><h2>Run summary</h2><span class="summary-hint">' + (open ? 'click to collapse' : 'click to expand') + '</span></summary>' +
		'<div class="run-summary-body">' +
		'<div class="run-summary-stats">' +
		'<span class="metric">Evolution snapshots: <strong>' + run.snapshots.length + '</strong></span>' +
		'<span class="metric">Excel instances: <strong>' + summary.instances.length + '</strong></span>' +
		'<span class="metric">Peak concurrent instances: <strong>' + summary.peak + '</strong></span>' +
		'</div>' +
		svgLifecycleChart(summary.instances, summary.start, summary.end) +
		'<div class="legend">' +
		'<span class="legend-item"><span class="swatch" style="background:#e09a3c"></span> keep-warm (persistent session)</span>' +
		'<span class="legend-item"><span class="swatch" style="background:#3a7ca5"></span> one-shot automation</span>' +
		'<span class="legend-item"><span class="swatch" style="background:#5f8fb5"></span> other / visible</span>' +
		'<span class="legend-item"><span class="swatch-dot" style="border-color:#fff;background:#3a7ca5"></span> death point</span>' +
		'</div></div></details>';
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
	content.innerHTML = summaryHtml(run) + (instances.length ? instances.join('') : '<div class="error">No Excel instances in this snapshot.</div>');
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
document.addEventListener("toggle", event => {
	const target = event.target;
	if (target && target.id === "run-summary-details") {
		runSummaryState[currentRun] = target.open;
	}
});
// Record the collapsed state synchronously on click (the native "toggle" event is
// queued asynchronously and could race a re-render when navigating snapshots).
document.addEventListener("click", event => {
	const summary = event.target && event.target.closest ? event.target.closest("#run-summary-details summary") : null;
	if (!summary) return;
	const details = document.getElementById("run-summary-details");
	if (!details) return;
	const open = !details.open;
	runSummaryState[currentRun] = open;
	const hint = details.querySelector(".summary-hint");
	if (hint) hint.textContent = open ? "click to collapse" : "click to expand";
});
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
