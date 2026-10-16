import { join, relative, extname } from "node:path";
import { fileURLToPath } from "node:url";
import fsExtra from "fs-extra";
import ls from "./lib/ls.js";

const { ensureDir } = fsExtra;

const root = fileURLToPath(new URL("../", import.meta.url));
const dist = join(root, "dist");

main().catch(err => {
	console.error(err);
	process.exit(1);
});

async function main() {
	const { default: zip } = await import("./lib/zip.mjs");
	await ensureDir(dist);
	await windows(zip);
	await mac(zip);
}

async function windows(zip) {
	const file = join(dist, `vbapm-win.zip`);
	const input = getInput("win32");

	await zip(input, file);
}

async function mac(zip) {
	const file = join(dist, `vbapm-mac.tar.gz`);
	const input = getInput("darwin");

	await zip(input, file, "tar", { gzip: true });
}

function getInput(platform) {
	const ignoreBackup = path => !/\.backup/.test(path);
	const isPowershellScript = path => path.endsWith(".ps1");
	const isAppleScript = path => /\.applescript/.test(path);
	const compatibleRunScript = path =>
		platform === "win32" ? isPowershellScript(path) : isAppleScript(path);
	const isCmd = path => /\.cmd/.test(path);
	const isPowershell = path => /\.ps1/.test(path);
	const isShell = path => !isCmd(path) && !isPowershell(path);
	const compatibleBin = path =>
		platform === "win32" ? isCmd(path) || isPowershell(path) : isShell(path);
	const compatibleExe = path =>
		platform === "win32" ? extname(path) === ".exe" : extname(path) === "";

	const lib = ls(join(root, "lib"));
	const addins = ls(join(root, "addins/build")).filter(ignoreBackup);
	const run_scripts = ls(join(root, "run-scripts")).filter(compatibleRunScript);
	const bin = ls(join(root, "bin")).filter(compatibleBin);
	const vendor = ls(join(root, "vendor")).filter(compatibleExe);

	const input = {};
	for (const file of addins.concat(run_scripts).concat(lib).concat(bin).concat(vendor)) {
		input[file] = relative(root, file);
	}

	return input;
}
