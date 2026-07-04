import builtin from "builtin-modules";
import commonjs from "@rollup/plugin-commonjs";
import json from "@rollup/plugin-json";
import resolve from "@rollup/plugin-node-resolve";
import replace from "@rollup/plugin-replace";
import terser from "@rollup/plugin-terser";
import typescript from "@rollup/plugin-typescript";
import fs from "fs";
import { createRequire } from "module";
import path from "path";

const mode = process.env.NODE_ENV || "production";
const builtins = new Set(builtin);

// Add shebang to CLI entry point and make it executable.
// Needed for npm's "bin" field to work. The standalone build
// ignores the shebang because it invokes lib/vbapm.js explicitly
// via the vendored node binary.
function shebang() {
	return {
		name: "shebang",
		renderChunk(code, chunk) {
			if (chunk.facadeModuleId && chunk.facadeModuleId.includes("vbapm.ts")) {
				return { code: "#!/usr/bin/env node\n" + code, map: null };
			}
			return null;
		},
		writeBundle(options, bundle) {
			for (const [fileName] of Object.entries(bundle)) {
				if (fileName === "vbapm.js") {
					const filePath = path.resolve(options.dir, fileName);
					try {
						fs.chmodSync(filePath, 0o755);
					} catch {
						// Ignore chmod errors on Windows
					}
				}
			}

			const templates = [
				"template.editorconfig",
				"template.gitattributes",
				"template.gitignore"
			];
			const templatesSourceDir = path.resolve("src", "actions", "templates");
			const templatesTargetDir = path.resolve(options.dir, "templates");
			fs.mkdirSync(templatesTargetDir, { recursive: true });
			for (const templateFile of templates) {
				fs.copyFileSync(path.join(templatesSourceDir, templateFile), path.join(templatesTargetDir, templateFile));
			}

			// editorconfig's one-ini parser may load this wasm file at runtime.
			// Uses require.resolve to find it even in pnpm's .pnpm/ directory layout.
			const localRequire = createRequire(import.meta.url);
			const wasmPkg = localRequire.resolve("@one-ini/wasm/package.json");
			const wasmSource = path.resolve(path.dirname(wasmPkg), "one_ini_bg.wasm");
			if (fs.existsSync(wasmSource)) {
				const wasmTarget = path.resolve(options.dir, "one_ini_bg.wasm");
				fs.copyFileSync(wasmSource, wasmTarget);
			}
		}
	};
}

export default [
	{
		input: ["src/index.ts", "src/bin/vbapm.ts", "src/debug.ts"],
		output: {
			format: "cjs",
			dir: "lib",
			sourcemap: false,
			exports: "auto"
		},
		external(id) {
			return builtins.has(id) || id.startsWith("node:");
		},
		plugins: [
			resolve(),
			replace({
				preventAssignment: true,
				"process.env.NODE_ENV": JSON.stringify(mode),
				"process.env.READABLE_STREAM": '"disable"',
				"require.cache": "{}"
			}),
			commonjs({
				include: "node_modules/**"
			}),
			json(),
			typescript(),
			mode === "production" && terser(),
			debug(),
			workerThreads(),
			shebang()
		].filter(Boolean),
		onwarn(warning, warn) {
			// Ignore known errors
			if (warning.code === "CIRCULAR_DEPENDENCY" && /glob/.test(warning.importer)) return;
			if (warning.code === "CIRCULAR_DEPENDENCY" && /readable-stream/.test(warning.importer || "")) return;
			// semver's Range <-> Comparator mutual dependency is an internal cycle
			// that cannot be avoided from outside the package; safe to ignore.			
			if (warning.code === "CIRCULAR_DEPENDENCY" && warning.ids?.some(id => /semver/.test(id))) return;
			if (warning.code === "UNRESOLVED_IMPORT" && (warning.source || "").startsWith("node:")) return;
			if (warning.code === "EVAL" && /minisat/.test(warning.id)) return;

			warn(warning);
		}
	}
];

function debug() {
	const isBrowser = /debug[/,\\]src[/,\\]browser\.js/;

	return {
		name: "debug",
		load(id) {
			if (isBrowser.test(id)) {
				return {
					code: `module.exports = {};`
				};
			}
		}
	};
}

function workerThreads() {
	const isWorkerThreads = /worker_threads/;

	return {
		name: "worker_threads",
		resolveId(importee) {
			if (isWorkerThreads.test(importee)) {
				return importee;
			}
		},
		load(id) {
			if (isWorkerThreads.test(id)) {
				return {
					code: `export const threadId = 0;`
				};
			}
		}
	};
}
