import builtin from "builtin-modules";
import commonjs from "@rollup/plugin-commonjs";
import json from "@rollup/plugin-json";
import resolve from "@rollup/plugin-node-resolve";
import replace from "@rollup/plugin-replace";
import terser from "@rollup/plugin-terser";
import typescript from "@rollup/plugin-typescript";
import fs from "fs";
import path from "path";

const mode = process.env.NODE_ENV || "production";
const builtins = new Set(builtin);

// Inline text template files (e.g. .gitignore, .gitattributes, .editorconfig) as string exports.
function inlineTemplates() {
	return {
		name: "inline-templates",
		transform(code, id) {
			if (/\.(gitignore|gitattributes|editorconfig)$/.test(id)) {
				return {
					code: `export default ${JSON.stringify(code)};`,
					map: { mappings: "" }
				};
			}
			return null;
		}
	};
}

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
					} catch (e) {
						// Ignore chmod errors on Windows
					}
				}
			}

			// editorconfig's one-ini parser may load this wasm file at runtime.
			const wasmSource = path.resolve("node_modules", "@one-ini", "wasm", "one_ini_bg.wasm");
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
			inlineTemplates(),
			typescript(),
			mode === "production" && terser(),
			debug(),
			workerThreads(),
			shebang()
		].filter(Boolean),
		onwarn(warning, warn) {
			// Ignore known errors
			if (warning.code === "CIRCULAR_DEPENDENCY" && /glob/.test(warning.importer)) return;
			if (warning.code === "CIRCULAR_DEPENDENCY" && /readable\-stream/.test(warning.importer || "")) return;
			// semver's Range <-> Comparator mutual dependency is an internal cycle
			// that cannot be avoided from outside the package; safe to ignore.			
			if (warning.code === "CIRCULAR_DEPENDENCY" && warning.ids?.some(id => /semver/.test(id))) return;
			if (warning.code === "UNRESOLVED_IMPORT" && /^node:/.test(warning.source || "")) return;
			if (warning.code === "EVAL" && /minisat/.test(warning.id)) return;

			warn(warning);
		}
	}
];

// Deprecated
// Explicitly export modern API from readable-stream
// (exclude fallback API)
function readableStream() {
	const isReadable = /readable\-stream[\\,\/]readable\.js/;
	const isPassthrough = /readable\-stream[\\,\/]passthrough\.js/;
	const isDuplex = /readable\-stream[\\,\/]duplex\.js/;

	return {
		name: "readable-stream",
		load(id) {
			if (isReadable.test(id)) {
				return {
					code: `
            const Stream = require('stream');

            exports = module.exports = Stream.Readable;
            exports.Readable = Stream.Readable;
            exports.Writable = Stream.Writable;
            exports.Duplex = Stream.Duplex;
            exports.Transform = Stream.Transform;
            exports.PassThrough = Stream.PassThrough;
            exports.Stream = Stream;
          `
				};
			}
			if (isPassthrough.test(id)) {
				return {
					code: `module.exports = require('stream').PassThrough;`
				};
			}
			if (isDuplex.test(id)) {
				return {
					code: `module.exports = require('stream').Duplex;`
				};
			}

			return null;
		}
	};
}

function debug() {
	const isBrowser = /debug[\\,\/]src[\\,\/]browser\.js/;

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
