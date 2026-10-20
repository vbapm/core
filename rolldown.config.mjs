import builtin from "builtin-modules";
import fs from "fs";
import { createRequire } from "module";
import path from "path";
import { replacePlugin } from "rolldown/plugins";

const mode = process.env.NODE_ENV || "production";
const standalone = process.env.VBAPM_STANDALONE === "1";
const selfContained = standalone || mode === "development";
const builtins = new Set(builtin);
const nodeGlobalsBanner = [
	'import { dirname as __dirnameFn } from "node:path";',
	'import { fileURLToPath as __fileURLToPath } from "node:url";',
	"const __filename = __fileURLToPath(import.meta.url);",
	"const __dirname = __dirnameFn(__filename);"
].join("\n");

export default [
	{
		input: ["src/index.ts", "src/bin/vbapm.ts", "src/debug.ts"],
		output: {
			format: "es",
			dir: "lib",
			sourcemap: false,
			banner: selfContained ? nodeGlobalsBanner : undefined,
			minify: mode === "production"
		},
		platform: "node",
		external(id) {
			return (
				builtins.has(id) ||
				id.startsWith("node:") ||
				(!selfContained && ["conf", "editorconfig", "debug"].includes(id))
			);
		},
		transform: {
			define: {
				"process.env.NODE_ENV": JSON.stringify(mode),
				"process.env.VBAPM_STANDALONE": JSON.stringify(selfContained),
				"process.env.READABLE_STREAM": '"disable"',
				"require.cache": "{}"
			},
			typescript: {
				onlyRemoveTypeImports: false
			}
		},
		plugins: [
			replacePlugin({
				"process.env.NODE_ENV": JSON.stringify(mode),
				"process.env.VBAPM_STANDALONE": JSON.stringify(selfContained),
				"process.env.READABLE_STREAM": '"disable"',
				"require.cache": "{}"
			}),
			debug(),
			workerThreads(),
			openPackageGlobals(),
			shebang()
		].filter(Boolean),
		onwarn(warning, warn) {
			if (warning.code === "CIRCULAR_DEPENDENCY" && /glob/.test(warning.importer)) return;
			if (warning.code === "CIRCULAR_DEPENDENCY" && /readable-stream/.test(warning.importer || ""))
				return;
			if (warning.code === "CIRCULAR_DEPENDENCY" && warning.ids?.some(id => /semver/.test(id)))
				return;
			if (warning.code === "UNRESOLVED_IMPORT" && (warning.source || "").startsWith("node:"))
				return;
			if (warning.code === "EVAL" && /minisat/.test(warning.id)) return;

			warn(warning);
		}
	}
];

function openPackageGlobals() {
	const openEntry = /[\\/]open@[^\\/]+[\\/]node_modules[\\/]open[\\/]index\.js$/;

	return {
		name: "open-package-globals",
		transform(code, id) {
			if (!openEntry.test(id)) return null;
			return { code: code.replace(/\b__dirname\b/g, "__openDirname"), map: null };
		}
	};
}

function debug() {
	const isBrowser = /debug[/,\\]src[/,\\]browser\.js/;

	return {
		name: "debug",
		load(id) {
			if (isBrowser.test(id)) {
				return {
					code: `export default {};`
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

			const templates = ["template.editorconfig", "template.gitattributes", "template.gitignore"];
			const templatesSourceDir = path.resolve("src", "actions", "templates");
			const templatesTargetDir = path.resolve(options.dir, "templates");
			fs.mkdirSync(templatesTargetDir, { recursive: true });
			for (const templateFile of templates) {
				fs.copyFileSync(
					path.join(templatesSourceDir, templateFile),
					path.join(templatesTargetDir, templateFile)
				);
			}

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