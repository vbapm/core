import { env } from "../env";
import { Message } from "../messages";
import { resolveSrcSubfolders, writeManifest } from "../manifest";
import { Reference } from "../manifest/reference";
import { Source } from "../manifest/source";
import { Project } from "../project";
import { ensureDir, remove, writeFile } from "../utils/fs";
import { parallel } from "../utils/parallel";
import { dirname, join } from "../utils/path";
import { Changeset } from "./changeset";
import { byComponentTypeThenName, Component, ComponentType, extensionToType } from "./component";
import { codepageToLabel, getSystemCodepage } from "./encoding-sniffer";

export async function applyChangeset(project: Project, changeset: Changeset) {
	const progress = env.reporter.progress("Updating src files");
	const start = progress.start;
	const done = progress.done;

	progress.start = () => {};
	progress.done = () => {};

	start();

	// Update src directory
	await parallel(
		changeset.components.changed,
		component => writeComponent(component.details.path!, component),
		{ progress }
	);

	await parallel(
		changeset.components.added,
		async component => {
			const sub = resolveSrcSubfolders(project.manifest.srcSubfolders, component.type);
			const path = join(project.paths.dir, "src", sub, component.filename);
			component.details.path = path;

			await writeComponent(path, component);
		},
		{ progress }
	);

	await parallel(
		changeset.components.removed,
		async component => {
			await remove(component.details.path!);
		},
		{ progress }
	);

	await updateManifest(project, changeset);

	done();
}

async function updateManifest(project: Project, changeset: Changeset) {
	const structure = project.manifest.srcStructure;
	const enforcement = project.manifest.srcProperties;

	// Grouped convention: globs cover all files, no manifest changes needed.
	if (structure?.grouped) {
		// Still handle removals (source entries with literal paths need cleanup).
		// But for pure grouped mode with globs, just skip.
		for (const component of changeset.components.removed) {
			const index = project.manifest.src.findIndex(
				(source: Source) => source.name === component.name
			);
			if (index >= 0) project.manifest.src.splice(index, 1);
		}
	} else {
		// Individual convention
		for (const component of changeset.components.added) {
			const sub = resolveSrcSubfolders(project.manifest.srcSubfolders, component.type);
			const srcPath = sub ? `src/${sub}/${component.filename}` : `src/${component.filename}`;
			const source: Source = {
				name: component.name,
				path: join(project.paths.dir, srcPath)
			};

			if (enforcement?.sort?.["by-types"] || enforcement?.sort?.alphabetical) {
				// Enforcement active: insert at sorted position
				const src = project.manifest.src;
				let insertAt = src.length;
				for (let i = 0; i < src.length; i++) {
					const existingExt = src[i].path.split(".").pop()?.toLowerCase() || "";
					const newExt = component.filename.split(".").pop()?.toLowerCase() || "";
					// Compare by type then name
					const existingType = extensionToType[`.${existingExt}`] || "class";
					const newType = extensionToType[`.${newExt}`] || "class";
					const cmp = compareByTypeThenName(newType, component.name, existingType, src[i].name);
					if (cmp < 0) {
						insertAt = i;
						break;
					}
				}
				src.splice(insertAt, 0, source);
			} else {
				// No enforcement: append
				project.manifest.src.push(source);
			}
		}

		for (const component of changeset.components.removed) {
			const index = project.manifest.src.findIndex(
				(source: Source) => source.name === component.name
			);
			if (index >= 0) project.manifest.src.splice(index, 1);
		}
	}

	for (let reference of changeset.references.added) {
		project.manifest.references.push(reference);
	}
	for (const reference of changeset.references.removed) {
		const index = project.manifest.references.findIndex(
			(ref: Reference) => ref.name === reference.name
		);
		if (index >= 0) project.manifest.references.splice(index, 1);
	}

	await writeManifest(project.manifest, project.paths.dir);
}

/** Type order for insertion sorting (Objects → Modules → Forms → Classes). */
const TYPE_ORDER_INSERT: Record<string, number> = {
	document: 1,
	module: 2,
	form: 3,
	class: 4
};

function compareByTypeThenName(
	typeA: string, nameA: string,
	typeB: string, nameB: string
): number {
	const orderA = TYPE_ORDER_INSERT[typeA] ?? 99;
	const orderB = TYPE_ORDER_INSERT[typeB] ?? 99;
	if (orderA !== orderB) return orderA - orderB;
	return nameA.toLowerCase() < nameB.toLowerCase() ? -1 : nameA.toLowerCase() > nameB.toLowerCase() ? 1 : 0;
}

async function writeComponent(path: string, component: Component) {
	const dir = dirname(path);
	await ensureDir(dir);

	// Write in the source encoding if declared, otherwise UTF-8.
	// This preserves the original encoding when exporting on a machine
	// whose system codepage differs from the source encoding.
	let data: string | Buffer = component.code;
	const srcEncoding = component.details.sourceEncoding;

	if (srcEncoding) {
		const iconv = require("iconv-lite");
		const systemLabel = codepageToLabel(getSystemCodepage());

		if (srcEncoding.toLowerCase() !== systemLabel.toLowerCase()) {
			// Source encoding differs from the system codepage — transcode.
			// Warn if the target encoding is not UTF-8/UTF-16 and the
			// content has non-ASCII characters that may not be representable.
			const isLossyTarget = !/^utf-?8$/i.test(srcEncoding) && !/^utf-?16/i.test(srcEncoding);

			if (isLossyTarget) {
				const hasNonAscii = [...component.code].some(c => c.charCodeAt(0) > 127);
				if (hasNonAscii) {
					env.reporter.log(
						Message.EncodingLossWarning,
						`Character loss possible: "${component.filename}" contains ` +
							`non-ASCII characters and is being encoded to ${srcEncoding}. ` +
							`Some characters may not be representable in the target encoding.`
					);
					// NOTE: We are working on an iconv-lite feature to allow a
					// handler to be passed when an invalid character is being
					// encoded, so that we can fail the export in the future when
					// there are such incompatibilities.
					// See https://github.com/pillarjs/iconv-lite/pull/388
				}
			}

			data = iconv.encode(component.code, srcEncoding);
		} else {
			// Same encoding as system codepage — encode directly
			data = iconv.encode(component.code, srcEncoding);
		}
	}

	await writeFile(path, data);

	if (component.binaryPath) {
		if (!component.details.binary) {
			throw new Error(`Binary data missing for component "${component.name}"`);
		}

		await writeFile(join(dir, component.binaryPath), component.details.binary);
	}
}
