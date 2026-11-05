import { writeFile } from "../utils/fs";
import { parallel } from "../utils/parallel";
import { basename, join } from "../utils/path";
import * as iconv from "iconv-lite";
import { env } from "../env";
import { createNamedHtmlEntityHandler } from "./named-html-entities";
import { BuildGraph, ImportGraph } from "./build-graph";
import { Codepage, codepageToLabel, getSystemCodepage } from "./encoding-sniffer";

const debug = env.debug("vbapm:build.encoding");

export async function stageBuildGraph(
	graph: BuildGraph,
	staging: string,
	targetCodepage?: Codepage
): Promise<ImportGraph> {
	// VBA's Component.Import reads files in the system ANSI codepage,
	// so stage source files in that encoding — or the target encoding
	// if one was explicitly declared in the manifest.
	const targetCp = targetCodepage ?? getSystemCodepage();
	const targetLabel = codepageToLabel(targetCp);

	const components = await parallel(graph.components, async component => {
		const path = join(staging, component.filename);

		const encoded =
			targetCp === Codepage.Ascii
				? iconv.encode(component.code, targetLabel, {
						invalidCharHandler: createNamedHtmlEntityHandler(
							component.code,
							(original, replacement, index) => {
								debug(
									'🔄️ Replaced "%s" with "%s" in %s at character %d',
									original,
									replacement,
									component.details.path ?? component.filename,
									index + 1
								);
							}
						)
					})
				: iconv.encode(component.code, targetLabel);
		if (!encoded) {
			throw new Error(`Encoding component "${component.name}" was cancelled`);
		}
		const code = Buffer.from(encoded);

		await writeFile(path, code);

		if (component.binaryPath) {
			const binaryPath = join(staging, basename(component.binaryPath));
			if (!component.details.binary) {
				throw new Error(`Binary data missing for component "${component.name}"`);
			}

			await writeFile(binaryPath, component.details.binary);
		}

		return { name: component.name, path, type: component.type };
	});

	return {
		name: graph.name,
		components,
		references: graph.references,
		...(graph.powerQueries ? { powerQueries: graph.powerQueries } : {})
	};
}
