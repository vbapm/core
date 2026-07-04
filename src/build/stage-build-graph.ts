import { writeFile } from "../utils/fs";
import { parallel } from "../utils/parallel";
import { basename, join } from "../utils/path";
import { BuildGraph, ImportGraph } from "./build-graph";
import { encodeForCodepage, getSystemCodepage } from "./encoding-sniffer";

export async function stageBuildGraph(graph: BuildGraph, staging: string): Promise<ImportGraph> {
	// VBA's Component.Import reads files in the system ANSI codepage,
	// so stage source files in that encoding.
	const codepage = getSystemCodepage();

	const components = await parallel(graph.components, async component => {
		const path = join(staging, component.filename);
		await writeFile(path, encodeForCodepage(component.code, codepage));

		if (component.binaryPath) {
			const binaryPath = join(staging, basename(component.binaryPath));
			if (!component.details.binary) {
				throw new Error(`Binary data missing for component "${component.name}"`);
			}

			await writeFile(binaryPath, component.details.binary);
		}

		return { name: component.name, path };
	});

	return {
		name: graph.name,
		components,
		references: graph.references
	};
}
