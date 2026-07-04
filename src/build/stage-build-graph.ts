import { writeFile } from "../utils/fs";
import { parallel } from "../utils/parallel";
import { basename, join } from "../utils/path";
import { BuildGraph, ImportGraph } from "./build-graph";
import {
	Codepage,
	codepageToLabel,
	encodeForCodepage,
	getSystemCodepage
} from "./encoding-sniffer";

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

		// Determine source encoding (from TOML declaration or system codepage)
		const sourceLabel = component.details.sourceEncoding ?? targetLabel;

		// Encode component code to the target encoding.
		// If source differs from target, iconv transcodes automatically.
		let code: string | Buffer = component.code;
		if (sourceLabel.toLowerCase() !== targetLabel.toLowerCase()) {
			const iconv = require("iconv-lite");
			code = iconv.encode(component.code, targetLabel);
		} else {
			code = encodeForCodepage(component.code, targetCp);
		}

		await writeFile(path, code);

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
