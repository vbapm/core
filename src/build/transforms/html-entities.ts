import { BuildGraph } from "../build-graph";
import { decodeHtmlEntities } from "../named-html-entities";

export async function toCompiled(graph: BuildGraph): Promise<BuildGraph> {
	return graph;
}

export async function toSrc(graph: BuildGraph): Promise<BuildGraph> {
	for (const component of graph.components) {
		component.code = component.code.replace(
			/^([ \t]*')([^\r\n]*)/gm,
			(_line, prefix: string, comment: string) => `${prefix}${decodeHtmlEntities(comment)}`
		);
	}

	return graph;
}
