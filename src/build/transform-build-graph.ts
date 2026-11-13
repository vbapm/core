import { asyncMap } from "../utils/async-map";
import { BuildGraph } from "./build-graph";
import * as editorConfig from "./transforms/editor-config";
import * as htmlEntities from "./transforms/html-entities";

const transforms = [htmlEntities, editorConfig];

export const toCompiled = asyncMap<BuildGraph>(
	...transforms.map(transform => transform.toCompiled).filter(Boolean)
);
export const toSrc = asyncMap<BuildGraph>(
	...transforms.map(transform => transform.toSrc).filter(Boolean)
);
