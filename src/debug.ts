import * as debugNamespace from "debug";
import { createRequire } from "node:module";

type DebugFactory = (id: string) => (...messages: any[]) => void;
type DebugBackend = DebugFactory & {
	enable: (namespaces: string) => void;
};

function resolveDebugBackend(value: unknown): DebugBackend {
	let current = value;
	for (let depth = 0; depth < 3; depth++) {
		if (typeof current === "function") {
			const backend = current as DebugBackend;
			if (typeof backend.enable === "function") return backend;
			break;
		}
		if (!current || typeof current !== "object" || !("default" in current)) break;
		current = (current as { default: unknown }).default;
	}
	throw new TypeError("The debug package did not provide a compatible backend.");
}

const debugBackend = process.env.VBAPM_STANDALONE
	? resolveDebugBackend(debugNamespace)
	: resolveDebugBackend(createRequire(import.meta.url)("debug"));

export const enableDebug = (namespaces: string): void => debugBackend.enable(namespaces);

export default debugBackend;
