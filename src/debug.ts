import * as debugNamespace from "debug";
import { createRequire } from "node:module";

type DebugFactory = (id: string) => (...messages: any[]) => void;

function resolveDebugFactory(value: unknown): DebugFactory {
	let current = value;
	for (let depth = 0; depth < 3; depth++) {
		if (typeof current === "function") return current as DebugFactory;
		if (!current || typeof current !== "object" || !("default" in current)) break;
		current = (current as { default: unknown }).default;
	}
	throw new TypeError("The debug package did not provide a factory function.");
}

const debugFactory = process.env.VBAPM_STANDALONE
	? resolveDebugFactory(debugNamespace)
	: (createRequire(import.meta.url)("debug") as DebugFactory);

export default debugFactory;
