import { createRequire } from "node:module";

const debugFactory = createRequire(import.meta.url)("debug") as (
	id: string
) => (...messages: any[]) => void;

export default debugFactory;
