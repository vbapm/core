import mri, { Args } from "mri";

export function parseCliArgs(argv: string[]): Args {
	return mri(argv, {
		alias: {
			v: "version",
			h: "help"
		},
		boolean: ["background", "global", "local", "list", "unset"]
	});
}
