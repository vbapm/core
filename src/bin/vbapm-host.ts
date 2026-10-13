import { Args } from "mri";
import { runHostCommand, HostCloseScope } from "../actions/host";
import { CliError, ErrorCode } from "../errors";

const help = [
	"Inspect or close running Excel instances on Windows.",
	"",
	"Usage: vbapm host [options]",
	"",
	"Options:",
	"  --close-all       Close every running Excel instance",
	"  --close-visible   Close every visible Excel instance",
	"  --close-invisible Close every invisible Excel instance",
	"  --force           Close visible instances even when unsaved workbooks are found"
].join("\n");

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const scopes: Array<[string, HostCloseScope]> = [
		["close-all", "all"],
		["close-visible", "visible"],
		["close-invisible", "invisible"]
	];
	const selected = scopes.filter(([name]) => !!args[name]);

	if (selected.length > 1) {
		throw new CliError(
			ErrorCode.HostOptionsConflict,
			"The host close options are mutually exclusive."
		);
	}

	const output = await runHostCommand(selected[0]?.[1], !!args.force);
	if (output) console.log(output);
}
