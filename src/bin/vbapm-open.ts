import { Args } from "mri";
import { checkTargetOpen, getTargetPath, openTarget } from "../actions/open-target";

const help = [
	"Open the current built target file or check whether it is open.",
	"",
	"Usage: vbapm open [options]",
	"",
	"Options:",
	"  --target=TYPE   Use built target of type TYPE",
	"  --check         Print whether the target is open in visible Excel"
].join("\n");

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;

	const path = await getTargetPath(target);
	if (args.check) {
		console.log(await checkTargetOpen(path));
		return;
	}

	console.log(`Opening built target: ${path}`);
	await openTarget(path);
}
