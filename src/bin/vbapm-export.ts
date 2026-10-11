import dedent from "@timhall/dedent";
import { yellowBright } from "@timhall/ansi-colors";
import { Args } from "mri";
import { exportProject } from "../actions/export-project";
import { resolveBackgroundMode } from "../config";

const help = dedent`
  (deprecated) Use "vbapm extract" instead.

  Export built project, including src, references, and target.

  Usage: vbapm export

  Options:
    --target=TYPE   Export target of type TYPE
    --xml-only      Only extract the target XML, skip VBA source export
	--vba-only      Only export the VBA source, skip target XML extraction
	--background    Use a hidden Excel instance

  Debugging options:
    --skip-sheet-name-normalization   Skip sheet name normalization (keep sheetN.xml names)`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	console.warn(
		`\n${yellowBright("WARN:")} "vbapm export" is deprecated. Use "vbapm extract" instead.\n`
	);

	const target = <string | undefined>args.target;
	const completed = <string | undefined>args.completed;
	const addin = <string | undefined>args.addin;
	const xmlOnly = !!args["xml-only"];
	const vbaOnly = !!args["vba-only"];
	const skipSheetNameNormalization = !!args["skip-sheet-name-normalization"];
	const background =
		typeof args.background === "boolean" ? args.background : await resolveBackgroundMode();

	await exportProject({
		target,
		completed,
		addin,
		xmlOnly,
		vbaOnly,
		skipSheetNameNormalization,
		background
	});
}
