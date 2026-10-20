import dedent from "@timhall/dedent";
import { Args } from "mri";
import { exportProject } from "../actions/export-project";
import { resolveBackgroundMode } from "../config";

const help = dedent`
  Extract built project, including src, references, and target.

  Usage: vbapm extract

  Options:
    --target=TYPE   Extract target of type TYPE
    --xml-only      Only extract the target XML, skip VBA source extraction
	--vba-only      Only extract the VBA source, skip target XML extraction
	--background    Use a hidden Excel instance

  Additional options:
	--skip-sheet-name-normalization   Skip sheet name normalization (keep sheetN.xml names)
	--skip-dxf-index-normalization   Skip differential format index normalization`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;
	const completed = <string | undefined>args.completed;
	const addin = <string | undefined>args.addin;
	const xmlOnly = !!args["xml-only"];
	const vbaOnly = !!args["vba-only"];
	const skipSheetNameNormalization = !!args["skip-sheet-name-normalization"];
	const skipDxfIndexNormalization = !!args["skip-dxf-index-normalization"];
	const background =
		typeof args.background === "boolean" ? args.background : await resolveBackgroundMode();

	await exportProject({
		target,
		completed,
		addin,
		xmlOnly,
		vbaOnly,
		skipSheetNameNormalization,
		skipDxfIndexNormalization,
		background
	});
}
