import { addins, extensionToApplication } from "../addin";
import { Target } from "../manifest/target";
import { Project } from "../project";
import { pathExists } from "../utils/fs";
import { join, resolve } from "../utils/path";
import { run } from "../utils/run";

/**
 * Checks whether the built target workbook has been saved in the running Excel instance.
 *
 * Returns:
 *   true  — workbook is saved (or Excel is not running / workbook is not open)
 *   false — workbook has unsaved changes
 *   null  — could not determine (addin missing, unexpected error)
 */
export async function isTargetSaved(target: Target, project: Project): Promise<boolean | null> {
	const application = extensionToApplication(target.type);
	const addin = addins[application];
	const resolvedFile = resolve(join(project.paths.build, target.filename));

	if (!(await pathExists(addin))) return null;

	try {
		const result = await run(application, addin, "Build.CheckFileSaved", [
			JSON.stringify({
				file: resolvedFile
			})
		]);
		// CheckFileSaved returns "saved:true" or "saved:false" in messages[0]
		const savedMessage = result.messages.find(m => m.startsWith("saved:"));
		if (savedMessage) {
			return savedMessage === "saved:true";
		}
		return null;
	} catch {
		return null;
	}
}
