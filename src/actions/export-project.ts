import dedent from "@timhall/dedent";
import { yellowBright } from "@timhall/ansi-colors";
import { exportTo } from "../addin";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Target, TargetType } from "../manifest/target";
import { Message } from "../messages";
import { fetchDependencies, loadProject } from "../project";
import { exportTarget } from "../targets";
import { emptyDir, ensureDir, remove } from "../utils/fs";
import { join, sanitize } from "../utils/path";
import { isTargetSaved } from "./check-saved";

export interface ExportOptions {
	target?: string;
	completed?: string;
	addin?: string;
	xmlOnly?: boolean;
	vbaOnly?: boolean;
	skipSheetNameNormalization?: boolean;
}

export async function exportProject(options: ExportOptions = {}) {
	if (options.xmlOnly && options.vbaOnly) {
		throw new CliError(
			ErrorCode.ExportOptionsConflict,
			"--xml-only and --vba-only are mutually exclusive."
		);
	}

	const skipStaging = !!options.completed || !!options.xmlOnly;
	const steps = skipStaging ? 2 : 3;
	env.reporter.log(Message.ExportProjectLoading, `[1/${steps}] Loading project...`);

	const project = await loadProject();

	if (!options.target && !project.manifest.target) {
		throw new CliError(
			ErrorCode.ExportNoTarget,
			dedent`
        No default target found for project,
        use --target TYPE to export from a specific target.
      `
		);
	}

	let target: Target | undefined;
	let blankTarget = false;
	if (project.manifest.target) {
		if (!options.target || options.target === project.manifest.target.type) {
			target = project.manifest.target;
		}
	} else {
		const type = <TargetType>options.target;
		const name = project.manifest.name;

		target = {
			type,
			name,
			path: `targets/${type}`,
			filename: `${sanitize(name)}.${type}`
		};
		blankTarget = true;
	}

	if (!target) {
		throw new CliError(
			ErrorCode.ExportNoMatching,
			`No matching target found for type "${options.target!}" in project.`
		);
	}

	// Warn if the workbook is open in Excel with unsaved changes — the exported XML
	// will only reflect the last saved state of the file on disk.
	if (!options.vbaOnly) {
		const saved = await isTargetSaved(target, project);
		if (saved === false) {
			console.warn(
				`\n${yellowBright("WARN:")} The workbook has unsaved changes. The XML data will only reflect the last saved version.`
			);
		}
	}

	const dependencies = await fetchDependencies(project);
	let staging: string;

	try {
		if (!options.completed) {
			staging = join(project.paths.staging, "export");

			await ensureDir(staging);
			await emptyDir(staging);

			if (!options.xmlOnly) {
				env.reporter.log(
					Message.ExportToStaging,
					`\n[2/3] Exporting src from "${target.filename}"`
				);
				await exportTo(project, target, staging, options);
			}
		} else {
			staging = options.completed;
		}

		env.reporter.log(Message.ExportToProject, `\n[${steps}/${steps}] Updating project`);
		await exportTarget(target, { project, dependencies, blankTarget }, staging, {
			xmlOnly: options.xmlOnly,
			vbaOnly: options.vbaOnly
		});
	} finally {
		// @ts-ignore Variable is used before being assigned
		if (staging) await remove(staging);
	}
}
