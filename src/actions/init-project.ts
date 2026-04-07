import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Manifest, writeManifest } from "../manifest";
import { TargetType } from "../manifest/target";
import { initProject as init } from "../project";
import { addTarget } from "../targets/add-target";
import { ensureDir, pathExists, writeFile } from "../utils/fs";
import { init as git_init } from "../utils/git";
import { basename, extname, join } from "../utils/path";

export interface InitOptions {
	name?: string;
	dir?: string;
	target?: string;
	from?: string;
	pkg: boolean;
	git: boolean;
}

export async function initProject(options: InitOptions) {
	let { name, dir = env.cwd, target: targetType, from, pkg: asPackage, git } = options;

	if (await pathExists(join(dir, "vbaproject.toml"))) {
		throw new CliError(
			ErrorCode.InitAlreadyInitialized,
			`A vbapm project already exists in this directory.`
		);
	}

	if (from && !(await pathExists(from))) {
		throw new CliError(ErrorCode.FromNotFound, `The "from" document was not found at "${from}".`);
	}

	name = name || (from ? basename(from, extname(from)) : basename(dir));

	if (!name) {
		throw new CliError(
			ErrorCode.InitNameRequired,
			dedent`
        Unable to determine name from current directory or --from.
        --name NAME is required to initialize this project.
      `
		);
	}
	if (!targetType && !from && name.includes(".")) {
		const parts = name.split(".");
		targetType = parts.pop();
		name = parts.join(".");
	}

	if (!asPackage && !targetType && !from) {
		throw new CliError(
			ErrorCode.InitTargetRequired,
			dedent`
        --target or --from is required for vbapm projects.
        (e.g. vbapm init --target xlsm)
      `
		);
	}

	await ensureDir(join(dir, "src"));

	if (git && !(await pathExists(join(dir, ".git")))) {
		await git_init(dir);
		await writeFile(join(dir, ".gitignore"), `/build`);
		await writeFile(
			join(dir, ".gitattributes"),
			`* text=auto\n*.bas text eol=crlf\n*.cls text eol=crlf`
		);
		await writeFile(
			join(dir, ".editorconfig"),
			dedent`
        # EditorConfig is awesome: http://EditorConfig.org
        # Template from: https://github.com/DecimalTurn/VBA-on-GitHub

        # top-most EditorConfig file
        root = true

        # Properties for VBA, VB6 and twinBASIC file extensions
        [*.{bas,cls,frm,vba,doccls,ctl,dsr,twin,tbform}]
        indent_style = space
        indent_size = 4
        end_of_line = crlf
        # Avoid line too long error (https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/line-too-long):
        max_line_length = 1023
        insert_final_newline = true
        trim_trailing_whitespace = true

        # The character set property isn't widely supported, but you can still add it. It just won't do anything if unsupported by your editor.
        # Reference: https://github.com/editorconfig/editorconfig/issues/209#issuecomment-445241830
        # Eg.:
        # charset = windows-1252
        # charset = us-ascii
        
        ###############################
        
        # Properties for Office OOXML files (e.g. .xml, .rels)
        [{*.rels,.rels,*.xml}]
        indent_style = space
        indent_size = 2
        end_of_line = crlf
      `
		);
	}

	const project = await init(name, dir, {
		type: asPackage ? "package" : "project"
	});

	if (from) {
		targetType = extname(from).replace(".", "");
	}
	if (targetType) {
		const dependencies: Manifest[] = [];
		await addTarget(<TargetType>targetType, { project, dependencies }, { from });
	}

	await writeManifest(project.manifest, project.paths.dir);
}
