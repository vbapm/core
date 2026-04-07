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
		await writeFile(
			join(dir, ".gitignore"),
			dedent`
        # Template from: https://github.com/DecimalTurn/VBA-on-GitHub

        # Hidden metadata files created by macOS (Desktop Services Store)
        .DS_Store

        # Office temporary files
        ~$*

        # Access database lock files (laccdb, ldb)
        *.[lL][aA][cC][cC][dD][bB]
        *.[lL][dD][bB]

        # The following sections constitute a list of Office file extensions that support VBA.
        # If you want to exclude Office files from your repo, uncomment the corresponding file extensions.

        # Excel (xls, xlsb, xlsm, xlt, xltm, xla, xlam)
        #/*.[xX][lL][sS]
        #/*.[xX][lL][sS][bB]
        #/*.[xX][lL][sS][mM]
        #/*.[xX][lL][tT]
        #/*.[xX][lL][tT][mM]
        #/*.[xX][lL][aA]
        #/*.[xX][lL][aA][mM]

        # Word (doc, docm, dot, dotm)
        #/*.[dD][oO][cC]
        #/*.[dD][oO][cC][mM]
        #/*.[dD][oO][tT]
        #/*.[dD][oO][tT][mM]

        # Access (accda, accdb, accde, mdb, mde)
        #/*.[aA][cC][cC][dD][aA]
        #/*.[aA][cC][cC][dD][bB]
        #/*.[aA][cC][cC][dD][eE]
        #/*.[mM][dD][bB]
        #/*.[mM][dD][eE]

        # PowerPoint (ppt, pptm, pot, potm, pps, ppsm)
        #/*.[pP][pP][tT]
        #/*.[pP][pP][tT][mM]
        #/*.[pP][oO][tT]
        #/*.[pP][oO][tT][mM]
        #/*.[pP][pP][sS]
        #/*.[pP][pP][sS][mM]

        # vbapm specific folders
        /build
      `
		);
		await writeFile(
			join(dir, ".gitattributes"),
			dedent`
        # Template from: https://github.com/DecimalTurn/VBA-on-GitHub

        # Don't perform any line ending conversions by default

        # VBA extensions - Prevent LF normalization (bas, cls, frm, vba, doccls)
        *.[bB][aA][sS]              -text diff
        *.[cC][lL][sS]              -text diff
        *.[fF][rR][mM]              -text diff
        *.[vV][bB][aA]              -text diff
        *.[dD][oO][cC][cC][lL][sS]  -text diff

        # VBA extensions - Mark as binary (frx)
        *.[fF][rR][xX]              binary

        ############################################################################
        # Other languages in the VB family 
        ############################################################################

        # VBS extensions - Prevent LF normalization (vbs)
        *.[vV][bB][sS]              -text diff

        # VB6 extensions - Prevent LF normalization (ctl, dsr, dob, pag, vbg, vbl, vbp, vbr, vbw)
        *.[cC][tT][lL]              -text diff
        *.[dD][sS][rR]              -text diff
        *.[dD][oO][bB]              -text diff
        *.[pP][aA][gG]              -text diff
        *.[vV][bB][gB]              -text diff
        *.[vV][bB][lL]              -text diff
        *.[vV][bB][pP]              -text diff
        *.[vV][bB][rR]              -text diff
        *.[vV][bB][wW]              -text diff

        # VB6 extensions - Mark as binary (ctx, dox, pgx)
        *.[cC][tT][xX]              binary
        *.[dD][oO][xX]              binary
        *.[pP][gG][xX]              binary

        # twinBASIC sources extensions - Prevent LF normalization (twin, tbform)
        *.[tT][wW][iI][nN]          -text diff linguist-language=VB6
        *.[tT][bB][fF][oO][rR][mM]  -text diff linguist-language=JSON

        # twinBASIC project extension - Mark as binary (twinproj)
        *.[tT][wW][iI][nN][pP][rR][oO][jJ]  binary

        ############################################################################
        # Other Windows-specific extensions
        ############################################################################

        # INI file extensions - Prevent LF normalization (ini)
        *.[iI][nN][iI]              -text diff

        # Batch scripts - Prevent LF normalization (cmd, bat)
        *.[cC][mM][dD]              -text diff
        *.[bB][aA][tT]              -text diff

        ############################################################################
        # Additional Sections
        ############################################################################

        # VBA packages from XVBA (https://marketplace.visualstudio.com/items?itemName=local-smart.excel-live-server)
        **/xvba_modules/**          linguist-vendored=true

        # Excel documents (xla, xlam, xls, xlsb, xlsm and xlsx)
        *.[xX][lL][aA]              binary
        *.[xX][lL][aA][mM]          binary
        *.[xX][lL][sS]              binary
        *.[xX][lL][sS][bB]          binary
        *.[xX][lL][sS][mM]          binary
        *.[xX][lL][sS][xX]          binary

        # Word documents (rtf, doc and docx)
        *.[rR][tT][fF]              diff=astextplain
        *.[dD][oO][cC]              diff=astextplain
        *.[dD][oO][cC][xX]          diff=astextplain

        # PowerPoint documents (ppt and pptx)
        *.[pP][pP][tT]              binary
        *.[pP][pP][tT][xX]          binary                  

        # Access documents (accdb)
        *.[aA][cC][cC][dD][bB]      binary

        # Images (jpg, png, bmp, gif, ico)           
        *.[jJ][pP][gG]              binary
        *.[pP][nN][gG]              binary
        *.[bB][mM][pP]              binary
        *.[gG][iI][fF]              binary
        *.[iI][cC][oO]              binary

        # Compressed files (zip, cab, 7z, gz)
        *.[zZ][iI][pP]              binary
        *.[cC][aA][bB]              binary
        *.[7][zZ]                   binary
        *.[gG][zZ]                  binary

        # Executables (exe, dll)
        *.[eE][xX][eE]              binary
        *.[dD][lL][lL]              binary

        # Other (pdf)
        *.[pP][dD][fF]              diff=astextplain
      `
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
