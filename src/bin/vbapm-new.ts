import dedent from "@timhall/dedent";
import { Args } from "mri";
import { createProject } from "../actions/create-project";

const help = dedent`
  Create a new project / package in a new directory

  Usage: vbapm new <name> [options]

  Options:
    <name>                  Project/package name (optionally, with extension)
    --target=TYPE           Add target of type TYPE to project (e.g. xlsm)
    --from=PATH             Create target and src from workbook/document
    --package               Create as package
    --no-git                Skip initializing git repository
    --no-conf               Skip copying .gitignore, .gitattributes, .editorconfig

  Examples:
  vbapm new analysis.xlsm
  vbapm new analysis --target xlsm
  vbapm new calculations --package
  `;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const [name] = args._;
	const target = <string | undefined>args.target;
	const from = <string | undefined>args.from;
	const pkg = !!args.package;
	const git = "git" in args ? <boolean>args.git : true;
	const configTemplates = "conf" in args ? <boolean>args.conf : true;

	await createProject({ name, target, from, pkg, git, configTemplates });
}
