import dedent from "@timhall/dedent";
import { Args } from "mri";
import { initProject } from "../actions/init-project";

const help = dedent`
  Initialize a new project in the current directory

  Usage: vbapm init [options]

  Options:
    --target=TYPE           Add target of type TYPE to project (e.g. xlsm)
    --from=PATH             Create target and src from workbook/document
    --name=NAME             Set project name (default = --from or directory name)
    --package               Initialize as package
    --no-git                Skip initializing git repository
    --no-conf               Skip copying .gitignore, .gitattributes, .editorconfig

  Examples:
  vbapm init --target xlsm
  vbapm init --name calculations --package
  `;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;
	const from = <string | undefined>args.from;
	const name = <string | undefined>args.name;
	const pkg = !!args.package;
	const git = "git" in args ? <boolean>args.git : true;
	const configTemplates = "conf" in args ? <boolean>args.conf : true;

	await initProject({ target, from, name, pkg, git, configTemplates });
}
