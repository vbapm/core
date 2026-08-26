import dedent from "@timhall/dedent";
import { greenBright, redBright, yellowBright } from "@timhall/ansi-colors";
import { existsSync } from "fs";
import meant from "meant";
import mri, { Args } from "mri";
import { version } from "../../package.json";
import { env } from "../env";
import { cleanError, CliError, ErrorCode, isCliError } from "../errors";
import { checkForUpdate, checkDualInstall, updateAvailable, updateVersion } from "../installer";
import { Message } from "../messages";
import { isRunError, closePowerShellSession } from "../utils/run";
import { joinCommas } from "../utils/text";

Error.stackTraceLimit = Infinity;

const debug = env.debug("vbapm:main");

type Command = (args: Args) => Promise<void>;
const commands: { [name: string]: () => Promise<Command> } = {
	new: async () => (await import("./vbapm-new")).default,
	init: async () => (await import("./vbapm-init")).default,
	add: async () => (await import("./vbapm-add")).default,
	build: async () => (await import("./vbapm-build")).default,
	config: async () => (await import("./vbapm-config")).default,
	test: async () => (await import("./vbapm-test")).default,
	extract: async () => (await import("./vbapm-extract")).default,
	export: async () => (await import("./vbapm-export")).default,
	update: async () => (await import("./vbapm-update")).default,
	open: async () => (await import("./vbapm-open")).default,
	close: async () => (await import("./vbapm-close")).default,
	run: async () => (await import("./vbapm-run")).default,
	version: async () => (await import("./vbapm-version")).default,
	manifest: async () => (await import("./vbapm-manifest")).default
};

const args = mri(process.argv.slice(2), {
	alias: {
		v: "version",
		h: "help"
	}
});

if (args.debug) {
	let debug = args.debug;
	if (debug === true) debug = "*";
	else if (Array.isArray(debug)) debug = debug.join(",");

	const filters = (<string>debug).split(",").map(filter => `vbapm:${filter}`);
	const existing = process.env.DEBUG ? process.env.DEBUG.split(",") : [];

	process.env.DEBUG = existing.concat(filters).join(",");
}

const help = dedent`
  vbapm v${version}

  Usage: vbapm [command] [options]

  Commands:
    - new           Create a new project / package in a new directory
    - init          Initialize a new project / package in the current directory
    - add           Create and register a new source file in vbaproject.toml
    - build         Build project from manifest
    - config        Get and set repository/global configuration
    - test          Run tests for built target
    - extract       Extract src from built target
    - export        (deprecated) Extract src from built target
    - update        Update VBA source in a built target
    - open          Open the current built target file
    - close         Close the current built target file
    - run           Run macro in document / add-in
    - help          Outputs this message or the help of the given command

  Options:
    -h, --help      Output usage information
    -v, --version   Output the version number

  Use 'vbapm help COMMAND' for help on specific commands.
  Visit https://vba-blocks.com to learn more about vbapm.`;

const updateAvailableMessage = () => {
	const isStandalone = existsSync(env.bin);
	if (isStandalone) {
		return dedent`
		  \n${greenBright("New Update!")} ${updateVersion()!}

		  A new version of vbapm is available.
		  Visit https://vba-blocks.com/update for more information.`;
	}
	return dedent`
	  \n${greenBright("New Update!")} ${updateVersion()!}

	  A new version of vbapm is available.
	  Run "npm update -g vbapm" to update.`;
};

process.title = "vbapm";
process.on("unhandledRejection", handleError);
process.on("uncaughtException", handleError);

main()
	.then(async command => {
		// Shut down any persistent PowerShell/Excel session we may own before
		// exiting, so we don't leak a background Excel instance.
		try {
			await closePowerShellSession();
		} catch {
			// best-effort
		}

		// Work around Node.js v23+ libuv assertion crash on Windows
		// (nodejs/node#56645). The export command, especially to an
		// empty project, can complete before V8's wasm task runner
		// has drained its pending delayed tasks. Calling
		// process.exit(0) fires while the handle is still closing:
		//   ASSERT(!(handle->flags & UV_HANDLE_CLOSING))
		// On the export path we omit process.exit entirely, letting
		// the event loop drain naturally so no pending task is
		// interrupted. All other commands exit immediately via
		// process.exit(0) as before.
		// TODO: Remove once nodejs/node#61999 is merged and released.
		if ((command === "extract" || command === "export") && process.platform === "win32") {
			// Let the event loop drain naturally — no process.exit()
			return;
		}
		process.exit(0);
	})
	.catch(handleError);

async function main(): Promise<string | undefined> {
	let [command] = args._;

	if (!command) {
		if (args.version) {
			console.log(version);
		} else {
			console.log(help);

			if (updateAvailable()) {
				env.reporter.log(Message.UpdateAvailable, updateAvailableMessage());
			}
		}

		warnIfDualInstall();

		return undefined;
	}

	if (command === "help") {
		command = args._[1];

		if (!command) {
			console.log(help);
			warnIfDualInstall();
			return undefined;
		}

		args._ = [command];
		args.help = true;
	}
	command = command.toLowerCase();

	const available = Object.keys(commands);
	if (!available.includes(command)) {
		const approximate = meant(command, available);
		const did_you_mean = approximate.length
			? `, did you mean "${meant(command, available)}"?`
			: ".";
		const list = joinCommas(available.map(name => `"${name}"`));

		throw new CliError(
			ErrorCode.UnknownCommand,
			dedent`
        Unknown command "${command}"${did_you_mean}

        Available commands are ${list}.
        Try "vbapm help" for more information.
      `
		);
	}

	// Remove command from args
	args._ = args._.slice(1);

	let subcommand: (args: Args) => Promise<void>;
	try {
		debug(`loading "./vbapm-${command}.js"`);
		subcommand = await commands[command]();
	} catch (err: any) {
		throw new Error(`Failed to load command "${command}".\n${err?.stack || err}`);
	}

	debug(`starting "${command}" with args ${JSON.stringify(args)}`);
	const [has_update_available] = await Promise.all([checkForUpdate(), subcommand(args)]);

	if (has_update_available) {
		env.reporter.log(Message.UpdateAvailable, updateAvailableMessage());
	}

	return command;
}

function warnIfDualInstall() {
	const dualInstallWarning = checkDualInstall();
	if (dualInstallWarning) {
		console.warn(`\n${yellowBright("Warning:")} ${dualInstallWarning}`);
	}
}

export function handleError(err: Error | CliError | any, _promise?: Promise<any>) {
	const { message } = cleanError(err);

	console.error(`${redBright("ERROR")} ${message}`);

	// TODO
	// if (err.code) {
	//   console.log(
	//     chalk`\n{dim Visit https://vba-blocks.com/errors/${
	//       err.code
	//     } for more information}`
	//   );
	// }

	// Couldn't import debug, so log directly if debugging anything
	if (process.env.DEBUG) {
		console.error(err);

		if (isCliError(err) && err.underlying) {
			console.error("underlying", err.underlying);
		}
		if (isRunError(err)) {
			console.error("result", err.result);
		}
	}

	process.exit(1);
}
