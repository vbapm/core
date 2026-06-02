import { differenceInCalendarDays } from "date-fns";
import { execSync } from "child_process";
import { existsSync } from "fs";
import { gt as semverGreaterThan } from "semver";
import { version as currentVersion } from "../package.json";
import { cache } from "./cache";
import { env } from "./env";
import { getLatestRelease } from "./utils/github";
import { join } from "./utils/path";

const debug = env.debug("vbapm:installer");

// When installed via npm, the bin/ directory doesn't exist inside
// node_modules/vbapm/. When running as standalone, it does.
const IS_STANDALONE = existsSync(env.bin);

export function updateVersion(): string | undefined {
	return cache.latest_version;
}

export function updateAvailable(): boolean {
	// Previously checked version is greater
	const latestKnownVersion = cache.latest_version;
	return !!latestKnownVersion && semverGreaterThan(latestKnownVersion, currentVersion);
}

export async function checkForUpdate(): Promise<boolean> {
	// npm users manage their own updates via npm update
	if (!IS_STANDALONE) return false;

	// Only check for new version once per day
	const lastChecked = cache.latest_version_checked;
	if (lastChecked && differenceInCalendarDays(new Date(lastChecked), Date.now()) < 1)
		return updateAvailable();

	// Allow skipping from the outside
	// set VBAPM_SKIP_UPDATE_CHECK=1
	//
	// (maybe this should be added to config)
	if (parseInt(env.values.VBAPM_SKIP_UPDATE_CHECK, 10)) return false;

	cache.latest_version_checked = Date.now();

	try {
		const { tag_name: latestVersion } = await getLatestRelease({
			owner: "vba-blocks",
			repo: "vba-blocks"
		});
		cache.latest_version = latestVersion;

		return semverGreaterThan(latestVersion, currentVersion);
	} catch (error) {
		debug("Error loading latest release");
		debug(error);
		return false;
	}
}

/**
 * Detect if both standalone CLI and npm global installations exist.
 * Returns a warning message if both are found, or undefined if only one is present.
 */
export function checkDualInstall(): string | undefined {
	const standaloneBin = env.isWindows
		? join(process.env.APPDATA || "", "vbapm", "bin", "vba.cmd")
		: join(process.env.HOME || "", ".vbapm", "bin", "vba");

	const hasStandalone = existsSync(standaloneBin);

	let hasNpm = false;
	try {
		// Ask npm where its global prefix is, then check for vbapm there
		const npmPrefix = execSync("npm prefix -g", {
			encoding: "utf8",
			timeout: 5000,
			stdio: ["ignore", "pipe", "ignore"]
		}).trim();
		const npmVbapm = join(npmPrefix, "node_modules", "vbapm", "package.json");
		hasNpm = existsSync(npmVbapm);
	} catch {
		debug("Could not determine npm global prefix");
	}

	if (hasStandalone && hasNpm) {
		return [
			"vbapm is installed both as a standalone CLI and as an npm global package.",
			"This may cause version conflicts. Consider removing one:",
			IS_STANDALONE
				? "  - To remove the npm copy: npm uninstall -g vbapm"
				: "  - To remove the standalone copy: delete the vbapm folder from " +
					(env.isWindows ? "%APPDATA%\\vbapm" : "~/.vbapm")
		].join("\n");
	}

	return undefined;
}
