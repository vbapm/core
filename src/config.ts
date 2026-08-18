import { env } from "./env";
import { GitSource, PathSource, RegistrySource, Sources } from "./sources";
import { pathExists, readFile, writeFile } from "./utils/fs";
import { join } from "./utils/path";
import { convert as convertToml, parse as parseToml } from "./utils/toml";

export type Registry = {} | { [name: string]: { index: string; packages: string } };

export interface Flags {}

export interface ToolSettings {
	background?: boolean;
}

export interface Config {
	registry: Registry;
	flags: Flags;
	sources: Sources;
}

export interface ConfigValue {
	registry?: Registry;
	flags?: Flags;
}

export function normalizeBackground(value: unknown): boolean {
	if (typeof value === "boolean") return value;
	if (typeof value === "number") return value !== 0;
	if (typeof value === "string") {
		return /^(1|true|yes|on)$/i.test(value.trim());
	}
	return false;
}

export async function loadToolSettings(file = join(env.bin, "vba.toml")): Promise<ToolSettings> {
	if (!(await pathExists(file))) return {};

	const raw = await readFile(file);
	const parsed = await parseToml(raw.toString());
	return parsed && typeof parsed === "object" ? (parsed as ToolSettings) : {};
}

export async function saveToolSettings(settings: ToolSettings, file = join(env.bin, "vba.toml")) {
	const existing = await loadToolSettings(file);
	const updated = { ...existing, ...settings };
	const content = await convertToml(updated);
	await writeFile(file, content);
}

export async function resolveBackgroundMode(value?: boolean): Promise<boolean> {
	if (typeof value === "boolean") return value;

	const envValue = env.values.VBA_BACKGROUND_BUILD;
	if (typeof envValue !== "undefined") return normalizeBackground(envValue);

	const settings = await loadToolSettings();
	return normalizeBackground(settings.background);
}

const empty: ConfigValue = { registry: {}, flags: {} };
const defaults: ConfigValue = {
	registry: {
		"vba-blocks": {
			index: "https://github.com/vba-blocks/registry",
			packages: "https://packages.vba-blocks.com"
		}
	},
	flags: {}
};

/**
 * Load config, from local, user, and environment values
 *
 * - env:config/config.toml (user)
 * - Search for .vbapm/config.toml up from cwd (local)
 * - Load VBAPM_* from environment (override)
 */
export async function loadConfig(): Promise<Config> {
	const user: ConfigValue = {
		...empty,
		...(await readConfig(env.config))
	};

	const dir = await findConfig(env.cwd);
	const local: ConfigValue = {
		...empty,
		...(dir ? await readConfig(dir) : undefined)
	};

	const override = loadConfigFromEnv();

	const registry: Registry = {
		...defaults.registry,
		...user.registry,
		...local.registry,
		...override.registry
	};
	const flags: Flags = {
		...defaults.flags,
		...user.flags,
		...local.flags,
		...override.flags
	};

	const sources: Sources = {
		registry: {},
		git: new GitSource(),
		path: new PathSource()
	};

	for (const [name, { index, packages }] of Object.entries(registry)) {
		sources.registry[name] = new RegistrySource({
			name,
			index,
			packages
		});
	}

	return { registry, flags, sources };
}

// Read config from dir (if present)
export async function readConfig(dir: string): Promise<ConfigValue | undefined> {
	const file = join(dir, "config.toml");
	if (!(await pathExists(file))) return {};

	const raw = await readFile(file);
	const parsed = await parseToml(raw.toString());

	return parsed;
}

// Find config up from and including given dir
// (looking for .vbapm/config.toml)
export async function findConfig(_dir: string): Promise<string | undefined> {
	// TODO Search for .vbapm/config.toml starting at cwd
	return;
}

// Load config values from environment
export function loadConfigFromEnv(): ConfigValue {
	// TODO Load override config from VBAPM_* env variables
	return {};
}
