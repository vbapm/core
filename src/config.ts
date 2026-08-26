import { env } from "./env";
import { GitSource, PathSource, RegistrySource, Sources } from "./sources";
import { ensureDir, pathExists, readFile, writeFile } from "./utils/fs";
import { dirname, join } from "./utils/path";
import { convert as convertToml, parse as parseToml } from "./utils/toml";

export type Registry = {} | { [name: string]: { index: string; packages: string } };

export interface Flags {}

export interface ToolSettings {
	background?: boolean;
	[name: string]: unknown;
}

export interface ToolSettingsOptions {
	global?: boolean;
	file?: string;
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

export function getGlobalConfigFile(): string {
	return join(env.bin, "vba.toml");
}

export async function getLocalConfigFile(dir = env.cwd): Promise<string | undefined> {
	let current = dir;

	while (true) {
		const candidate = join(current, "vbaproject.toml");
		if (await pathExists(candidate)) return join(current, "vba.toml");

		const parent = dirname(current);
		if (parent === current) return undefined;
		current = parent;
	}
}

export async function loadToolSettings(
	options: ToolSettingsOptions = {}
): Promise<ToolSettings> {
	const file = options.file ?? (options.global ? getGlobalConfigFile() : await getLocalConfigFile());
	if (!file || !(await pathExists(file))) return {};

	const raw = await readFile(file);
	const parsed = await parseToml(raw.toString());
	return parsed && typeof parsed === "object" ? (parsed as ToolSettings) : {};
}

export async function loadEffectiveToolSettings(): Promise<ToolSettings> {
	const globalSettings = await loadToolSettings({ global: true });
	const localSettings = await loadToolSettings({ global: false });
	return { ...globalSettings, ...localSettings };
}

export async function saveToolSettings(
	settings: ToolSettings,
	options: ToolSettingsOptions = {}
): Promise<string> {
	const file =
		options.file ??
		(options.global ? getGlobalConfigFile() : (await getLocalConfigFile()) ?? join(env.cwd, "vba.toml"));

	await ensureDir(dirname(file));
	const existing = await loadToolSettings({ file });
	const updated = { ...existing, ...settings };
	const content = await convertToml(updated);
	await writeFile(file, content);
	return file;
}

export async function resolveBackgroundMode(value?: boolean): Promise<boolean> {
	if (typeof value === "boolean") return value;

	const settings = await loadEffectiveToolSettings();
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
