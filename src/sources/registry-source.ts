import dedent from "@timhall/dedent";
import { existsSync, renameSync, unlinkSync } from "fs";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Dependency, RegistryDependency } from "../manifest/dependency";
import { Message } from "../messages";
import { download } from "../utils/download";
import {
	checksum as getChecksum,
	ensureDir,
	move,
	pathExists,
	readFile,
	remove,
	tmpFile
} from "../utils/fs";
import { clone, isGitRepository, pull } from "../utils/git";
import { basename, dirname, join, sanitize } from "../utils/path";
import { unzip } from "../utils/zip";
import { getRegistrationId, getRegistrationSource, Registration } from "./registration";
import { Source } from "./source";

const debug = env.debug("vbapm:registry-source");

export interface RegistryOptions {
	name: string;
	index: string;
	packages: string;
}

export class RegistrySource implements Source {
	name: string;
	local: { index: string; packages: string };
	remote: { index: string; packages: string };
	private sources: string;
	private pulling?: Promise<void>;
	private upToDate: boolean;

	constructor({ name, index, packages }: RegistryOptions) {
		this.name = name;
		this.local = {
			index: join(env.registry, name),
			packages: join(env.packages, name)
		};
		this.remote = { index, packages };

		this.sources = join(env.sources, name);
		this.upToDate = false;
	}

	async resolve(dependency: Dependency): Promise<Registration[]> {
		if (!this.upToDate) await this.pull();

		const { name } = <RegistryDependency>dependency;
		const path = getPath(this.local.index, name);

		if (!(await pathExists(path))) {
			throw new CliError(
				ErrorCode.DependencyNotFound,
				`Dependency "${name}" not found in registry "${this.name}".`
			);
		}

		const data = await readFile(path, "utf8");
		const registrations: Registration[] = data
			.split(/\r?\n/)
			.filter((line: string) => !!line)
			.map((line: string) => JSON.parse(line))
			.filter((value: any) => value && !value.yanked)
			.map((value: string) => parseRegistration(this.name, value));

		return registrations;
	}

	async fetch(registration: Registration): Promise<string> {
		const url = getRemotePackage(this.remote.packages, registration);
		const file = getLocalPackage(this.local.packages, registration);

		const [_, checksum] = registration.source.split("#", 2);
		const [algorithm, hash] = checksum.split("-");

		if (!(await pathExists(file))) {
			const unverifiedFile = await tmpFile();
			try {
				await download(url, unverifiedFile);
			} catch (err: any) {
				throw new CliError(
					ErrorCode.SourceDownloadFailed,
					`Failed to download "${registration.source}".`,
					err
				);
			}

			const comparison = await getChecksum(unverifiedFile, algorithm);
			if (comparison !== hash) {
				const expectedSignature = `${algorithm}-${hash}`;
				const receivedSignature = `${algorithm}-${comparison}`;

				debug(`Checksum failed for ${unverifiedFile}`);
				debug(`Expected ${expectedSignature}, received ${receivedSignature}`);

				throw new CliError(
					ErrorCode.DependencyInvalidChecksum,
					dedent`
            Dependency "${registration.name}" failed validation.

            The downloaded file signature for ${registration.id}
            does not match the signature in the registry.

            Expected: ${expectedSignature}
            Received: ${receivedSignature}
          `
				);
			}

			await move(unverifiedFile, file);
		}

		return await extractSource(file, this.sources, registration);
	}

	async pull() {
		if (this.pulling) return this.pulling;

		this.pulling = pullIndex(this.local.index, this.remote.index);
		await this.pulling;

		this.upToDate = true;
		this.pulling = undefined;
	}
}

export async function pullIndex(local: string, remote: string) {
	const hasLocalDirectory = await pathExists(local);
	if (hasLocalDirectory && !isGitRepository(local)) {
		// For local registry, skip clone + pull
		// if directory exists without git repository
		env.reporter.log(
			Message.RegistrySourceLocalOnly,
			"(local registry is not a git repository, skipping pull)"
		);
		return;
	}

	if (!hasLocalDirectory) {
		await ensureDir(dirname(local));

		try {
			await clone(remote, basename(local), dirname(local));
		} catch (err: any) {
			throw new CliError(
				ErrorCode.RegistryCloneFailed,
				`Failed to clone registry from ${remote}`,
				err
			);
		}
	}

	try {
		await pull(local);
	} catch (err) {
		debug("Pull failed", err);

		// If pull fails (but repository exists)
		// treat as offline and still attempt to resolve
		env.reporter.log(
			Message.RegistrySourceSkipPull,
			"(failed to update local registry, resolving with previously loaded values)"
		);
	}
}

export function parseRegistration(registry: string, value: any): Registration {
	const { name, vers: version, cksum: checksum } = value;

	const dependencies: RegistryDependency[] = value.deps.map((dep: any) => {
		const { name, req } = dep;
		const dependency: RegistryDependency = {
			name,
			registry,
			version: req
		};

		return dependency;
	});

	return {
		id: getRegistrationId(name, version),
		source: getRegistrationSource("registry", registry, checksum),
		name,
		version,
		dependencies
	};
}

export function sanitizePackageName(name: string): string {
	return sanitize(name.replace("/", "--"));
}

function getPath(index: string, name: string): string {
	return join(index, sanitizePackageName(name));
}

export function getRemotePackage(packages: string, registration: Registration): string {
	const { name, version } = registration;
	return `${packages}/${sanitizePackageName(name)}-v${version}.block`;
}

export function getLocalPackage(packages: string, registration: Registration): string {
	const { name, version } = registration;
	return join(packages, `${sanitizePackageName(name)}-v${version}.block`);
}

export function getSource(sources: string, registration: Registration): string {
	const { name, version } = registration;
	return join(sources, `${sanitizePackageName(name)}-v${version}`);
}

/**
 * Extract a downloaded package into the shared sources cache atomically.
 *
 * `fetch` can run concurrently (parallel Jest workers each run their own `vba`
 * process), so this must be safe when several processes fetch the same
 * registration at once. Each caller extracts into a unique temp directory and
 * then atomically renames it into place; a caller that loses the rename race
 * reuses the directory the winner published. This avoids the
 * `ENOENT: utime` crash that happened when two unzips wrote to the same
 * directory while one renamed `vba-block.toml` → `vbaproject.toml`.
 */
export async function extractSource(
	file: string,
	sources: string,
	registration: Registration
): Promise<string> {
	const src = getSource(sources, registration);

	// Fast path: already extracted (legacy or current manifest present).
	if (await isSourceExtracted(src)) {
		return src;
	}

	// Extract into a unique sibling temp dir on the same volume so the final
	// rename is atomic.
	const tmp = join(sources, `.${basename(src)}-tmp-${process.pid}-${randomSuffix()}`);
	await ensureDir(tmp);

	try {
		await unzip(file, tmp);

		// Normalize the legacy manifest inside the private temp dir. Uses native
		// fs (not the mockable wrapper) because unzip writes real files to disk.
		const legacyManifest = join(tmp, "vba-block.toml");
		const newManifest = join(tmp, "vbaproject.toml");
		if (existsSync(legacyManifest)) {
			if (existsSync(newManifest)) {
				unlinkSync(legacyManifest);
			} else {
				renameSync(legacyManifest, newManifest);
			}
		}

		try {
			renameSync(tmp, src);
		} catch (err) {
			// Another process published `src` first; reuse theirs. If it isn't
			// actually ready, surface the rename error instead.
			if (!(await isSourceExtracted(src))) {
				throw err;
			}
		}
	} finally {
		// Best-effort cleanup of our temp dir (no-op after a successful rename).
		await remove(tmp).catch(() => {});
	}

	return src;
}

async function isSourceExtracted(src: string): Promise<boolean> {
	return (
		(await pathExists(join(src, "vbaproject.toml"))) ||
		(await pathExists(join(src, "vba-block.toml")))
	);
}

function randomSuffix(): string {
	return Math.random().toString(36).slice(2, 10);
}
