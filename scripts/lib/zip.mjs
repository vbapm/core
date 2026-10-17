/**
 * zip.mjs — Utility function to create a zip archive from a set of files.
 *
 * This is used by the publish.js script to create a .block archive from the package files.
 *
 * Usage:
 *   const { default: zip } = await import("./lib/zip.mjs");
 *   await zip({ "path/to/file1": "name/in/archive1", "path/to/file2": "name/in/archive2" }, "output.zip");
 */

import { createWriteStream } from "node:fs";
import { TarArchive, ZipArchive } from "archiver";

export default async function zip(input, dest, type = "zip", options = {}) {
	return new Promise((resolve, reject) => {
		try {
			const output = createWriteStream(dest);
						const archive = type === "tar" ? new TarArchive(options) : new ZipArchive(options);

			output.on("close", resolve);
			output.on("error", reject);

			archive.pipe(output);
			archive.on("error", reject);

			for (const [path, name] of Object.entries(input)) {
				archive.file(path, { name });
			}

			archive.finalize();
		} catch (err) {
			reject(err);
		}
	});
}
