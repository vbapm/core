import { dirname } from "node:path";
import { createWriteStream } from "node:fs";
import fetch from "node-fetch";
import { ensureDir } from "fs-extra";

export default async function download(url, dest) {
	await ensureDir(dirname(dest));

	return fetch(url).then(response => {
		const file = createWriteStream(dest);

		return new Promise((resolve, reject) => {
			response.body
				.pipe(file)
				.on("finish", () => resolve())
				.on("error", reject);
		});
	});
};
