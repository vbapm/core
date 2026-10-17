import { createWriteStream } from "fs";
import fetch from "node-fetch";
import { ensureDir } from "./fs";
import { dirname } from "./path";

export async function download(url: string, dest: string): Promise<void> {
	await ensureDir(dirname(dest));

	return fetch(url).then(response => {
		const body = response.body;
		if (!body) {
			throw new Error(`Response body is empty for ${url}`);
		}

		const file = createWriteStream(dest);

		return new Promise<void>((resolve, reject) => {
			body
				.pipe(file)
				.on("finish", () => resolve())
				.on("error", reject);
		});
	});
}
