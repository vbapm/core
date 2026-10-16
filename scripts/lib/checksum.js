import fsExtra from "fs-extra";
import hash from "./hash.js";

const { readFile } = fsExtra;

export default async function checksum(file, algorithm = "sha256") {
	const data = await readFile(file);
	return hash(data, { algorithm });
};
