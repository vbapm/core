import { join } from "node:path";
import walkSync from "walk-sync";

export default function ls(dir) {
	return walkSync(dir, { directories: false }).map(path => join(dir, path));
};
