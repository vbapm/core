import { promisify } from "node:util";
import { execFile as execFileCallback } from "node:child_process";

const execFile = promisify(execFileCallback);

const remote = "https://github.com/vba-blocks/registry.git";

export default {
	async pull({ dir }) {
		try {
			await execFile("git", ["pull"], { cwd: dir });
		} catch (error) {
			throw new Error(`git pull failed in "${dir}". ${error.message}`);
		}
	},

	async commit({ dir, message }) {
		try {
			await execFile("git", ["add", "-A"], { cwd: dir });
			await execFile("git", ["commit", "-am", message], { cwd: dir });
		} catch (error) {
			throw new Error(
				`git add -A && git commit -am "${message}" failed in "${dir}". ${error.message}`
			);
		}
	},

	async push({ dir }) {
		try {
			await execFile("git", ["push"], { cwd: dir });
		} catch (error) {
			throw new Error(`git push failed in "${dir}". ${error.message}`);
		}
	},

	async clone({ dir }) {
		try {
			console.log(`git clone ${remote} registry`, dir);
			await execFile("git", ["clone", remote, "registry"], { cwd: dir });
		} catch (error) {
			throw new Error(`git clone failed in "${dir}". ${error.message}`);
		}
	}
};
