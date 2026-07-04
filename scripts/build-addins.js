const { join } = require("path");
const vba = require("../lib/index");

main().catch(err => {
	console.error(err.message);
	if (err.underlying && err.underlying.result) {
		console.log(err.underlying.result);
	}
	process.exit(1);
});

async function main() {
	await vba.buildProject({ addin: join(__dirname, "bootstrap/build/bootstrap.xlsm") });
}
