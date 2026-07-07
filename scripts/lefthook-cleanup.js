// Patch .git/hooks/pre-commit to produce clean, ANSI-free output.
// lefthook emits ANSI color codes and box-drawing that render as garbage
// in VS Code's Git dialog. This script:
//   1. Injects NO_COLOR + companion env vars after the shebang.
//   2. Wraps the final lefthook call to strip any remaining ANSI via sed.
// Runs after every `lefthook install` since the hook file is regenerated.
const fs = require("fs");
const path = require("path");

const hookPath = path.resolve(__dirname, "..", ".git", "hooks", "pre-commit");

if (!fs.existsSync(hookPath)) {
	console.log("patch-hook-no-color: .git/hooks/pre-commit not found, skipping");
	process.exit(0);
}

let content = fs.readFileSync(hookPath, "utf8");
let patched = false;

// Patch 1: inject ANSI-suppressing env vars after the shebang
if (!content.includes("export NO_COLOR=1")) {
	content = content.replace(
		"#!/bin/sh\n",
		"#!/bin/sh\nexport NO_COLOR=1\nexport FORCE_COLOR=0\nexport CLICOLOR=0\nexport TERM=dumb\n"
	);
	patched = true;
}

// Patch 2: wrap the lefthook call to strip any remaining ANSI via sed
const callLefthook = 'call_lefthook run "pre-commit" "$@"';
if (content.includes(callLefthook)) {
	const wrapper = `# Strip all ANSI escape codes so VS Code's Git dialog shows clean text.\noutput=$(mktemp)\n${callLefthook} >"$output" 2>&1\nexit_code=$?\nsed -E -e 's/\\x1b\\[[0-9;]*m//g' -e '/^[╭╰]/d' -e 's/^│ //; s/ │$//' "$output"\nrm -f "$output"\nexit $exit_code`;
	content = content.replace(callLefthook, wrapper);
	patched = true;
}

if (patched) {
	fs.writeFileSync(hookPath, content);
	console.log("patch-hook-no-color: patched .git/hooks/pre-commit for clean output");
} else {
	console.log("patch-hook-no-color: already patched, skipping");
}
