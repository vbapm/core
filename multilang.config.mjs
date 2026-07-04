/**
 * Jest configuration for multilingual encoding tests.
 *
 * These tests verify that VBA source files in different Windows ANSI
 * codepages survive the full roundtrip. They require a Windows machine
 * with the matching system locale and the E2E_ML or CI environment
 * variable set.
 *
 * Run:
 *   $env:E2E_ML=1; npx jest --config multilang.config.mjs --runInBand
 */

export default {
	testEnvironment: "node",
	testTimeout: 300000,
	testMatch: ["**/tests/**/*.multilang.ts"],
	testPathIgnorePatterns: ["/node_modules/", "/lib/", "<rootDir>/worktrees/"],
	modulePathIgnorePatterns: ["<rootDir>/worktrees/"],
	transformIgnorePatterns: [
		"/node_modules/(?!.*(?:@decimalturn/toml-patch|env-paths|is-safe-filename)/)"
	],
	transform: {
		"^.+\\.[tj]sx?$": ["ts-jest", { tsconfig: "tests/tsconfig.json", diagnostics: false }]
	},
	moduleNameMapper: {
		"^vbapm$": "<rootDir>/src/index.ts",
		"^@timhall/dedent$": "<rootDir>/node_modules/@timhall/dedent/dist/dedent.js"
	},
	snapshotFormat: {
		escapeString: true,
		printBasicPrototype: true
	}
};
