export default {
	testEnvironment: "node",
	testTimeout: 120000,
	testMatch: ["**/tests/**/*.e2e.ts"],
	testPathIgnorePatterns: ["/node_modules/", "/lib/", "<rootDir>/worktrees/"],
	modulePathIgnorePatterns: ["<rootDir>/worktrees/"],
	transformIgnorePatterns: [
		"/node_modules/(?!.*(?:@decimalturn/toml-patch|env-paths|is-safe-filename)/)",
		"<rootDir>/lib/"
	],
	transform: {
		"^.+\\.[tj]sx?$": ["ts-jest", { tsconfig: "tests/tsconfig.json", diagnostics: false }]
	},
	moduleNameMapper: {
		"^vbapm$": "<rootDir>/lib/index.js",
		"^@timhall/dedent$": "<rootDir>/node_modules/@timhall/dedent/dist/dedent.js"
	},
	snapshotFormat: {
		escapeString: true,
		printBasicPrototype: true
	}
};
