export default {
	testEnvironment: "node",
	testTimeout: 120000,
	globalSetup: "<rootDir>/tests/global-setup.js",
	globalTeardown: "<rootDir>/tests/global-teardown.js",
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
		"^@timhall/dedent$": "<rootDir>/node_modules/@timhall/dedent/dist/dedent.js",
		"^open$": "<rootDir>/tests/__helpers__/open-stub.ts"
	},
	snapshotFormat: {
		escapeString: true,
		printBasicPrototype: true
	}
};
