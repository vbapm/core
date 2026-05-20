export default {
	testEnvironment: "node",
	testTimeout: 120000,
	testMatch: ["**/tests/**/*.e2e.ts"],
	testPathIgnorePatterns: ["/node_modules/", "/lib/", "/worktrees/"],
	modulePathIgnorePatterns: ["<rootDir>/worktrees/"],
	transformIgnorePatterns: ["/node_modules/(?!@decimalturn/toml-patch/|env-paths/|is-safe-filename/)"],
	transform: {
		"^.+\\.(gitignore|gitattributes|editorconfig)$": "<rootDir>/tests/jest-text-template-transform.cjs",
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
