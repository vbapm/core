export default {
	testEnvironment: "node",
	testPathIgnorePatterns: ["/node_modules/", "/lib/", "<rootDir>/worktrees/"],
	modulePathIgnorePatterns: ["<rootDir>/worktrees/"],
	// ESM-only packages. Jest (CJS mode via ts-jest) skips transforming
	// node_modules by default. The ignore pattern excludes everything EXCEPT
	// these packages so ts-jest can downcompile them to CJS.
	// Pattern accounts for both npm flat and pnpm .pnpm/ directory layouts.
	transformIgnorePatterns: ["/node_modules/(?!.*(?:@decimalturn/toml-patch|env-paths|is-safe-filename|jschardet)/)"],	
	// allowJs is required so ts-jest can process the plain .js ESM dist file of toml-patch.
	transform: {
		"^.+\\.[tj]sx?$": ["ts-jest", { tsconfig: { allowJs: true } }]
	},
	moduleNameMapper: {
		"^@timhall/dedent$": "<rootDir>/node_modules/@timhall/dedent/dist/dedent.js",
		"^jschardet$": "<rootDir>/node_modules/jschardet/build/index.js"
	},
	snapshotFormat: {
		escapeString: true,
		printBasicPrototype: true
	}
};
