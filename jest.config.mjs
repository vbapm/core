export default {
	testEnvironment: "node",
	testPathIgnorePatterns: ["/node_modules/", "/lib/"],
	// @decimalturn/toml-patch v1+ is ESM-only. Jest (running in CJS mode via ts-jest) skips
	// transforming node_modules by default, which causes "Unexpected token export" errors.
	// This pattern excludes toml-patch from that skip list so ts-jest can downcompile it to CJS.
	transformIgnorePatterns: ["/node_modules/(?!@decimalturn/toml-patch/)"],
	// allowJs is required so ts-jest can process the plain .js ESM dist file of toml-patch.
	transform: {
		"^.+\\.[tj]sx?$": ["ts-jest", { tsconfig: { allowJs: true } }]
	},
	moduleNameMapper: {
		"^@timhall/dedent$": "<rootDir>/node_modules/@timhall/dedent/dist/dedent.js"
	},
	snapshotFormat: {
		escapeString: true,
		printBasicPrototype: true
	}
};
