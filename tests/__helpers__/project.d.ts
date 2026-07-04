interface EnvironmentOptions {
	silent?: boolean;
}
export declare function setupEnvironment(cwd: string, options?: EnvironmentOptions): void;
export declare function setup(
	cwd: string,
	options?: EnvironmentOptions
): Promise<{
	project: import("../../src/project").Project;
	dependencies: import("../../src/manifest").Manifest[];
}>;
export declare function setupWorkspace(cwd: string): Promise<{
	manifest: import("../../src/manifest").Manifest;
	workspace: import("../../src/professional/workspace").Workspace;
	config: import("../../src/config").Config;
}>;
export declare function reset(): void;
