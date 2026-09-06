declare module "vbapm" {
	interface RunResult {
		success: boolean;
		messages: string[];
		warnings: string[];
		errors: string[];
		stdout?: string;
		stderr?: string;
	}

	export function run(
		application: string,
		file: string,
		macro: string,
		args: string[]
	): Promise<RunResult>;

	export function closePowerShellSession(): Promise<void>;
	export function resetPersistentPowerShellSession(): Promise<void>;
}
