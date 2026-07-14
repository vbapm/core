declare module "walk-sync" {
	export interface WalkOptions {
		directories?: boolean;
		globs?: string[];
	}
	export default function walkSync(dir: string, options?: WalkOptions): string[];
}
