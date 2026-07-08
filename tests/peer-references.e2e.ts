import { readJson } from "fs-extra";
import { join } from "path";
import { tmp, execute, run } from "./__helpers__/execute";

/** Path to the built vbapm.xlam addin (resolved relative to project root). */
const VBAPM_ADDIN = join(__dirname, "../addins/build/vbapm.xlam");

/**
 * Phase 0: Exploration tests for peer (VBA project-to-project) references.
 *
 * These tests validate that `References.AddFromFile` works correctly by
 * calling the vbapm addin's VBA macros directly via `run()`, bypassing
 * the TypeScript pipeline (which hasn't been updated for peer refs yet).
 */
describe("peer references (Phase 0)", () => {
	test("scenario 0: verify Build.ImportGraph works with COM reference via run()", async () => {
		await tmp("peer-ref-0", async cwd => {
			// Create and build host project
			await execute(cwd, "new host --target xlsm --no-git");
			const hostDir = join(cwd, "host");
			await execute(hostDir, "add TestModule --type module");
			const { stderr: hostBuildErr } = await execute(hostDir, "build");
			expect(hostBuildErr).not.toContain("Error");
			const hostFile = join(hostDir, "build", "host.xlsm");

			// Call Build.ImportGraph with a known COM reference (VBIDE)
			const importPayload = JSON.stringify({
				file: hostFile,
				name: "host",
				src: [],
				references: [
					{
						name: "VBIDE",
						guid: "{0002E157-0000-0000-C000-000000000046}",
						major: 5,
						minor: 3
					}
				]
			});

			const result = await run("excel", VBAPM_ADDIN, "Build.ImportGraph", [importPayload]);
			if (!result.success) {
				console.error("ImportGraph errors:", JSON.stringify(result.errors));
			}
			expect(result.success).toBe(true);
			expect(result.errors).toHaveLength(0);
		});
	}, 120000);

	test("scenario 1: add peer reference to .xlam addin via AddFromFile", async () => {
		await tmp("peer-ref-1", async cwd => {
			// 1. Create and build peer addin project (.xlam)
			await execute(cwd, "new addin-peer --target xlam --no-git");
			const peerDir = join(cwd, "addin-peer");
			await execute(peerDir, "add Greeting --type module");
			const { stderr: peerBuildErr } = await execute(peerDir, "build");
			expect(peerBuildErr).not.toContain("Error");
			const peerFile = join(peerDir, "build", "AddinPeer.xlam");

			// 2. Create and build host project (.xlsm)
			await execute(cwd, "new host --target xlsm --no-git");
			const hostDir = join(cwd, "host");
			await execute(hostDir, "add CallPeer --type module");
			const { stderr: hostBuildErr } = await execute(hostDir, "build");
			expect(hostBuildErr).not.toContain("Error");
			const hostFile = join(hostDir, "build", "host.xlsm");

			// 3. Use run() to call Build.ImportGraph on the vbapm addin.
			//    The addin opens the host file internally via the JSON payload.
			const importPayload = JSON.stringify({
				file: hostFile,
				name: "host",
				src: [],
				references: [
					{
						name: "AddinPeer",
						peer: true,
						path: peerFile
					}
				]
			});

			const importResult = await run("excel", VBAPM_ADDIN, "Build.ImportGraph", [importPayload]);
			if (!importResult.success) {
				console.error("ImportGraph errors:", JSON.stringify(importResult.errors));
			}
			expect(importResult.success).toBe(true);
			expect(importResult.errors).toHaveLength(0);

			// 4. Export via the addin to verify the reference was written correctly
			const stagingDir = join(hostDir, "staging");
			const exportPayload = JSON.stringify({
				file: hostFile,
				staging: stagingDir
			});

			const exportResult = await run("excel", VBAPM_ADDIN, "Build.ExportTo", [exportPayload]);
			expect(exportResult.success).toBe(true);

			// 5. Read project.json and verify peer reference metadata
			const projectJson = await readJson(join(stagingDir, "project.json"));
			const peerRef = projectJson.references.find((r: any) => r.name === "AddinPeer");
			expect(peerRef).toBeDefined();
			expect(peerRef.guid).toBe("");
			expect(peerRef.major).toBe(0);
			expect(peerRef.minor).toBe(0);
			expect(peerRef.peer).toBe(true);
			expect(peerRef.path).toBe(peerFile);
		});
	}, 120000);
});
