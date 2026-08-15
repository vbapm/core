import { readFile } from "fs-extra";
import { join } from "path";
import { execute, setup } from "./__helpers__/execute";
import { peerHost } from "./__fixtures__";

/**
 * Peer (VBA project-to-project) reference integration tests.
 *
 * Uses the `peer-host` fixture, which contains an embedded peer addin project
 * (`src/AddinPeer`). The host references the peer with a RELATIVE path, so these
 * tests are CI-safe (no absolute paths). The host source uses the peer, which is
 * required for the reference to persist across save (usage-gating).
 *
 * To run this file directly, use:
 * pnpm run test:e2e -- tests/peer-host.e2e.ts
 * or
 * pnpm run test:e2e:updateSnapshots -- tests/peer-host.e2e.ts
 *
 */
describe("peer references", () => {
	test("builds host with peer reference and round-trips the relative path", async () => {
		await setup(peerHost, "peer-host", async cwd => {
			// 1. Build the embedded peer addin first so the .xlam exists at
			//    the relative path the host references.
			console.log("Step 1: Building the embedded peer addin (src/AddinPeer)");
			const { stderr: peerErr } = await execute(join(cwd, "src/AddinPeer"), "build");
			expect(peerErr).not.toContain("Error");
			expect(await readFile(join(cwd, "src/AddinPeer/build/AddinPeer.xlam"))).toBeTruthy();

			// 2. Build the host. importGraph resolves the relative peer path to
			//    absolute and calls References.AddFromFile.
			console.log("Step 2: Building the host project");
			const { stderr: hostErr } = await execute(cwd, "build");
			expect(hostErr).not.toContain("Error");

			// 3. Extract. The reference should persist (host uses AddinPeer),
			//    and the stored path should round-trip back to relative.
			console.log("Step 3: Extracting the host (--target xlsm)");
			const { stderr: extractErr } = await execute(cwd, "extract --target xlsm");
			expect(extractErr).not.toContain("Error");

			console.log("Step 4: Reading the extracted manifest (vbaproject.toml)");
			const manifest = await readFile(join(cwd, "vbaproject.toml"), "utf8");
			expect(manifest).toContain(
				`AddinPeer = { peer = true, path = 'src/AddinPeer/build/AddinPeer.xlam' }`
			);

			console.log("Step 5: Snapshotting the extracted manifest (vbaproject.toml)");
			expect(manifest).toMatchSnapshot();
		});
	}, 180000);
});
