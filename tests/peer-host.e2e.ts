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
 */
describe("peer references", () => {
	test("builds host with peer reference and round-trips the relative path", async () => {
		await setup(peerHost, "peer-host", async cwd => {
			// 1. Build the embedded peer addin first so the .xlam exists at
			//    the relative path the host references.
			const { stderr: peerErr } = await execute(join(cwd, "src/AddinPeer"), "build");
			expect(peerErr).not.toContain("Error");
			expect(await readFile(join(cwd, "src/AddinPeer/build/AddinPeer.xlam"))).toBeTruthy();

			// 2. Build the host. importGraph resolves the relative peer path to
			//    absolute and calls References.AddFromFile.
			const { stderr: hostErr } = await execute(cwd, "build");
			expect(hostErr).not.toContain("Error");

			// 3. Extract. The reference should persist (host uses AddinPeer),
			//    and the stored path should round-trip back to relative.
			const { stderr: extractErr } = await execute(cwd, "extract --target xlsm");
			expect(extractErr).not.toContain("Error");

			const manifest = await readFile(join(cwd, "vbaproject.toml"), "utf8");
			expect(manifest).toContain(
				`AddinPeer = { peer = true, path = 'src/AddinPeer/build/AddinPeer.xlam' }`
			);
		});
	}, 180000);
});
