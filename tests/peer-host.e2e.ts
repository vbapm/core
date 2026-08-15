import { readFile } from "fs-extra";
import { join } from "path";
import { execute, setup } from "./__helpers__/execute";
import { peerHost } from "./__fixtures__";

/**
 * Peer (VBA project-to-project) reference integration tests.
 *
 * Uses the `peer-host` fixture, which contains the host project and a PRE-BUILT
 * peer addin at `src/AddinPeer/build/AddinPeer.xlam`. The addin is built offline
 * and committed to the fixture so this test doesn't need to spin up Excel to
 * build it (one fewer Excel instance = fewer file-lock/cleanup races on CI).
 * The host references the peer with a RELATIVE path, so these tests are CI-safe
 * (no absolute paths). The host source uses the peer, which is required for the
 * reference to persist across save (usage-gating).
 *
 * To rebuild the committed addin after changing `src/AddinPeer`, run `vba build`
 * inside `src/AddinPeer` and commit the regenerated
 * `src/AddinPeer/build/AddinPeer.xlam`.
 *
 * To run this file directly, use:
 * pnpm run test:e2e -- tests/peer-host.e2e.ts
 * or
 * pnpm run test:e2e:updateSnapshots -- tests/peer-host.e2e.ts
 */
describe("peer references", () => {
	test("builds host with peer reference and round-trips the relative path", async () => {
		await setup(peerHost, "peer-host", async cwd => {
			// The peer addin is pre-built and committed in the fixture, so it
			// already exists at the relative path the host references.
			expect(await readFile(join(cwd, "src/AddinPeer/build/AddinPeer.xlam"))).toBeTruthy();

			// Build the host. importGraph resolves the relative peer path to
			// absolute and calls References.AddFromFile.
			const { stderr: hostErr } = await execute(cwd, "build");
			expect(hostErr).not.toContain("Error");

			// Extract. The reference should persist (host uses AddinPeer),
			// and the stored path should round-trip back to relative.
			const { stderr: extractErr } = await execute(cwd, "extract --target xlsm");
			expect(extractErr).not.toContain("Error");

			const manifest = await readFile(join(cwd, "vbaproject.toml"), "utf8");
			expect(manifest).toContain(
				`AddinPeer = { peer = true, path = 'src/AddinPeer/build/AddinPeer.xlam' }`
			);

			expect(manifest).toMatchSnapshot();
		});
	}, 180000);
});
