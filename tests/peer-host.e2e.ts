/**
 * End-to-end test for VBA project-to-project ("peer") references.
 *
 * Uses the `peer-host` fixture, which contains the host project and a PRE-BUILT
 * peer addin at `src/AddinPeer/build/AddinPeer.xlam`. The addin is built offline
 * and committed to the fixture so this test doesn't need to spin up Excel to
 * build it (one fewer Excel instance = fewer file-lock/cleanup races on CI).
 * The host references the peer with a RELATIVE path, so the test is CI-safe
 * (no absolute paths). The host source uses the peer, which is required for the
 * reference to persist across save (usage-gating).
 *
 * This file forces background mode (fresh hidden instances) so it never attaches
 * to a lingering Excel instance from another suite — the original cause of the
 * intermittent EBUSY ("file locked") failures on the peer addin.
 *
 * To rebuild the committed addin after changing `src/AddinPeer`, run `vba build`
 * inside `src/AddinPeer` and commit the regenerated
 * `src/AddinPeer/build/AddinPeer.xlam`.
 */

import { readFile } from "fs-extra";
import { join } from "path";
import { peerHost } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";

// Always run this file against fresh hidden Excel instances: in visible mode
// `run.ps1` attaches to any running Excel and never quits it, which left the
// peer addin locked (`~$AddinPeer.xlam`) when a lingering instance was reused.
process.env.VBA_BACKGROUND_BUILD = "1";

jest.setTimeout(180000);

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
				`AddinPeer = { peer = true, path = "src/AddinPeer/build/AddinPeer.xlam" }`
			);

			expect(manifest).toMatchSnapshot();
		});
	}, 180000);
});
