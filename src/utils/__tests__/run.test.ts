import { jest } from "@jest/globals";

await jest.unstable_mockModule("../../env", () => ({
	env: {
		debug: () => () => undefined,
		isWindows: true
	}
}));

const { escape } = await import("../run");

test("should escape quotes and spaces", () => {
	expect(escape(`" '!?a1`)).toEqual(`^q '!?a1`);
});
