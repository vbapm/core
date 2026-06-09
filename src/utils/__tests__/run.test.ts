jest.mock("../../env", () => ({
	env: {
		debug: () => () => undefined,
		isWindows: true
	}
}));

import { escape } from "../run";

test("should escape quotes and spaces", () => {
	expect(escape(`" '!?a1`)).toEqual(`^q '!?a1`);
});
