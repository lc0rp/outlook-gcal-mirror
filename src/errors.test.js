import { describe, expect, it } from "vitest";
import { UserError } from "./errors.js";

describe("UserError", () => {
	it("sets name and message", () => {
		const err = new UserError("nope");
		expect(err.message).toBe("nope");
		expect(err.name).toBe("UserError");
	});
});
