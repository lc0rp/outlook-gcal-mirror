import { describe, expect, it } from "vitest";
import { importOptional } from "./imports.js";
import { UserError } from "../errors.js";

describe("importOptional", () => {
	it("imports existing module", async () => {
		const mod = await importOptional("node:path");
		expect(mod.join).toBeDefined();
	});

	it("throws UserError when all modules missing", async () => {
		await expect(importOptional(["nonexistent-pkg-1", "nonexistent-pkg-2"])).rejects.toBeInstanceOf(UserError);
	});

	it("single string name also works", async () => {
		const mod = await importOptional("node:fs");
		expect(mod.promises).toBeDefined();
	});
});
