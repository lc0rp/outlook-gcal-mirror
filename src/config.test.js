import { describe, expect, it } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import { loadConfig, saveConfig } from "./config.js";
import { UserError } from "./errors.js";

describe("config", () => {
	it("saveConfig writes and loadConfig reads", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-config-"));
		const filePath = path.join(dir, "config.json");
		const cfg = {
			outlook: { includeCalendars: ["Team"] },
			google: { credentialsPath: "cred.json", tokenPath: "token.json", calendarName: "Outlook Mirror" },
			sync: { windowDays: 14, markCancelled: false },
		};

		await saveConfig(filePath, cfg);
		const loaded = await loadConfig(filePath);
		expect(loaded).toEqual(cfg);
	});

	it("loadConfig throws UserError when missing", async () => {
		await expect(loadConfig("/tmp/does-not-exist.json")).rejects.toBeInstanceOf(UserError);
	});
});
