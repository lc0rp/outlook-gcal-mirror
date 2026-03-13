import { describe, expect, it } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import { buildProgram } from "./cli.js";

function getCommand(program, name) {
	return program.commands.find((command) => command.name() === name);
}

describe("cli", () => {
	it("removes legacy direct-outlook commands", () => {
		const program = buildProgram();
		const commandNames = program.commands.map((command) => command.name());

		expect(commandNames).not.toContain("keepalive");
		expect(commandNames).not.toContain("discover-owa");
		expect(commandNames).not.toContain("discover-owa-log");
		expect(commandNames).not.toContain("capture-owa");
		expect(commandNames).not.toContain("fetch-owa");
	});

	it("setup stays cli-365 centric", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-cli-"));
		const configPath = path.join(dir, "config.json");
		const program = buildProgram();
		const setup = getCommand(program, "setup");

		expect(setup?.options.map((option) => option.long)).not.toContain("--cdp-port");
		expect(setup?.options.map((option) => option.long)).not.toContain("--engine");
		expect(setup?.options.map((option) => option.long)).not.toContain("--target-url");

		await program.parseAsync(
			[
				"--config",
				configPath,
				"setup",
				"--google-credentials",
				"/tmp/credentials.json",
				"--calendar",
				"Outlook Mirror",
			],
			{ from: "user" }
		);

		const saved = JSON.parse(await fs.readFile(configPath, "utf8"));
		expect(saved).toEqual({
			outlook: {},
			google: {
				credentialsPath: "/tmp/credentials.json",
				tokenPath: expect.stringContaining("google-token.json"),
				calendarName: "Outlook Mirror",
			},
			sync: {
				windowDays: 14,
				markCancelled: false,
			},
		});
	});
});
