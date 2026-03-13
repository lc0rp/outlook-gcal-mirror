import { describe, expect, it } from "vitest";
import { promises as fs } from "node:fs";
import { execFile } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function runCli(args) {
	return await execFileAsync(process.execPath, ["src/cli.js", ...args], {
		cwd: repoRoot,
		env: { ...process.env, FORCE_COLOR: "0" },
	});
}

describe("cli", () => {
	it("setup writes cli-365-first config only", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-cli-"));
		const configPath = path.join(dir, "config.json");

		await runCli([
			"--config",
			configPath,
			"setup",
			"--google-credentials",
			"/tmp/creds.json",
			"--calendar",
			"Mirror",
		]);

		const cfg = JSON.parse(await fs.readFile(configPath, "utf8"));

		expect(cfg).toEqual({
			google: {
				credentialsPath: "/tmp/creds.json",
				tokenPath: expect.stringContaining("google-token.json"),
				calendarName: "Mirror",
			},
			sync: {
				windowDays: 14,
				markCancelled: false,
			},
		});
	});

	it("help omits retired OWA commands", async () => {
		const { stdout } = await runCli(["--help"]);

		expect(stdout).not.toContain("keepalive");
		expect(stdout).not.toContain("discover-owa");
		expect(stdout).not.toContain("discover-owa-log");
		expect(stdout).not.toContain("capture-owa");
		expect(stdout).not.toContain("fetch-owa");
	});

	it("README quickstart stays cli-365-first", async () => {
		const readme = await fs.readFile(path.join(repoRoot, "README.md"), "utf8");
		const docs = readme.split("## TODO:")[0];

		expect(docs).not.toContain("keepalive");
		expect(docs).not.toContain("discover-owa");
		expect(docs).not.toContain("discover-owa-log");
		expect(docs).not.toContain("capture-owa");
		expect(docs).not.toContain("fetch-owa");
		expect(docs).not.toMatch(/(^|[\s`])--engine(?=$|[\s`])/m);
		expect(docs).not.toMatch(/(^|[\s`])--target-url(?=$|[\s`])/m);
		expect(docs).not.toMatch(/(^|[\s`])--cdp-port(?=$|[\s`])/m);
	});

	it("spec and package stay free of retired OWA runtime paths", async () => {
		const spec = await fs.readFile(path.join(repoRoot, "SPEC.md"), "utf8");
		const pkg = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));

		expect(spec).not.toContain("OWA/CDP");
		expect(spec).not.toContain("owa-tooling");
		expect(spec).not.toContain("Playwright");
		expect(spec).not.toContain("Puppeteer");

		expect(pkg.dependencies).not.toHaveProperty("playwright");
		expect(pkg.dependencies).not.toHaveProperty("puppeteer");
	});
});
