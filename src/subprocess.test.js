import { describe, expect, it, vi } from "vitest";

import { runCommand, runJsonCommand, _internal } from "./subprocess.js";
import { UserError } from "./errors.js";

describe("subprocess", () => {
	it("runCommand returns stdout and stderr", async () => {
		const runner = vi.fn().mockResolvedValue({ stdout: "ok", stderr: "warn" });
		const result = await runCommand({
			command: "echo",
			args: ["ok"],
			runner,
		});
		expect(result).toEqual({ stdout: "ok", stderr: "warn" });
	});

	it("runCommand wraps failures in UserError", async () => {
		const runner = vi.fn().mockRejectedValue({
			message: "exit 1",
			stderr: "boom",
			stdout: "",
		});
		await expect(
			runCommand({ command: "bad", args: ["--x"], runner })
		).rejects.toBeInstanceOf(UserError);
	});

	it("runJsonCommand parses direct JSON", async () => {
		const runner = vi.fn().mockResolvedValue({ stdout: '{"ok":true}', stderr: "" });
		const json = await runJsonCommand({ command: "x", runner, label: "test" });
		expect(json).toEqual({ ok: true });
	});

	it("runJsonCommand parses trailing JSON line", async () => {
		const runner = vi.fn().mockResolvedValue({
			stdout: "log line\n{\"ok\":true}",
			stderr: "",
		});
		const json = await runJsonCommand({ command: "x", runner, label: "test" });
		expect(json).toEqual({ ok: true });
	});

	it("runJsonCommand throws on non-JSON output", async () => {
		const runner = vi.fn().mockResolvedValue({ stdout: "hello", stderr: "" });
		await expect(
			runJsonCommand({ command: "x", runner, label: "test" })
		).rejects.toBeInstanceOf(UserError);
	});

	it("parseJsonOutput handles arrays", () => {
		expect(_internal.parseJsonOutput("[1,2,3]")).toEqual([1, 2, 3]);
	});
});
