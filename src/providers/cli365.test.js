import { describe, expect, it, vi } from "vitest";

import { createCli365Client, normalizeOutlookEvent, _internal } from "./cli365.js";
import { UserError } from "../errors.js";

describe("providers/cli365", () => {
	it("normalizes calendar event", () => {
		const ev = normalizeOutlookEvent({
			ItemId: "out-1",
			Subject: "Daily",
			Start: "2026-01-02T09:00:00-05:00",
			End: "2026-01-02T09:30:00-05:00",
			Location: { DisplayName: "Zoom" },
			Body: { Value: "desc" },
			RequiredAttendees: [{ Mailbox: { Address: "a@example.com", Name: "Alice" } }],
		});

		expect(ev.id).toBe("out-1");
		expect(ev.summary).toBe("Daily");
		expect(ev.start?.dateTime).toBe("2026-01-02T14:00:00.000Z");
		expect(ev.location).toBe("Zoom");
		expect(ev.attendeeNames).toEqual(["Alice"]);
	});

	it("listEvents builds args and parses Events", async () => {
		const runJson = vi.fn().mockResolvedValue({
			Events: [
				{ ItemId: "1", Subject: "A", Start: "2026-01-01T00:00:00Z", End: "2026-01-01T01:00:00Z" },
			],
		});
		const client = createCli365Client({ runJson, command: "go", commandArgs: ["run", "./cmd/cli-365"] });
		const events = await client.listEvents({
			start: "2026-01-01T00:00:00Z",
			end: "2026-01-07T00:00:00Z",
			limit: 10,
		});

		expect(events).toHaveLength(1);
		const call = runJson.mock.calls[0][0];
		expect(call.args).toContain("calendar");
		expect(call.args).toContain("list");
		expect(call.args).toContain("--json");
	});

	it("defaults to cli-365 on PATH and does not add CDP flags", async () => {
		const runJson = vi.fn().mockResolvedValue({ Events: [] });
		const client = createCli365Client({ runJson });
		await client.listEvents();
		const call = runJson.mock.calls[0][0];
		expect(call.command).toBe("cli-365");
		expect(call.cwd).toBeUndefined();
		expect(call.args).not.toContain("--cdp-port");
		expect(call.args).not.toContain("--ensure-cdp");
		expect(call.args).not.toContain("--ensure-cdp-timeout");
	});

	it("passes CDP flags only when explicitly configured", async () => {
		const runJson = vi.fn().mockResolvedValue({ Events: [] });
		const client = createCli365Client({
			runJson,
			configPath: "/tmp/cli-365.yaml",
			cdpPort: 36429,
			ensureCdp: true,
			ensureCdpTimeout: "2m",
		});

		await client.listEvents();
		const call = runJson.mock.calls[0][0];
		expect(call.args).toEqual(
			expect.arrayContaining([
				"--json",
				"--config",
				"/tmp/cli-365.yaml",
				"--cdp-port",
				"36429",
				"--ensure-cdp",
				"--ensure-cdp-timeout",
				"2m",
				"calendar",
				"list",
			])
		);
	});

	it("updateEvent recreates on known update failure", async () => {
		const runJson = vi
			.fn()
			.mockRejectedValueOnce(new UserError("ErrorSendMeetingInvitationsOrCancellationsRequired"))
			.mockResolvedValueOnce({
				ItemId: "new-1",
				Subject: "A",
				Start: "2026-01-01T00:00:00Z",
				End: "2026-01-01T01:00:00Z",
			});
		const runText = vi.fn().mockResolvedValue({ stdout: "Event deleted", stderr: "" });
		const client = createCli365Client({ runJson, runText });
		const res = await client.updateEvent({
			eventId: "old-1",
			patch: {
				summary: "A",
				start: { dateTime: "2026-01-01T00:00:00Z" },
				end: { dateTime: "2026-01-01T01:00:00Z" },
			},
		});

		expect(res.action).toBe("recreated");
		expect(res.previousId).toBe("old-1");
		expect(runText).toHaveBeenCalledTimes(1);
	});

	it("formatCliDate keeps all-day values date-only", () => {
		expect(_internal.formatCliDate({ dateTime: "2026-01-01T09:00:00Z" }, true)).toBe("2026-01-01");
		expect(_internal.formatCliDate({ date: "2026-01-02" }, true)).toBe("2026-01-02");
	});
});
