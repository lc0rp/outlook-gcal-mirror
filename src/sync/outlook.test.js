import { describe, expect, it } from "vitest";

import { cli365EventToNormalized } from "./outlook.js";

describe("sync/outlook", () => {
	it("maps cli-365 event to normalized event", () => {
		const normalized = cli365EventToNormalized({
			id: "out-1",
			summary: "Standup",
			start: { dateTime: "2026-01-01T10:00:00Z" },
			end: { dateTime: "2026-01-01T10:30:00Z" },
			attendeeNames: ["Alice", "Bob"],
			raw: {
				Organizer: { Mailbox: { Address: "organizer@example.com" } },
				ParentFolder: { DisplayName: "Calendar" },
				Owner: { EmailAddress: { Address: "owner@example.com" } },
			},
		});

		expect(normalized.sourceKey).toBe("out-1");
		expect(normalized.sourceId).toBe("out-1");
		expect(normalized.subject).toBe("Standup");
		expect(normalized.organizerEmail).toBe("organizer@example.com");
		expect(normalized.sourceCalendarName).toBe("Calendar");
		expect(normalized.sourceOwnerEmail).toBe("owner@example.com");
		expect(normalized.attendeeNames).toEqual(["Alice", "Bob"]);
	});

	it("builds hashed sourceKey when id missing", () => {
		const normalized = cli365EventToNormalized({
			summary: "Untitled",
			start: { date: "2026-01-01" },
			end: { date: "2026-01-02" },
			attendeeNames: [],
			raw: {},
		});

		expect(normalized.sourceId).toBeNull();
		expect(normalized.sourceKey).toMatch(/^[a-f0-9]{40}$/);
		expect(normalized.start).toEqual({ date: "2026-01-01" });
		expect(normalized.end).toEqual({ date: "2026-01-02" });
	});
});
