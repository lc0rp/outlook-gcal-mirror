import { describe, expect, test } from "vitest";

import { extractOutlookEventIdsFromJson, extractOutlookEventsFromJson } from "./extract.js";

describe("extractOutlookEventsFromJson", () => {
	test("extracts a basic event shape", () => {
		const json = {
			Subject: "Weekly 1:1",
			Start: "2026-01-01T10:00:00Z",
			End: "2026-01-01T10:30:00Z",
			Attendees: [{ EmailAddress: { Name: "Alice" } }, { EmailAddress: { Name: "Bob" } }],
			Organizer: { EmailAddress: { Address: "organizer@example.com" } },
			Id: "AAMk-test-id",
		};

		const events = extractOutlookEventsFromJson(json);
		expect(events).toHaveLength(1);
		expect(events[0].subject).toBe("Weekly 1:1");
		expect(events[0].sourceId).toBe("AAMk-test-id");
		expect(events[0].organizerEmail).toBe("organizer@example.com");
		expect(events[0].attendeeNames).toEqual(["Alice", "Bob"]);
	});

	test("finds events nested in arrays/objects and de-dupes", () => {
		const event = {
			subject: "Planning",
			start: { DateTime: "2026-02-02T09:00:00", TimeZone: "America/New_York" },
			end: { DateTime: "2026-02-02T10:00:00", TimeZone: "America/New_York" },
			attendees: [{ name: "Carol" }, { attendee: { displayName: "Dave" } }],
			itemId: "item-123",
		};

		const json = { a: [event, { deep: { event } }] };
		const events = extractOutlookEventsFromJson(json);
		expect(events).toHaveLength(1);
		expect(events[0].subject).toBe("Planning");
		expect(events[0].sourceId).toBe("item-123");
		expect(events[0].attendeeNames).toEqual(["Carol", "Dave"]);
	});

	test("captures ItemId.Id values", () => {
		const json = {
			Subject: "Demo",
			Start: "2026-03-01T09:00:00Z",
			End: "2026-03-01T10:00:00Z",
			ItemId: { Id: "AAk-test-item" },
		};

		const events = extractOutlookEventsFromJson(json);
		expect(events).toHaveLength(1);
		expect(events[0].sourceId).toBe("AAk-test-item");

		const ids = extractOutlookEventIdsFromJson(json);
		expect(ids).toEqual(["AAk-test-item"]);
	});
});
