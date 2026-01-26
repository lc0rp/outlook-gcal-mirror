import { describe, expect, it } from "vitest";
import { buildGoogleDescription, cancelledSummary } from "./format.js";

describe("format", () => {
	it("buildGoogleDescription formats event details", () => {
		const ev = {
			subject: "Meeting",
			sourceKey: "key-123",
			sourceId: "id-456",
			attendeeNames: ["Alice", "Bob"],
			organizerEmail: "organizer@example.com",
			sourceCalendarName: "Work",
			sourceOwnerEmail: "owner@example.com",
		};

		const desc = buildGoogleDescription(ev);
		expect(desc).toContain("Mirrored from Outlook");
		expect(desc).toContain("Attendees:");
		expect(desc).toContain("- Alice");
		expect(desc).toContain("- Bob");
		expect(desc).toContain("Organizer: organizer@example.com");
		expect(desc).toContain("Source calendar: Work (owner@example.com)");
		expect(desc).toContain("Source key: key-123");
		expect(desc).toContain("Source id: id-456");
	});

	it("buildGoogleDescription handles minimal event", () => {
		const ev = { subject: "Quick", sourceKey: "k", attendeeNames: [] };
		const desc = buildGoogleDescription(ev);
		expect(desc).toContain("Mirrored from Outlook");
		expect(desc).not.toContain("Attendees:");
	});

	it("cancelledSummary adds prefix", () => {
		expect(cancelledSummary("Meeting")).toBe("CANCELLED: Meeting");
		expect(cancelledSummary("CANCELLED: Meeting")).toBe("CANCELLED: Meeting");
		expect(cancelledSummary("cancelled: Foo")).toBe("cancelled: Foo");
	});
});
