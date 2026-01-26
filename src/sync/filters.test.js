import { describe, expect, it } from "vitest";
import { shouldSyncEvent } from "./filters.js";

describe("shouldSyncEvent", () => {
	const baseEvent = {
		sourceKey: "key-1",
		sourceCalendarName: "Work Calendar",
		sourceOwnerEmail: "alice@example.com",
	};

	it("returns true with no rules", () => {
		expect(shouldSyncEvent(baseEvent, {})).toBe(true);
	});

	it("skipCalendars excludes matching calendars", () => {
		expect(shouldSyncEvent(baseEvent, { skipCalendars: ["work"] })).toBe(false);
		expect(shouldSyncEvent(baseEvent, { skipCalendars: ["personal"] })).toBe(true);
	});

	it("includeCalendars requires match", () => {
		expect(shouldSyncEvent(baseEvent, { includeCalendars: ["Work"] })).toBe(true);
		expect(shouldSyncEvent(baseEvent, { includeCalendars: ["Personal"] })).toBe(false);
	});

	it("skipOwnerEmails excludes matching emails (case-insensitive)", () => {
		expect(shouldSyncEvent(baseEvent, { skipOwnerEmails: ["ALICE@example.com"] })).toBe(false);
		expect(shouldSyncEvent(baseEvent, { skipOwnerEmails: ["bob@example.com"] })).toBe(true);
	});

	it("includeOwnerEmails requires match (case-insensitive)", () => {
		expect(shouldSyncEvent(baseEvent, { includeOwnerEmails: ["Alice@Example.com"] })).toBe(true);
		expect(shouldSyncEvent(baseEvent, { includeOwnerEmails: ["bob@example.com"] })).toBe(false);
	});

	it("handles missing sourceCalendarName gracefully", () => {
		const ev = { ...baseEvent, sourceCalendarName: undefined };
		expect(shouldSyncEvent(ev, { skipCalendars: ["Work"] })).toBe(true);
	});
});
