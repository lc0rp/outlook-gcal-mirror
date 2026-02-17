import { describe, expect, it, vi } from "vitest";

import {
	createGogClient,
	normalizeGoogleEvent,
	GOOGLE_LINK_OUTLOOK_ID,
	GOOGLE_LINK_VERSION,
	GOOGLE_LINK_VERSION_VALUE,
	_internal,
} from "./gog.js";

describe("providers/gog", () => {
	it("normalizes google event", () => {
		const ev = normalizeGoogleEvent({
			id: "g-1",
			summary: "Standup",
			start: { dateTime: "2026-01-01T10:00:00-05:00" },
			end: { dateTime: "2026-01-01T10:30:00-05:00" },
			attendees: [{ email: "a@example.com", displayName: "Alice" }],
			extendedProperties: { private: { [GOOGLE_LINK_OUTLOOK_ID]: "o-1" } },
		});

		expect(ev.id).toBe("g-1");
		expect(ev.summary).toBe("Standup");
		expect(ev.allDay).toBe(false);
		expect(ev.attendeeNames).toEqual(["Alice"]);
		expect(ev.privateProps[GOOGLE_LINK_OUTLOOK_ID]).toBe("o-1");
	});

	it("resolveCalendarId matches by name", async () => {
		const runJson = vi
			.fn()
			.mockResolvedValueOnce({ calendars: [{ id: "cal-1", summary: "Outlook Mirror" }] });
		const client = createGogClient({ runJson });
		const id = await client.resolveCalendarId("Outlook Mirror");
		expect(id).toBe("cal-1");
	});

	it("listEvents requests calendar events", async () => {
		const runJson = vi
			.fn()
			.mockResolvedValueOnce({ events: [{ id: "g-1", summary: "A", start: { date: "2026-01-01" }, end: { date: "2026-01-02" } }] });
		const client = createGogClient({ runJson, account: "me@example.com" });
		const events = await client.listEvents({
			calendarId: "primary",
			from: "2026-01-01T00:00:00Z",
			to: "2026-01-07T00:00:00Z",
		});
		expect(events).toHaveLength(1);
		const args = runJson.mock.calls[0][0].args;
		expect(args).toContain("--account");
		expect(args).toContain("calendar");
		expect(args).toContain("events");
	});

	it("createEvent includes send-updates and private props", async () => {
		const runJson = vi
			.fn()
			.mockResolvedValueOnce({ event: { id: "g-2", summary: "A", start: { date: "2026-01-01" }, end: { date: "2026-01-02" } } });
		const client = createGogClient({ runJson });
		await client.createEvent({
			calendarId: "primary",
			event: {
				summary: "A",
				start: { date: "2026-01-01" },
				end: { date: "2026-01-02" },
				allDay: true,
				privateProps: {
					[GOOGLE_LINK_OUTLOOK_ID]: "out-1",
					[GOOGLE_LINK_VERSION]: GOOGLE_LINK_VERSION_VALUE,
				},
			},
		});
		const args = runJson.mock.calls[0][0].args;
		expect(args).toContain("--send-updates");
		expect(args).toContain("none");
		expect(args).toContain(`${GOOGLE_LINK_OUTLOOK_ID}=out-1`);
	});

	it("formatGoogleTime returns date for all-day", () => {
		expect(_internal.formatGoogleTime({ dateTime: "2026-01-01T09:00:00Z" }, true)).toBe("2026-01-01");
	});
});
